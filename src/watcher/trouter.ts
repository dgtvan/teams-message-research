import type { Page } from "playwright";
import { htmlToText } from "./chatsvc.js";

/**
 * Teams delivers message events to the web client in real time over
 * "Trouter" — Microsoft's WebSocket-based notification transport — as a
 * relayed HTTP-shaped POST frame, one connection carrying events for
 * *every* conversation on the account, not per-conversation. This is what
 * makes near-real-time detection across many conversations possible on a
 * single tab — polling a single tab can't get there (see config.pool's
 * comment), but a single passive listener on this one connection covers
 * every tracked conversation at once.
 *
 * Frame shape (Socket.IO-flavored, not documented anywhere public):
 *   3:::{"id":...,"method":"POST","url":"/v4/f/<id>/messaging",...,"body":"<JSON string>"}
 * `body` is itself a JSON-encoded string; decoding it twice gives the
 * actual notification: { resourceType: "...", resource: {...} }.
 *
 * Three resourceType values confirmed live, all carrying the same shape:
 * - "NewMessage": a new message. resource.content is the HTML content.
 * - "MessageUpdate" with resource.properties.edittime set: an edit —
 *   same resource.id as the original message, resource.content is the
 *   new content.
 * - "MessageUpdate" with resource.properties.deletetime set (no
 *   edittime): a deletion — same resource.id, resource.content is "".
 * "ConversationUpdate" also fires alongside NewMessage/MessageUpdate,
 * carrying the same message redundantly as a left-rail preview refresh —
 * ignored, since the two above already cover it.
 */

export type PushedMessageKind = "new" | "edit" | "delete";

export interface PushedMessage {
  kind: PushedMessageKind;
  conversationId: string;
  messageId: string;
  content: string;
  sender: string | null;
  timestamp: string;
  /**
   * True when the RAW resource.content (before htmlToText strips every
   * tag) contains an `<img>` tag — confirmed live (teams.ts) that both a
   * directly-pasted image and a real uploaded image render as an actual
   * `<img>` in Teams' own DOM, and resource.content is that same
   * underlying message HTML, so the tag survives into the push payload
   * too. Lets handlePushedMessage (watcher/index.ts) recognize a message
   * that mixes real text with an image — content is non-empty (so the
   * existing empty-content pending check doesn't fire) but still
   * push-incomplete, since push can never carry the actual image bytes.
   * Without this, a text+image message just sat on the normal
   * interval-based rotation instead of being prioritized, which on a busy
   * account can leave it unresolved far longer than the "eventually"
   * self-healing is supposed to mean (confirmed live: a "123"/image,
   * "456"/image, "789"/image message in "48:notes" sent via three
   * paragraphs stayed at its original push-only, image-less revision for
   * 8+ minutes with zero DOM revisits, well past the 5min+jitter backup
   * interval that should have caught it on the normal rotation).
   */
  hasInlineImage: boolean;
}

function parseTrouterFrame(text: string): PushedMessage | null {
  const jsonStart = text.indexOf("{");
  if (jsonStart === -1) return null;

  let outer: unknown;
  try {
    outer = JSON.parse(text.slice(jsonStart));
  } catch {
    return null;
  }
  if (typeof outer !== "object" || outer === null) return null;
  const outerObj = outer as { url?: unknown; body?: unknown };
  if (typeof outerObj.url !== "string" || !outerObj.url.endsWith("/messaging")) return null;
  if (typeof outerObj.body !== "string") return null;

  let notification: unknown;
  try {
    notification = JSON.parse(outerObj.body);
  } catch {
    return null;
  }
  if (typeof notification !== "object" || notification === null) return null;
  const n = notification as { resourceType?: unknown; resource?: unknown };
  if (typeof n.resource !== "object" || n.resource === null) return null;
  const r = n.resource as Record<string, unknown>;

  let kind: PushedMessageKind;
  if (n.resourceType === "NewMessage") {
    kind = "new";
  } else if (n.resourceType === "MessageUpdate") {
    const props = (r.properties && typeof r.properties === "object" ? r.properties : {}) as Record<
      string,
      unknown
    >;
    kind = props.deletetime ? "delete" : "edit";
  } else {
    return null; // e.g. ConversationUpdate — redundant with NewMessage/MessageUpdate, ignored
  }

  const conversationId = typeof r.to === "string" ? r.to : null;
  const messageId = typeof r.id === "string" ? r.id : null;
  if (!conversationId || !messageId) return null;

  const rawContent = typeof r.content === "string" ? r.content : "";
  const content = htmlToText(rawContent);
  const sender = typeof r.imdisplayname === "string" ? r.imdisplayname : null;
  const hasInlineImage = /<img\b/i.test(rawContent);

  // Confirmed live (two phantom rows in the "Daily Scrum" conversation,
  // message_ids 1787818820426 / 1787818877035): Teams also pushes
  // non-message events — almost certainly reactions (👍/❤️/etc. on an
  // existing message), based on the timing and the fact both rows appeared
  // seconds apart with nothing typed — over this same connection tagged
  // resourceType "NewMessage", carrying a fresh, real `to`/`id` (so they
  // pass the check above) but an empty resource.content and
  // resource.imdisplayname explicitly set to "" (not absent — sender comes
  // back "", not null). A real chat message — even an image/attachment-only
  // one with no text — always carries the actual sender's display name, so
  // "kind is new, but content AND sender are both empty" is a safe signal
  // this isn't really a new message at all, without needing to know
  // reactions' real field names (attempting to capture a live example to
  // confirm those was blocked before it could run — interacting with the
  // real account, even in the safe self-chat, tripped the automation
  // classifier — so this fix is deliberately just "don't record it as a
  // message" rather than also surfacing what was reacted with/by whom;
  // see the TODO below).
  //
  // Scoped to kind === "new" only: a genuine deletion legitimately has
  // empty content (see the module comment) and must still be recorded, and
  // an edit/delete for a message we don't already own is dropped earlier
  // regardless (handlePushedMessage in watcher/index.ts).
  //
  // TODO: if a real reaction payload is ever captured (e.g. by logging
  // JSON.stringify(notification) here for a while and watching for one),
  // extract who reacted + with what and expose it as an annotation on the
  // original message instead of just discarding it.
  if (kind === "new" && content === "" && !sender) return null;

  return {
    kind,
    conversationId,
    messageId,
    content,
    sender,
    timestamp:
      typeof r.originalarrivaltime === "string" ? r.originalarrivaltime : new Date().toISOString(),
    hasInlineImage,
  };
}

/**
 * Attaches a passive listener to whichever Trouter WebSocket(s) the page
 * opens, calling onMessage for every real-time new/edit/delete push it
 * sees. Safe to call once per page — Playwright's `page.on("websocket",
 * ...)` catches connections made at any point, including ones already
 * open by the time this runs, since Teams reconnects Trouter periodically
 * anyway.
 */
export function attachMessagePushListener(page: Page, onMessage: (msg: PushedMessage) => void): void {
  page.on("websocket", (ws) => {
    if (!ws.url().includes("trouter")) return;
    ws.on("framereceived", (frame) => {
      if (typeof frame.payload !== "string") return;
      const parsed = parseTrouterFrame(frame.payload);
      if (parsed) onMessage(parsed);
    });
  });
}
