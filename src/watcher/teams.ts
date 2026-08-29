import type { Page } from "playwright";
import crypto from "node:crypto";
import { selectors } from "./selectors.js";
import type { ObservedMessage } from "../db/index.js";

export interface DiscoveredConversation {
  id: string;
  title: string;
  kind: "chat" | "channel";
}

export interface RawAttachment {
  name: string;
  url: string;
}

export interface RawMessage {
  messageId: string;
  sender: string | null;
  content: string;
  editedFlag: boolean;
  isDeleted: boolean;
  attachments: RawAttachment[];
  /** Image attachments' actual bytes, as data URLs — see attachments.ts. */
  images: { name: string; dataUrl: string }[];
}

/**
 * Scans the left rail for conversations you could choose to track. The
 * conversation id comes from the treeitem's `data-tabster` attribute
 * (a JSON blob Teams uses for its own focus management, but it happens to
 * carry the real Teams thread id at observed.names[0] — more reliable
 * than trying to parse the composite data-fui-tree-item-value string).
 */
export async function discoverConversations(page: Page): Promise<DiscoveredConversation[]> {
  const items = await page.$$eval(selectors.conversationListItem, (nodes) =>
    nodes.map((node) => {
      let id = "";
      try {
        const tabster = JSON.parse(node.getAttribute("data-tabster") || "{}");
        id = tabster?.observed?.names?.[0] ?? "";
      } catch {
        id = "";
      }
      return {
        id,
        title: (node.textContent || "").trim(),
        kind: node.getAttribute("data-item-type") === "channel" ? "channel" : "chat",
      };
    })
  );
  return items.filter((i) => i.id) as DiscoveredConversation[];
}

/**
 * Opens a conversation by clicking its left-rail entry, and waits for the
 * message pane's actual content to settle before returning.
 *
 * Verification is two-layered:
 * 1. Wait for `[id="chat-header-<conversationId>"]` to be attached to the
 *    DOM — Teams renders this element (in the chat header/toolbar area,
 *    confirmed live via a full-page HTML search) with the open
 *    conversation's own id baked into its element id. This is a direct,
 *    unambiguous link between the DOM and a specific conversation —
 *    unlike every "did the message list change" heuristic tried before,
 *    which could be fooled by the *previous* conversation itself
 *    receiving new messages during the transition (confirmed live: this
 *    reproduced a real misattributed-message bug even after other
 *    settle-timing fixes). This element only exists for whichever
 *    conversation is genuinely open right now, so waiting for it removes
 *    the "which conversation is this?" ambiguity entirely.
 * 2. Only then check the message *content* for stability — the header can
 *    confirm identity before the message list has finished loading.
 */
export async function openConversation(page: Page, conversationId: string): Promise<void> {
  const item = page.locator(`${selectors.conversationListItem}[data-tabster*="${conversationId}"]`);
  // Bounded, not Playwright's 30s default: a tracked conversation that's
  // no longer in the left rail (archived/removed — confirmed live as an
  // ongoing case, "JFrog Onboarding bi-weekly") would otherwise eat a
  // full 30s on EVERY single attempt, monopolizing the one shared tab.
  // That used to just be a throughput annoyance; it became a measured,
  // concrete latency problem once catchUpAllTracked's priority queue for
  // pending image resolution shipped — a priority item unlucky enough to
  // queue up behind an in-flight 30s failure waited the full 30s too,
  // confirmed live (32s end-to-end for a test send). A conversation that
  // genuinely exists is already rendered in the left rail from the
  // initial discovery scan, no extensive scroll/pagination needed to find
  // it, so 8s is generous for the real case while capping the cost of the
  // "no longer exists" case at a fraction of before.
  await item.first().click({ timeout: 8000 });
  await waitForMessagePane(page);
  await waitForConversationHeader(page, conversationId);
  await waitForStableMessageList(page);
}

async function readMessageIds(page: Page): Promise<string> {
  return page
    .$$eval(selectors.messageListItem, (nodes) => nodes.map((n) => n.getAttribute("data-mid")).join(","))
    .catch(() => "");
}

async function waitForConversationHeader(page: Page, conversationId: string, timeoutMs = 15_000): Promise<void> {
  await page.locator(`[id="chat-header-${conversationId}"]`).first().waitFor({ state: "attached", timeout: timeoutMs });
}

/**
 * Polls the rendered message ids until they stop changing across a few
 * consecutive checks, so callers can trust the pane has finished swapping
 * to whichever conversation was just opened.
 */
async function waitForStableMessageList(
  page: Page,
  { requiredStableChecks = 3, intervalMs = 200, timeoutMs = 5000 } = {}
): Promise<void> {
  const start = Date.now();
  let lastIds = "";
  let stableCount = 0;
  while (Date.now() - start < timeoutMs) {
    const ids = await readMessageIds(page);
    if (ids === lastIds) {
      stableCount++;
      if (stableCount >= requiredStableChecks) return;
    } else {
      stableCount = 0;
      lastIds = ids;
    }
    await page.waitForTimeout(intervalMs);
  }
}

/**
 * Waits for the message pane to actually render. Needed both after
 * clicking a conversation and after navigating a fresh tab straight to a
 * deep link — in the latter case `domcontentloaded` fires well before the
 * SPA has hydrated and rendered the thread.
 */
export async function waitForMessagePane(page: Page, timeoutMs = 30_000): Promise<void> {
  await page.locator(selectors.messagePane).waitFor({ state: "visible", timeout: timeoutMs });
}

export function hashContent(content: string): string {
  return crypto.createHash("sha256").update(content).digest("hex");
}

/**
 * In-memory cache of a chatsvc bearer token, captured off the Authorization
 * header of any real chatsvc request this page makes — either one of our
 * own REST fast-path sends (self-perpetuating: each send re-caches the
 * token it used) or the DOM path's compose-box POST below (so the first
 * send after a restart, which has no cached token yet and must fall back
 * to the slow DOM path, still warms the cache as a side effect, making
 * every send after it eligible for the fast path again). Confirmed live
 * this is necessary, not just belt-and-suspenders: reading a pre-existing
 * token out of localStorage (an earlier version of this) only worked
 * because that debug session happened to be ~20 minutes old already: a
 * freshly restarted session's localStorage has no chatsvcagg-scoped token
 * yet, and driving the DOM compose box directly (fill + Enter, no
 * different from what sendMessage below already does) does NOT cause one
 * to appear there either — confirmed live across two separate attempts
 * ~45s apart. Whatever mechanism actually mints and caches that token in
 * long-lived localStorage runs on its own schedule, not in response to a
 * send — so the only reliable source is capturing it directly off a live
 * request as it goes out, not fishing for it in storage after the fact.
 */
let cachedToken: { secret: string; expiresAtMs: number } | null = null;
const tokenListenerAttached = new WeakSet<Page>();

function decodeJwtExpiryMs(token: string): number | null {
  try {
    const payload = token.split(".")[1];
    const json = Buffer.from(payload.replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8");
    const exp = JSON.parse(json).exp;
    return typeof exp === "number" ? exp * 1000 : null;
  } catch {
    return null;
  }
}

/** Idempotent per page — a second call on the same page is a no-op. */
function ensureChatTokenCapture(page: Page): void {
  if (tokenListenerAttached.has(page)) return;
  tokenListenerAttached.add(page);
  page.on("request", (req) => {
    if (req.method() !== "POST" || !req.url().includes("/api/chatsvc/")) return;
    const auth = req.headers()["authorization"];
    const token = auth?.match(/^Bearer\s+(.+)$/i)?.[1];
    if (!token) return;
    const expiresAtMs = decodeJwtExpiryMs(token) ?? Date.now() + 5 * 60_000;
    cachedToken = { secret: token, expiresAtMs };
  });
}

function getCachedChatServiceToken(): string | null {
  const minRemainingMs = 60_000;
  if (!cachedToken || cachedToken.expiresAtMs - Date.now() < minRemainingMs) return null;
  return cachedToken.secret;
}

/**
 * Sends a message from the dashboard's chatbox: opens the conversation
 * (same as any other visit) and types into the real compose box, exactly
 * as if you'd typed it yourself in Teams — this account's own send, not a
 * synthetic API call, so it's visible to real recipients immediately and
 * can't be taken back once Enter is pressed. Callers must have already
 * confirmed the user actually wants to send (see the dashboard's compose
 * bar) — this function itself has no confirmation step of its own.
 *
 * `.fill()` on a contenteditable works the same way it did for the manual
 * test send used to verify the catch-up fix earlier — Playwright treats it
 * like any other editable element, no special-casing needed.
 */

function toComposeHtml(text: string): string {
  return `<p>${text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\n/g, "<br>")}</p>`;
}

/**
 * Fast path for sending: POSTs straight to Teams' own chatsvc REST API
 * using the bearer token above, instead of opening the conversation and
 * typing into the compose box. Confirmed live: completes in under a
 * second (808ms round trip observed), vs. several seconds for the DOM path
 * below, largely because it skips openConversation's click/wait-for-
 * header/wait-for-stable-list sequence entirely — the message shows up
 * back in the DOM afterward via Teams' own real-time push, same as any
 * other message.
 *
 * Doesn't touch the page's rendered DOM or navigation state at all — the
 * token cache is plain in-memory state, and page.request is a
 * context-level API independent of whatever conversation happens to be
 * open — so this is safe for callers to run WITHOUT queuing behind
 * withPage's shared-tab lock (see sendMessageTo in watcher/index.ts). It
 * can't collide with an in-flight visit the way typing into the compose
 * box could.
 *
 * Never throws: returns false on anything short of a clean success (no
 * cached token, a non-OK response, a network error), so callers can
 * silently fall back to the slower but more resilient DOM path below.
 */
export async function sendMessageViaRest(page: Page, conversationId: string, text: string): Promise<boolean> {
  try {
    ensureChatTokenCapture(page);
    const token = getCachedChatServiceToken();
    if (!token) {
      console.log("[send] no cached chatsvc token found (cold cache or expired)");
      return false;
    }

    const body = {
      id: "-1",
      type: "Message",
      conversationid: conversationId,
      conversationLink: `blah/${conversationId}`,
      messagetype: "RichText/Html",
      contenttype: "Text",
      content: toComposeHtml(text),
      clientmessageid: `${Date.now()}${Math.floor(Math.random() * 1000)}`,
      properties: { importance: "", subject: "", cards: "[]", links: "[]", mentions: "[]", files: "[]" },
    };

    // "au" is this account's own region prefix (confirmed live by
    // capturing a real send's request URL) — not a general Teams API
    // constant. A different tenant/region would need this adjusted.
    const res = await page.request.post(
      `https://teams.microsoft.com/api/chatsvc/au/v1/users/ME/conversations/${encodeURIComponent(conversationId)}/messages`,
      {
        headers: {
          authorization: `Bearer ${token}`,
          "content-type": "application/json",
          behavioroverride: "redirectAs404",
        },
        data: JSON.stringify(body),
        timeout: 8000,
      }
    );
    if (!res.ok()) {
      console.log(`[send] REST attempt got non-OK status ${res.status()}`);
    }
    return res.ok();
  } catch (err) {
    console.log(`[send] REST attempt threw: ${err instanceof Error ? err.message : String(err)}`);
    return false;
  }
}

export async function sendMessage(page: Page, conversationId: string, text: string): Promise<void> {
  // Piggybacks the token-capture listener on this path too, so a send that
  // has to fall back here (no cached token yet) still warms the cache as a
  // side effect — the compose box's own POST to chatsvc carries a bearer
  // token same as the REST fast path's own requests do.
  ensureChatTokenCapture(page);
  await openConversation(page, conversationId);
  const editor = page.locator(selectors.composeBox);
  await editor.click();
  await editor.fill(text);
  await page.keyboard.press("Enter");
  // Best-effort confirmation that the send actually went through — Teams
  // clears the compose box once the message is accepted. Not fatal if this
  // times out: the message may still have sent, and the caller has no
  // better signal to fall back to anyway (no API response to check, this
  // is DOM automation).
  await page
    .waitForFunction(
      (sel) => {
        const el = document.querySelector(sel);
        return !!el && (el.textContent || "").trim() === "";
      },
      selectors.composeBox,
      { timeout: 5000 }
    )
    .catch(() => {});
}

/**
 * Reads every message currently rendered in the open thread's DOM.
 *
 * Author/timestamp/edited-marker elements are looked up by id
 * (<prefix><data-mid>) rather than queried as descendants — Teams groups
 * consecutive messages from one sender under a single shared header, so
 * those elements often aren't nested inside the message body element they
 * describe (confirmed live: message-author-name span was NOT a descendant
 * of chat-pane-message, but shared its data-mid as an id suffix).
 */
export async function extractVisibleMessages(page: Page): Promise<RawMessage[]> {
  const raw = await page.$$eval(
    selectors.messageListItem,
    async (nodes, args) => {
      // Every helper below is an inline, immediately-invoked, ANONYMOUS
      // function expression — never bound to a name via `function foo(){}`
      // or `const foo = () => {}`. Confirmed live: either of those forms
      // throws "ReferenceError: __name is not defined" the moment this
      // runs against the real page, for every conversation. esbuild's
      // name-preservation transform wraps any nameable function (a
      // declaration, or an arrow assigned to a variable — JS itself infers
      // that variable's name onto the function) in a call to its `__name`
      // helper; that helper lives at module scope in the bundle, but
      // Playwright's $$eval serializes only this callback's own source
      // text (via toString()) to run standalone in the page, where the
      // helper doesn't exist. An IIFE has no name to preserve, so it's
      // never wrapped in the first place. Same reasoning applies to
      // `nodes.map(async (node) => ...)` below — it's async (to read an
      // image blob's bytes per message) but still just an inline argument,
      // never assigned to anything.
      return await Promise.all(nodes.map(async (node) => {
        const mid = node.getAttribute("data-mid") || "";
        const authorEl = document.getElementById(args.authorPrefix + mid);
        const bodyEl = node.querySelector(args.bodySelector) as HTMLElement | null;
        const editedEl = document.getElementById(args.editedPrefix + mid);
        // A quoted/replied-to message renders as its own structured card
        // (data-tid="quoted-reply-card") followed by the actual reply as
        // separate sibling element(s) of bodyEl — extracted structurally
        // here (author/timestamp/text pulled from their own elements),
        // not guessed from flattened text the way this used to work. That
        // guess (matching a "Name + timestamp" pattern at the start of
        // innerText) broke on a real self-quote where the timestamp
        // element existed in the DOM but rendered invisibly (confirmed
        // live — never showed up in innerText at all), leaving no
        // reliable text pattern to match against and no way to recover
        // from it at render time. This can't have that failure mode: it
        // never depends on the timestamp actually being visible.
        //
        // The reply portion is computed by innerText-ing every direct
        // child of bodyEl EXCEPT the one containing the quote card,
        // rather than innerText-ing bodyEl as a whole and trying to strip
        // the quote back out of the flattened string afterward. Every one
        // of those children is already attached and rendered exactly
        // where Teams put it — nothing here is moved, hidden, or cloned
        // out of the live page — so innerText on each is exactly as
        // reliable as it is on bodyEl itself.
        const quoteCard = bodyEl?.querySelector(args.quotedReplyCard) ?? null;
        let quote: { author: string; timestamp: string; text: string } | null = null;
        let content: string;

        // Declared here — before content is built, not after (contrast
        // with the old layout) — so the plain-text branch below can push
        // into them itself while walking the body's own children. That's
        // what lets a pasted image end up as a PENDING_ATTACHMENT
        // placeholder sitting exactly where it was in the message, instead
        // of the old approach: read all the text first via one
        // bodyEl.innerText call, separately query the WHOLE body for every
        // image afterward, and always report them trailing all the text
        // regardless of which line they were actually under. Confirmed
        // live: a "123"/image, "456"/image, "789"/image message (three
        // separate paragraphs, one image pasted under each) rendered as
        // "123 456 789" followed by all three images bundled in one block
        // — technically nothing was lost, but nothing lined up with Teams'
        // own per-line layout either. attachments.ts resolves
        // PENDING_ATTACHMENT back into a real [[ATTACHMENT|name|url]]
        // marker IN PLACE, one level up, once it knows the final URL.
        const attachments: RawAttachment[] = [];
        const images: { name: string; dataUrl: string }[] = [];
        let pastedIndex = 0;

        if (bodyEl && quoteCard) {
          const timestampEl = quoteCard.querySelector(args.quotedReplyTimestamp);
          const authorEl = timestampEl?.previousElementSibling ?? null;
          const contentEl = quoteCard.querySelector(args.quotedReplyContent);
          quote = {
            author: (authorEl?.textContent || "").trim(),
            timestamp: (timestampEl?.textContent || "").trim(),
            text: (contentEl?.textContent || "").trim(),
          };

          let quoteTopLevel: Element = quoteCard;
          while (quoteTopLevel.parentElement && quoteTopLevel.parentElement !== bodyEl) {
            quoteTopLevel = quoteTopLevel.parentElement;
          }
          const replyParts: string[] = [];
          for (const child of Array.from(bodyEl.children)) {
            if (child === quoteTopLevel) continue;
            const text = ((child as HTMLElement).innerText || child.textContent || "").trim();
            if (text) replyParts.push(text);
          }
          content = `[[QUOTE|${quote.author}|${quote.timestamp}|${quote.text}]]\n${replyParts.join("\n\n")}`;
        } else if (bodyEl) {
          // One pass over each top-level child (Teams renders one per
          // line) instead of a single whole-body innerText read, so a
          // pasted image's own child element can be swapped for its
          // placeholder right there in the loop — see the comment on
          // attachments/images above for why this matters.
          const parts: string[] = [];
          for (const child of Array.from(bodyEl.children)) {
            const inlineImages = Array.from(
              child.querySelectorAll('[itemtype="http://schema.skype.com/AMSImage"] img')
            ) as HTMLImageElement[];
            const leftoverText = ((child as HTMLElement).innerText || child.textContent || "").trim();
            if (!inlineImages.length) {
              if (leftoverText) parts.push(leftoverText);
              continue;
            }
            if (leftoverText) parts.push(leftoverText);
            for (const imgEl of inlineImages) {
              pastedIndex++;
              const remoteUrl = imgEl.getAttribute("data-orig-src") || "";
              const objectIdMatch = /\/objects\/([^/]+)\//.exec(remoteUrl);
              const objectId = objectIdMatch?.[1] || imgEl.id.replace(/^x_/, "") || `unknown-${pastedIndex}`;
              const name = `pasted-image-${objectId}`;
              attachments.push({ name, url: remoteUrl });
              parts.push(`[[PENDING_ATTACHMENT|${name}]]`);
              if (imgEl.src.startsWith("blob:")) {
                try {
                  const res = await fetch(imgEl.src);
                  const blob = await res.blob();
                  const dataUrl: string = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result as string);
                    reader.onerror = () => reject(reader.error);
                    reader.readAsDataURL(blob);
                  });
                  images.push({ name, dataUrl });
                } catch {
                  // Same as the file-attachment case below — the
                  // (unloadable without auth) remote reference still made
                  // it into attachments[], just without bytes to embed.
                }
              }
            }
          }
          content = parts.join("\n\n");
        } else {
          content = "";
        }

        // Mentions and attachments aren't part of the DB schema (see
        // renderBody in app.js) — they're appended to the plain-text
        // content as trailing `[[...]]` marker lines instead, the same
        // render-time-detected approach the quote/reply split already
        // uses, so no migration was needed to add them.
        //
        // A mentioned person's display name is split across one sibling
        // element per WORD, all sharing the same data-person-mri
        // (confirmed live: "Tung Nguyen Ba" is three separate elements
        // joined by literal &nbsp; text nodes, not one element holding the
        // full name) — group consecutive same-mri siblings back into a
        // single name. An @Everyone mention is a single, unsplit element,
        // so it falls out of the same loop as a one-element group for
        // free.
        const mentionNames: string[] = (() => {
          if (!bodyEl) return [];
          const els = Array.from(bodyEl.querySelectorAll(`[${args.mentionAttr}]`));
          const seen = new Set<Element>();
          const names: string[] = [];
          for (const el of els) {
            if (seen.has(el)) continue;
            const mri = el.getAttribute(args.mentionAttr);
            const parts: string[] = [];
            let cur: Element | null = el;
            while (cur && cur.getAttribute(args.mentionAttr) === mri) {
              seen.add(cur);
              const text = (cur.textContent || "").trim();
              if (text) parts.push(text);
              cur = cur.nextElementSibling;
            }
            const name = parts.join(" ").trim();
            if (name) names.push(name);
          }
          return names;
        })();
        if (mentionNames.length) content += `\n[[MENTIONS|${mentionNames.join("|")}]]`;

        // Attachments render as a SIBLING of the message body, not inside
        // it — the body is empty for an attachment-only message, which is
        // why one showed up as a completely blank message before this.
        // See selectors.ts's attachmentGrid comment for the two confirmed
        // shapes an item can have (a document card vs. a real image) and
        // why they need different handling. Markers aren't baked into
        // `content` here anymore (unlike mentions above) — deciding a real
        // image's final URL needs DB access to check for/save a local
        // snapshot (see attachments.ts), which this DOM-only layer
        // deliberately doesn't have; that happens one level up, in
        // watcher/index.ts. (attachments/images themselves are declared
        // further up, alongside the content-building walk — see there.)
        const grid = node.querySelector(args.attachmentGrid);
        if (grid) {
          for (const item of Array.from(grid.querySelectorAll(args.attachmentItem))) {
            const previewUrl = item.getAttribute(args.attachmentPreviewUrlAttr);
            const imgEl = item.querySelector("img") as HTMLImageElement | null;
            if (previewUrl) {
              const name = item.getAttribute("aria-label") || "image";
              attachments.push({ name, url: previewUrl });
              // The thumbnail's own src is a page-local blob: URL — reading
              // its bytes needs no auth at all (unlike previewUrl, which
              // 401s: confirmed live with a plain curl), since the browser
              // already fetched and decoded it to render the thumbnail
              // that's sitting right here in the DOM.
              if (imgEl?.src.startsWith("blob:")) {
                try {
                  const res = await fetch(imgEl.src);
                  const blob = await res.blob();
                  const dataUrl: string = await new Promise((resolve, reject) => {
                    const reader = new FileReader();
                    reader.onload = () => resolve(reader.result as string);
                    reader.onerror = () => reject(reader.error);
                    reader.readAsDataURL(blob);
                  });
                  images.push({ name, dataUrl });
                } catch {
                  // No image bytes this time — attachments[] above still
                  // carries the (unloadable without auth) remote URL, so
                  // the attachment reference itself isn't lost.
                }
              }
              continue;
            }
            const titleEl = item.querySelector(args.attachmentTitle);
            const label = titleEl?.getAttribute("aria-label") || "";
            const [labelName, labelUrl] = label.split("\n").map((s) => s.trim());
            attachments.push({
              name: labelName || item.getAttribute("aria-label") || "attachment",
              url: labelUrl || "",
            });
          }
        }

        // A directly-pasted (Ctrl+V) image is a THIRD, distinct shape —
        // confirmed live it is NOT a file-attachment-grid sibling at all
        // (that selector matches nothing for one), but an actual `<img>`
        // rendered INSIDE the message body itself, wrapped in a span
        // carrying itemtype="http://schema.skype.com/AMSImage". Since it
        // has no visible text, this is exactly why a pasted image showed
        // up as a completely empty message (not even an [[ATTACHMENT]]
        // marker) before this — bodyEl.innerText saw nothing there at all.
        // No real filename exists for a pasted image (unlike an uploaded
        // file), so a name is synthesized here from Teams' own object id —
        // NOT from position in the loop. It used to be positional
        // ("pasted-image-1", "pasted-image-2", ...), which is exactly what
        // caused a confirmed-live bug: editing a multi-image message to
        // remove any image other than the last one shifts every later
        // image's DOM position by one, so its positional name collides
        // with an EXISTING attachment_snapshots row already saved under
        // that name for the ORIGINAL (now different) image that used to
        // sit there — resolveAttachmentMarkers then silently reuses those
        // stale bytes instead of downloading the real current image,
        // which showed up as "wrong image"/"wrong order" after an edit.
        // The object id (embedded in both data-orig-src and the element's
        // own id, "x_<objectId>") identifies the actual image itself and
        // survives edits/reorders/revisits unchanged, so the same photo
        // always maps to the same name regardless of what position it's
        // currently rendered in. attachments.ts appends the real extension
        // once it knows the actual content type from the captured bytes.
        //
        // Only reached for the QUOTE branch above — the plain-text branch
        // already walked every top-level child itself and captured any
        // inline image there, in position (see the content-building walk
        // up top). A quoted reply's own inline image doesn't get that same
        // positional treatment (it still ends up trailing the reply text,
        // same as before this change) — lower priority than the plain-text
        // case actually reported, and quote + inline-image is rare enough
        // combined that it wasn't worth doubling the positional-walk logic
        // for here too. It's still captured either way, just not
        // positioned.
        if (bodyEl && quoteCard) {
          const inlineImages = Array.from(
            bodyEl.querySelectorAll('[itemtype="http://schema.skype.com/AMSImage"] img')
          ) as HTMLImageElement[];
          for (const imgEl of inlineImages) {
            pastedIndex++;
            const remoteUrl = imgEl.getAttribute("data-orig-src") || "";
            const objectIdMatch = /\/objects\/([^/]+)\//.exec(remoteUrl);
            const objectId = objectIdMatch?.[1] || imgEl.id.replace(/^x_/, "") || `unknown-${pastedIndex}`;
            const name = `pasted-image-${objectId}`;
            attachments.push({ name, url: remoteUrl });
            if (imgEl.src.startsWith("blob:")) {
              try {
                const res = await fetch(imgEl.src);
                const blob = await res.blob();
                const dataUrl: string = await new Promise((resolve, reject) => {
                  const reader = new FileReader();
                  reader.onload = () => resolve(reader.result as string);
                  reader.onerror = () => reject(reader.error);
                  reader.readAsDataURL(blob);
                });
                images.push({ name, dataUrl });
              } catch {
                // Same as above — the (unloadable without auth) remote
                // reference still made it into attachments[].
              }
            }
          }
        }

        return {
          messageId: mid,
          sender: authorEl ? (authorEl.textContent || "").trim() : null,
          attachments,
          images,
          content,
          editedFlag: !!(editedEl && (editedEl.textContent || "").trim()),
        };
      }));
    },
    {
      authorPrefix: selectors.messageAuthorIdPrefix,
      editedPrefix: selectors.messageEditedMarkerIdPrefix,
      bodySelector: selectors.messageBody,
      mentionAttr: selectors.mentionAttr,
      attachmentGrid: selectors.attachmentGrid,
      attachmentItem: selectors.attachmentItem,
      attachmentTitle: selectors.attachmentTitle,
      attachmentPreviewUrlAttr: selectors.attachmentPreviewUrlAttr,
      quotedReplyCard: selectors.quotedReplyCard,
      quotedReplyTimestamp: selectors.quotedReplyTimestamp,
      quotedReplyContent: selectors.quotedReplyContent,
    }
  );

  return raw
    .filter((m) => m.messageId)
    .map((m) => ({
      ...m,
      isDeleted: selectors.deletedMessageTextPatterns.some((p) => p.test(m.content)),
    }));
}

/**
 * Like extractVisibleMessages, but re-reads until two independent reads
 * (spaced apart) agree, instead of trusting a single snapshot.
 *
 * Specifically for baseline-setting on a freshly tracked conversation:
 * that step doesn't go through recordObservedMessage (only the newest id
 * gets used, to set the watermark — no message row is written), so the
 * db-level "already owned by a different conversation" guard has nothing
 * to check yet and can't catch a bad read here. Re-confirming at the
 * point of use closed a contamination pattern that openConversation's own
 * internal settle checks did not (both are about the same underlying
 * async-render timing, checked independently rather than trusting one
 * pass through it).
 */
export async function extractVisibleMessagesStable(
  page: Page,
  { attempts = 4, delayMs = 400 } = {}
): Promise<RawMessage[]> {
  let last = await extractVisibleMessages(page);
  for (let i = 0; i < attempts; i++) {
    await page.waitForTimeout(delayMs);
    const next = await extractVisibleMessages(page);
    if (next.map((m) => m.messageId).join(",") === last.map((m) => m.messageId).join(",")) return next;
    last = next;
  }
  return last;
}

export function toObservedMessages(
  conversationId: string,
  raw: RawMessage[],
  observedAt: string
): ObservedMessage[] {
  return raw.map((m) => ({
    messageId: m.messageId,
    conversationId,
    sender: m.sender,
    content: m.content,
    contentHash: hashContent(m.content),
    editedFlag: m.editedFlag,
    isDeleted: m.isDeleted,
    observedAt,
  }));
}

