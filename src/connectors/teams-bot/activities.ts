import { htmlToText } from "../../core/html.js";
import type { DiscoveredConversation, IncrementalMessage } from "../../core/types.js";

/**
 * Turns one Bot Framework "activity" (the JSON Teams POSTs at the bot's
 * endpoint) into something core/ingest understands.
 *
 * Kept as pure functions, separate from the adapter/HTTP plumbing in
 * ./index.ts, because this is the part with all the Teams-shaped guesswork
 * in it — and the part the local harness (scripts/bot-activity-harness.mjs)
 * exercises without needing a real bot registration or tunnel.
 *
 * Only the fields actually used are typed. The real payload has far more,
 * and Teams adds to it over time; anything unrecognized is deliberately
 * ignored rather than treated as an error, since an unknown activity type
 * showing up is normal (typing indicators, reactions, card actions,
 * installation events, and so on all arrive on this same endpoint).
 */

export interface TeamsAttachment {
  contentType?: string;
  contentUrl?: string;
  name?: string;
  content?: unknown;
}

export interface TeamsActivity {
  type?: string;
  id?: string;
  text?: string;
  textFormat?: string;
  timestamp?: string;
  from?: { id?: string; name?: string; aadObjectId?: string };
  recipient?: { id?: string; name?: string };
  conversation?: { id?: string; conversationType?: string; name?: string; isGroup?: boolean; tenantId?: string };
  channelData?: {
    eventType?: string;
    team?: { id?: string; name?: string };
    channel?: { id?: string; name?: string };
  };
  membersAdded?: Array<{ id?: string }>;
  membersRemoved?: Array<{ id?: string }>;
  attachments?: TeamsAttachment[];
}

/**
 * A channel reply carries the thread it belongs to in the conversation id
 * itself, as a `;messageid=<root id>` suffix — so the same channel arrives
 * under a different conversation id for every thread in it. Stripped so a
 * channel is ONE tracked conversation with all its threads in it, which is
 * both how a person thinks about a channel and how the teams-web connector
 * already treats conversation ids (a group chat's id has no such suffix, so
 * this is a no-op there).
 *
 * The consequence worth knowing: threading is flattened. A reply is stored
 * as just another message in the channel, with no record of which thread it
 * belonged to, because there's nowhere in the schema to put that yet.
 */
export function normalizeConversationId(rawId: string): string {
  const cut = rawId.indexOf(";messageid=");
  return cut === -1 ? rawId : rawId.slice(0, cut);
}

/** DB `kind` for a conversation. A 1:1 chat with the bot and a group chat are both "chat"; only a real team channel is "channel". */
function conversationKind(activity: TeamsActivity): "chat" | "channel" {
  return activity.conversation?.conversationType === "channel" ? "channel" : "chat";
}

/**
 * Best available human-readable name. Teams is inconsistent here: a group
 * chat with no name set sends no name at all, and a team's default channel
 * ("General") often sends `channel.id` with no `channel.name`. Falls back
 * through progressively worse options rather than showing a raw thread id
 * in the dashboard's conversation list.
 */
function conversationTitle(activity: TeamsActivity): string {
  const team = activity.channelData?.team?.name;
  const channel = activity.channelData?.channel?.name;
  if (team) return channel ? `${team} / ${channel}` : `${team} / General`;
  if (activity.conversation?.name) return activity.conversation.name;
  if (activity.conversation?.conversationType === "personal") {
    return activity.from?.name ? `${activity.from.name} (via bot)` : "Direct chat (via bot)";
  }
  return activity.conversation?.isGroup ? "Group chat (via bot)" : "Chat (via bot)";
}

export function describeConversation(activity: TeamsActivity): DiscoveredConversation | null {
  const rawId = activity.conversation?.id;
  if (!rawId) return null;
  return { id: normalizeConversationId(rawId), title: conversationTitle(activity), kind: conversationKind(activity) };
}

/**
 * Teams file/image attachments, reduced to the {name, url} pairs
 * core/attachments.ts builds its `[[ATTACHMENT|…]]` markers from.
 *
 * A file that was uploaded carries its real download URL one level down in
 * `content.downloadUrl`; an inline/pasted image instead uses the top-level
 * `contentUrl`. Rich cards are skipped entirely — they aren't attachments
 * in any sense the dashboard could render, and a bot receives them mostly
 * as echoes of its own or another bot's output.
 *
 * NOTE: no bytes are downloaded, unlike the teams-web connector, which
 * reads an image's actual bytes out of the blob the browser already
 * decoded. These URLs generally need the bot's own bearer token to fetch,
 * so an image attachment stored by this connector renders as a link rather
 * than an inline preview. See the README's connector comparison.
 */
export function extractAttachments(activity: TeamsActivity): { name: string; url: string }[] {
  const out: { name: string; url: string }[] = [];
  for (const att of activity.attachments ?? []) {
    const contentType = att.contentType ?? "";
    if (contentType.startsWith("application/vnd.microsoft.card.")) continue;

    if (contentType === "application/vnd.microsoft.teams.file.download.info") {
      const downloadUrl = (att.content as { downloadUrl?: string } | undefined)?.downloadUrl;
      if (downloadUrl && att.name) out.push({ name: att.name, url: downloadUrl });
      continue;
    }

    if (contentType.startsWith("image/") && att.contentUrl) {
      // An inline image usually has no name at all — synthesize a stable-ish
      // one so the marker has something to display and key on.
      out.push({ name: att.name ?? `image-${out.length + 1}`, url: att.contentUrl });
    }
  }
  return out;
}

/**
 * Message text as plain text. `activity.text` is usually already plain, but
 * Teams sends real HTML when a message has any formatting at all (bold, a
 * link, a mention), flagged with textFormat "xml" — and sometimes without
 * flagging it. Sniffing for a tag rather than trusting textFormat, then
 * reusing the same htmlToText the teams-web connector's push path uses, so
 * both connectors produce comparable stored content.
 */
export function activityText(activity: TeamsActivity): string {
  const text = activity.text ?? "";
  return /<[a-z][^>]*>/i.test(text) ? htmlToText(text) : text.trim();
}

/**
 * What this activity means for ingest, or null for one we don't record.
 *
 * Teams sends an edit as a `messageUpdate` and a delete as a
 * `messageDelete`, both reusing the ORIGINAL message's id — which is
 * exactly what ingest needs in order to attach a new revision to an
 * existing message rather than creating a second one.
 * `channelData.eventType` distinguishes the sub-cases
 * ("editMessage" / "softDeleteMessage" / "undeleteMessage").
 */
export interface ObservedActivity {
  message: Omit<IncrementalMessage, "content">;
  /** Resolved into `[[ATTACHMENT|…]]` markers by the connector, which is where storage access lives. */
  attachments: { name: string; url: string }[];
  text: string;
}

export function classifyActivity(activity: TeamsActivity): ObservedActivity | null {
  const rawConversationId = activity.conversation?.id;
  const messageId = activity.id;
  if (!rawConversationId || !messageId) return null;

  let kind: IncrementalMessage["kind"];
  if (activity.type === "message") {
    kind = "new";
  } else if (activity.type === "messageUpdate") {
    // Covers both "editMessage" and "undeleteMessage": an un-delete is
    // recorded as an edit back to the restored content, which is the closest
    // the revision model gets to un-deleting, and leaves the delete itself
    // visible in the message's history rather than erasing it.
    kind = "edit";
  } else if (activity.type === "messageDelete") {
    kind = "delete";
  } else {
    return null; // typing indicators, reactions, invokes, installs, conversationUpdate — handled elsewhere or ignored
  }

  const isDelete = kind === "delete";
  return {
    message: {
      kind,
      conversationId: normalizeConversationId(rawConversationId),
      messageId,
      sender: activity.from?.name ?? null,
    },
    attachments: isDelete ? [] : extractAttachments(activity),
    text: isDelete ? "" : activityText(activity),
  };
}

/**
 * True when this activity is Teams telling us the bot itself was just added
 * to a conversation — the bot's entire equivalent of the teams-web
 * connector's left-rail discovery scan. There is nothing to scan: a bot
 * only ever learns about a conversation by being put in one.
 *
 * Compared against `recipient.id` (which is the bot in an inbound activity)
 * rather than a configured app id, so this works identically whether or not
 * credentials are set — including against the local harness.
 */
export function isBotAddedToConversation(activity: TeamsActivity): boolean {
  if (activity.type !== "conversationUpdate") return false;
  const botId = activity.recipient?.id;
  if (!botId) return false;
  return (activity.membersAdded ?? []).some((m) => m.id === botId);
}

/** True when the bot was removed — it will stop receiving anything from this conversation, but whatever was already captured is kept. */
export function isBotRemovedFromConversation(activity: TeamsActivity): boolean {
  if (activity.type !== "conversationUpdate") return false;
  const botId = activity.recipient?.id;
  if (!botId) return false;
  return (activity.membersRemoved ?? []).some((m) => m.id === botId);
}
