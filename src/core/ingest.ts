import crypto from "node:crypto";
import {
  advanceWatermark,
  getLatestRevision,
  getMessageOwner,
  getTrackedConversation,
  listTrackedConversations,
  recordObservedMessage,
  startTracking,
  touchActivity,
  touchSynced,
  type TrackedConversation,
} from "../db/index.js";
import type {
  DiscoveredConversation,
  IncrementalMessage,
  Ingest,
  IngestResult,
  SnapshotMessage,
  TrackedConversationView,
} from "./types.js";

/**
 * The connector-agnostic half of what used to be `watcher/index.ts`:
 * everything about turning "I observed this message" into stored state —
 * ownership guarding, watermarks, revision dedupe, unread/activity
 * bookkeeping, and telling the dashboard something changed.
 *
 * Nothing here knows how the observation was made. The scheduling,
 * transport, auth and extraction that produced it all live in whichever
 * connector is running (src/connectors/*), behind the Connector interface in
 * ./types.ts.
 */

export interface IngestListeners {
  /** A conversation's stored content changed and any dashboard showing it should refresh. */
  onConversationUpdate(conversationId: string): void;
}

export function hashContent(content: string): string {
  return crypto.createHash("sha256").update(content).digest("hex");
}

function toBigIntOrNull(id: string | null | undefined): bigint | null {
  if (!id) return null;
  try {
    return BigInt(id);
  } catch {
    return null; // unparseable id: don't filter on it, better to over-record than silently drop a message
  }
}

function toView(conv: TrackedConversation): TrackedConversationView {
  return {
    conversationId: conv.conversation_id,
    title: conv.title,
    kind: conv.kind,
    watermarkMessageId: conv.last_watermark_message_id,
    lastCheckedAt: conv.last_synced_at,
  };
}

const EMPTY: IngestResult = { recorded: false, hasNewMessage: false, priorContent: null };

export function createIngest(listeners: IngestListeners): Ingest {
  return {
    discover(conversations) {
      return conversations.filter((c) => !getTrackedConversation(c.id));
    },

    baseline(conversation: DiscoveredConversation, messageId: string | null) {
      if (getTrackedConversation(conversation.id)) return; // idempotent — already tracked, don't reset its baseline
      startTracking(conversation.id, conversation.title, conversation.kind);
      const now = new Date().toISOString();
      if (messageId) advanceWatermark(conversation.id, messageId, now);
      touchSynced(conversation.id);
    },

    /**
     * Records one incremental update.
     *
     * Deliberately additive-only: a "new" message at or below the
     * watermark is dropped rather than re-recorded, and an edit/delete of
     * a message this app never captured in the first place (it predates
     * tracking — by design there's no backfill, so there's no original
     * content to meaningfully update) is dropped too. See
     * IncrementalMessage's own comment for why an incremental source
     * isn't allowed to overwrite below the watermark.
     */
    incremental(msg: IncrementalMessage): IngestResult {
      const conv = getTrackedConversation(msg.conversationId);
      if (!conv) return EMPTY; // not tracked yet — discovery will pick it up on its own schedule

      const owner = getMessageOwner(msg.messageId);
      if (owner !== null && owner !== msg.conversationId) return EMPTY; // belongs to another conversation — never trust it

      const floor = toBigIntOrNull(conv.last_watermark_message_id);
      const id = toBigIntOrNull(msg.messageId);

      const isNew = msg.kind === "new" && id !== null && (floor === null || id > floor);
      if (msg.kind === "new") {
        if (id !== null && floor !== null && id <= floor) return EMPTY; // already known or older than our baseline
      } else if (owner !== msg.conversationId) {
        return EMPTY;
      }

      // Fetched BEFORE recordObservedMessage below, which — if this
      // actually changes anything — makes msg.content the new latest
      // revision. Handed back so the connector can tell whether it just
      // wrote over something richer than it's able to produce.
      const priorContent = getLatestRevision(msg.messageId)?.content ?? null;

      const observedAt = new Date().toISOString();
      const rev = recordObservedMessage({
        messageId: msg.messageId,
        conversationId: msg.conversationId,
        sender: msg.sender,
        content: msg.content,
        contentHash: hashContent(msg.content),
        editedFlag: msg.kind === "edit",
        isDeleted: msg.kind === "delete",
        observedAt,
      });

      if (isNew) advanceWatermark(msg.conversationId, msg.messageId, observedAt);
      touchSynced(msg.conversationId);

      if (rev) {
        touchActivity(msg.conversationId, observedAt);
        listeners.onConversationUpdate(msg.conversationId);
      }
      return { recorded: !!rev, hasNewMessage: isNew, priorContent };
    },

    /**
     * Records a full re-read of a conversation — every message given,
     * not just ones past the watermark.
     *
     * This used to pre-filter to `id > floor`, which meant a message
     * already at or below the watermark was never looked at again, so
     * neither a real edit NOR a fix to a connector's own extraction
     * accuracy could ever correct an already-recorded message.
     * Confirmed live: a message captured before the quote/reply
     * extraction fix stayed permanently fused (no separator between the
     * quoted text and the reply) even after the fix shipped and the
     * backup loop kept revisiting that conversation, because it was
     * already below the watermark. recordObservedMessage already no-ops
     * (via its content-hash compare) when nothing actually changed, so
     * re-checking everything on every pass is cheap and self-heals this
     * class of bug going forward without a one-off manual backfill.
     *
     * `hasNewMessage` (id > floor) drives activity/unread — a retroactive
     * content correction to an old message isn't new activity and
     * shouldn't bump sort order or mark the conversation unread.
     * `recorded` (a revision was actually written, new OR corrected)
     * drives the live-update notify, so a dashboard with that
     * conversation open refreshes either way.
     */
    snapshot(conversationId: string, messages: SnapshotMessage[]): IngestResult {
      const conv = getTrackedConversation(conversationId);
      const floor = toBigIntOrNull(conv?.last_watermark_message_id);
      const observedAt = new Date().toISOString();

      let hasNewMessage = false;
      let recorded = false;
      let newFloor = floor;

      for (const msg of messages) {
        const owner = getMessageOwner(msg.messageId);
        if (owner !== null && owner !== conversationId) continue; // belongs to a different conversation — never trust it

        const id = toBigIntOrNull(msg.messageId);
        const isNew = id !== null && (floor === null || id > floor);

        const rev = recordObservedMessage({
          messageId: msg.messageId,
          conversationId,
          sender: msg.sender,
          content: msg.content,
          contentHash: hashContent(msg.content),
          editedFlag: msg.edited,
          isDeleted: msg.deleted,
          observedAt,
        });
        if (rev) {
          recorded = true;
          if (isNew) hasNewMessage = true;
        }

        if (id !== null && (newFloor === null || id > newFloor)) newFloor = id;
      }

      if (newFloor !== floor) advanceWatermark(conversationId, newFloor!.toString(), observedAt);
      touchSynced(conversationId);
      if (hasNewMessage) touchActivity(conversationId, observedAt);
      if (recorded) listeners.onConversationUpdate(conversationId);

      return { recorded, hasNewMessage, priorContent: null };
    },

    trackedConversations: () => listTrackedConversations().map(toView),
    trackedConversation: (id) => {
      const conv = getTrackedConversation(id);
      return conv ? toView(conv) : undefined;
    },
    touchChecked: (id) => touchSynced(id),
  };
}
