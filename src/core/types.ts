/**
 * The contract between a *connector* — whatever actually talks to
 * Microsoft Teams, in BOTH directions: observing messages coming out, and
 * sending messages back in — and the rest of the app (storage +
 * dashboard). Everything in `src/core` and `src/server` is written
 * against these types only, and knows nothing about whether the other end
 * is a browser being driven around the Teams web client, an API, or
 * anything else.
 *
 * The one connector that exists today is `src/connectors/teams-web`
 * (Playwright driving a real logged-in Edge profile). It's deliberately
 * the *only* place that imports Playwright, holds Teams DOM selectors, or
 * knows about Trouter pushes — see the layering note in README.
 */

/**
 * A two-way link to Teams. A connector owns its own transport, its own
 * authentication, and its own scheduling: the app starts it, stops it,
 * and asks it to send — everything in between is the connector's business.
 */
export interface Connector {
  /** Stable identifier, matching the key it's registered under in src/connectors/index.ts. */
  readonly name: string;

  /**
   * Runs until stop() is called, reporting everything it observes through
   * `ctx`. Owns its whole lifecycle internally, INCLUDING any re-auth
   * cycle — if a connector needs a human to re-authenticate, it reports
   * that via ctx.reportAuth and handles the recovery itself rather than
   * returning and expecting the caller to know what to do about it.
   */
  run(ctx: ConnectorContext): Promise<void>;

  /** Shuts the connector down and releases whatever it holds (browser process, socket, token refresh timer). Must be safe to call at any point, including before run(). */
  stop(): Promise<void>;

  /**
   * Sends a message into a conversation — the inbound half of the link,
   * driving the dashboard's chatbox.
   *
   * Optional: a connector that can only read leaves this off, and the send
   * endpoint reports that sending isn't available rather than failing
   * obscurely.
   *
   * Whether a connector records what it just sent depends on whether its
   * own observation path will deliver it back, which differs per connector
   * and is NOT something this interface can decide:
   *
   * - `teams-web` does NOT record it. Teams echoes a sent message back
   *   through the same real-time push as anything anyone else sends, so
   *   recording here would duplicate logic that already handles ownership,
   *   watermarks and notification — two code paths that could drift apart.
   * - `teams-bot` DOES record it. Bot Framework never delivers a bot's own
   *   messages back to it, so nothing else ever would, and the message
   *   would be missing from the dashboard entirely.
   *
   * The rule, then: record it only if nothing else will. Preferring the
   * echo where one exists keeps recording on a single path.
   */
  send?(conversationId: string, text: string): Promise<void>;
}

/** A conversation a connector found on the account. */
export interface DiscoveredConversation {
  id: string;
  title: string;
  kind: "chat" | "channel";
}

/**
 * One message observed as an *incremental* update — a real-time push, a
 * change notification, a delta-query row: something that reports "this
 * one message just changed" rather than "here is the current state of the
 * conversation".
 *
 * Incremental sources are treated as strictly additive: an update at or
 * below the conversation's watermark is dropped rather than recorded (see
 * Ingest.incremental), because such a source is usually a *lossier* view
 * of a message than a full re-read is, and letting it overwrite an
 * already-good stored revision loses information. The teams-web connector
 * is the live example — its pushes can't carry attachments at all.
 */
export interface IncrementalMessage {
  kind: "new" | "edit" | "delete";
  conversationId: string;
  messageId: string;
  sender: string | null;
  /** Plain text, already carrying any `[[ATTACHMENT|…]]` / `[[QUOTE|…]]` markers the dashboard renders (see core/attachments.ts). */
  content: string;
}

/**
 * One message observed as part of a *snapshot* — a full re-read of
 * whatever the connector can currently see of a conversation.
 *
 * Snapshots are authoritative and self-healing: every message in one is
 * re-recorded regardless of the watermark, so an improvement to a
 * connector's own extraction retroactively corrects already-captured
 * messages on the next pass, with no manual backfill. Storage no-ops when
 * content genuinely hasn't changed, so this is cheap.
 */
export interface SnapshotMessage {
  messageId: string;
  sender: string | null;
  content: string;
  edited: boolean;
  deleted: boolean;
}

/** What a single ingest call actually did, so a connector can react to it. */
export interface IngestResult {
  /** A new revision was actually written (content genuinely changed). */
  recorded: boolean;
  /** At least one message was past the conversation's watermark, i.e. real new activity. */
  hasNewMessage: boolean;
  /**
   * The stored content of the affected message *before* this call, or
   * null if there was none. Only populated by `incremental` (a snapshot
   * covers many messages at once); a connector uses it to notice that it
   * may have just overwritten something richer than it can produce — see
   * the teams-web connector's pending-resolution handling.
   */
  priorContent: string | null;
}

/**
 * Whether the connector currently needs a human to do something before it
 * can keep working. `message` is shown nowhere yet (the dashboard banner
 * has fixed copy) but is logged, so a second connector can say what its own
 * remedy actually is instead of inheriting "go look at the Edge window".
 */
export type AuthState = { state: "ok" } | { state: "needs-user-action"; message: string };

/** Everything a connector is given to report what it sees. Supplied by the composition root (src/index.ts). */
export interface ConnectorContext {
  ingest: Ingest;
  reportAuth(auth: AuthState): void;
  /** "I am actively re-reading this conversation right now" — drives the dashboard's per-tab spinner. */
  reportSync(conversationId: string, syncing: boolean): void;
  /** "This specific message's stored content is known-incomplete and I'm working on it" — drives the per-message spinner. */
  reportMessagePending(messageId: string, pending: boolean): void;
}

/** Writes observations into storage. Connector-agnostic; see core/ingest.ts. */
export interface Ingest {
  /** Records conversations a connector discovered, tracking any that aren't tracked yet. Returns the ones that were newly tracked. */
  discover(conversations: DiscoveredConversation[]): DiscoveredConversation[];
  /** Marks a conversation tracked with `messageId` as its baseline, capturing no history before it. */
  baseline(conversation: DiscoveredConversation, messageId: string | null): void;
  incremental(msg: IncrementalMessage): IngestResult;
  snapshot(conversationId: string, messages: SnapshotMessage[]): IngestResult;
  /** Every conversation the app is tracking, for a connector's own scheduling. */
  trackedConversations(): TrackedConversationView[];
  trackedConversation(conversationId: string): TrackedConversationView | undefined;
  /** Records that a conversation was just checked, whether or not anything changed. */
  touchChecked(conversationId: string): void;
}

/** The subset of a tracked conversation a connector needs in order to schedule its own work. */
export interface TrackedConversationView {
  conversationId: string;
  title: string;
  kind: "chat" | "channel";
  /** Newest message id captured so far — the floor for "is this new?". */
  watermarkMessageId: string | null;
  /** ISO timestamp of the last check, or null if never checked. */
  lastCheckedAt: string | null;
}
