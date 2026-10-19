import type { BrowserContext, Page } from "playwright";
import { config } from "../../config.js";
import { resolveAttachmentMarkers } from "../../core/attachments.js";
import type {
  DiscoveredConversation,
  Connector,
  ConnectorContext,
  SnapshotMessage,
  TrackedConversationView,
} from "../../core/types.js";
import { isOnLoginPage, waitUntilLoggedIn, watchForReauth } from "./auth.js";
import { launchTeamsBrowser } from "./browser.js";
import {
  discoverConversations,
  extractVisibleMessagesStable,
  openConversation,
  sendMessage as sendMessageInPage,
  sendMessageViaRest,
} from "./teams.js";
import { attachMessagePushListener, type PushedMessage } from "./trouter.js";
import { selectors } from "./selectors.js";

/**
 * The "teams-web" connector: observes Teams by driving a real, logged-in
 * Microsoft Edge profile with Playwright.
 *
 * Everything Playwright-, DOM-, Trouter- and Edge-profile-specific is
 * contained in this directory. The rest of the app only ever sees the
 * Connector interface (core/types.ts) — which is what makes adding a second
 * connector a matter of writing a sibling directory rather than touching
 * storage, the server or the dashboard.
 *
 * Single-tab design: exactly one browser tab handles login, discovery,
 * tracking, and every ongoing check — no dedicated tab per conversation.
 *
 * Detection has two layers:
 * 1. Real-time push (trouter.ts) — the primary path. Teams delivers
 *    new/edited/deleted messages for every conversation on the account
 *    over one WebSocket connection, so a single passive listener catches
 *    them all near-instantly with no per-conversation visiting at all.
 * 2. The scheduled visit loop below — a slower backup, in case the push
 *    connection ever drops. Every tracked conversation is treated
 *    uniformly here (visited in turn, oldest-checked first); there used to
 *    be a 'live'/'polling' distinction that checked recently-active
 *    conversations more often, but once push covers real-time detection
 *    for everything, that distinction stopped changing anything
 *    meaningful and was removed.
 *
 * This replaced an earlier design that kept a dedicated tab open per live
 * conversation plus a small pool of polling tabs — mainly for the
 * resource cost of many open tabs. A cross-conversation message-leak bug
 * was chased at length under that design and initially blamed on
 * cross-tab DOM confusion, but it turned out to reproduce identically
 * with only one tab open: the real cause was the post-navigation read
 * being a single-shot DOM read instead of the re-confirmed (stable) read
 * used elsewhere — see visitConversation. The single-tab design is kept
 * regardless, on its own resource-usage merits.
 */
export function createTeamsWebConnector(): Connector {
  return new TeamsWebConnector();
}

class TeamsWebConnector implements Connector {
  readonly name = "teams-web";

  // Serializes every operation against the single shared tab — discovery
  // scans, tracking a new conversation, and each scheduled visit all queue
  // through here, so nothing ever runs concurrently against it.
  #pageQueue: Promise<unknown> = Promise.resolve();

  #context: BrowserContext | null = null;
  #page: Page | null = null;
  #stopped = false;

  // Conversations with at least one message that arrived via push in a
  // state push can't fully represent on its own — push can't extract
  // attachments at all, so three shapes count as "probably incomplete,
  // needs a DOM pass to reconcile":
  //   1. A genuinely NEW message with empty content and a real sender —
  //      the signature of an attachment-only message (a real reaction has
  //      an empty SENDER too, and never reaches ingest at all — see
  //      trouter.ts, so this can't misfire on those).
  //   1b. Any push (new or edit) carrying an inline image alongside real
  //      text (trouter.ts's hasInlineImage) — content isn't empty here, so
  //      (1) alone doesn't catch it, but push still can't carry the actual
  //      image bytes, only the plain-text part. Confirmed live: without
  //      this, a "123"/image + "456"/image + "789"/image message sat on the
  //      normal interval-based rotation instead of being prioritized, and
  //      went unresolved for 8+ minutes (past the 5min+jitter backup
  //      interval) on a busy day where other conversations' own churn kept
  //      pushing it back in the queue.
  //   2. An EDIT (or any push) to a message whose previous revision had
  //      attachment markers — push's own content for the edit is always
  //      attachment-blind, so if we just recorded push's version straight
  //      over a revision that had real attachments, whatever push gave us
  //      may have silently dropped them. Confirmed live: editing a
  //      3-image message (removing one, adding text) pushed a plain-text
  //      revision with zero attachment markers — not just the one removed,
  //      all of them — and it stayed that way, because this same
  //      conversation's OWN follow-up push traffic kept resetting the
  //      last-checked timestamp and deferring the backup visit that would
  //      have caught and fixed it, for far longer than "eventually".
  //   3. Any push (new or edit) whose raw content carries a quoted-reply
  //      card (trouter.ts's hasQuote) — push's flattened text is all
  //      app.js's render-time heuristic has to identify the quote from, and
  //      that heuristic requires a parseable timestamp it isn't always
  //      given (confirmed live: a self-quote's raw pushed HTML had no
  //      timestamp text at all for the quoted-reply-card's timestamp
  //      element), so an unresolved quote push doesn't just miss its
  //      timestamp — it renders as broken run-on text until a DOM pass
  //      overwrites it with the structural [[QUOTE|...]] extraction
  //      (teams.ts), which doesn't depend on the timestamp being present.
  // The backup visit loop checks this set first, ahead of its normal
  // interval-based rotation, since a REST/HTTP alternative to DOM
  // extraction isn't viable here — both the chatsvc endpoint and the
  // image's own `amspreviewurl` 401 without an authenticated Teams session
  // neither page.request nor an in-page fetch can supply (confirmed live,
  // see README). DOM extraction via the browser is the only thing that's
  // actually worked; this doesn't replace it, it just stops a busy
  // conversation's frequent push traffic from endlessly deferring the
  // backup visit that would resolve it — every push touches the
  // last-checked timestamp (new message, edit, or delete), so an active
  // conversation's normal due-interval can keep getting pushed back
  // indefinitely, which is exactly the scenario most likely to have
  // something pending in the first place.
  #pendingConversations = new Set<string>();

  // Same signal as #pendingConversations, but keeping the specific message
  // id(s) involved rather than just the conversation — purely for
  // reportMessagePending (the dashboard's per-message "still loading"
  // indicator), which needs to know WHICH message to show a spinner on,
  // not just which conversation to prioritize revisiting. Kept as a
  // second, separate structure rather than folding #pendingConversations
  // into a Map<conversationId, Set<messageId>> so the already-working
  // priority-queue routing in the visit loop's tick() stays untouched.
  #pendingMessages = new Map<string, Set<string>>();

  // Set while the visit loop is running, so a push that just marked a
  // conversation pending can cancel whatever sleep the loop is currently
  // in and make it check again right away — otherwise a pending image
  // could still sit for up to the loop's own idle-poll interval (2s)
  // before anything picks it up, which isn't the "instant" the pending
  // indicator implies.
  #wakeVisitLoop: (() => void) | null = null;

  #withPage<T>(fn: () => Promise<T>): Promise<T> {
    const result = this.#pageQueue.then(fn, fn);
    this.#pageQueue = result.then(
      () => undefined,
      () => undefined
    );
    return result;
  }

  /**
   * Runs until stop(): alternates between a headless session (the normal
   * case — the persistent profile carries the Teams login across restarts)
   * and a headed window opened purely so a human can sign in again when
   * Teams drops back to a login page, which Playwright can't render with
   * no window at all.
   *
   * This loop used to live in src/index.ts. It's here now because it's
   * entirely an artifact of how THIS connector authenticates — another
   * connector's "needs a human" story (an expired token, a device-code
   * prompt) has nothing to do with headless-vs-headed browsers, and the
   * composition root shouldn't have to know the difference.
   */
  async run(ctx: ConnectorContext): Promise<void> {
    while (!this.#stopped) {
      await this.#runHeadlessSession(ctx);
      if (this.#stopped) break;
      await this.#runHeadedLoginSession(ctx);
    }
  }

  async stop(): Promise<void> {
    this.#stopped = true;
    this.#wakeVisitLoop = null;
    const context = this.#context;
    this.#context = null;
    this.#page = null;
    await context?.close().catch(() => {});
  }

  /**
   * Sends a message from the dashboard's chatbox.
   *
   * Tries the REST fast path first (see sendMessageViaRest in teams.ts) —
   * outside #withPage's queue entirely, since it never touches the page's
   * DOM/navigation state and so can't collide with an in-flight visit. Only
   * on that failing (no cached token, expired, network error) does this fall
   * back to the DOM-typing path, which DOES need #withPage: it can't run
   * concurrently with a scheduled visit or catch-up without risking typing
   * into whichever conversation happened to be open for a different reason.
   *
   * Deliberately does NOT record the sent message itself, on either path:
   * Teams' own real-time push (trouter.ts) will deliver it back as a
   * genuine "new message" event within moments, same as anything anyone
   * else sends, which keeps this on the one code path that already handles
   * ownership checks, watermark advancement, and dashboard notification
   * correctly — duplicating that logic here would just be two places that
   * could drift out of sync with each other.
   */
  async send(conversationId: string, text: string): Promise<void> {
    const page = this.#page;
    if (!page) throw new Error("Watcher isn't connected right now — check the Edge window.");

    const t0 = Date.now();
    if (await sendMessageViaRest(page, conversationId, text)) {
      console.log(`[send] REST fast path for "${conversationId}" took ${Date.now() - t0}ms`);
      return;
    }
    await this.#withPage(() => sendMessageInPage(page, conversationId, text));
    console.log(`[send] DOM fallback for "${conversationId}" took ${Date.now() - t0}ms total`);
  }

  /**
   * One headless session: launches, and if already logged in, wires up
   * push detection + catch-up + discovery + the backup visit loop and runs
   * until Teams redirects to a login page. Resolves once the session has
   * torn itself down, either because a login page showed up or the session
   * was already logged out when it launched.
   */
  async #runHeadlessSession(ctx: ConnectorContext): Promise<void> {
    const context = await launchTeamsBrowser(true);
    this.#context = context;
    const [page] = context.pages();
    this.#page = page;
    await page.goto(config.teamsWeb.teamsUrl, { waitUntil: "domcontentloaded" });

    if (await isOnLoginPage(page)) {
      this.#page = null;
      await context.close();
      return;
    }

    console.log("[teams-web] Headless session active.");
    ctx.reportAuth({ state: "ok" });

    await new Promise<void>((resolve) => {
      let stopDiscovery = () => {};
      const stopVisits = this.#startVisitLoop(page, ctx);

      const stopWatching = watchForReauth(page, () => {
        console.log("[teams-web] Teams needs you to log in again — switching to a visible browser.");
        ctx.reportAuth({
          state: "needs-user-action",
          message: "Teams signed out — a visible Edge window is opening for you to log in.",
        });
        this.#page = null; // this page is about to close — nothing left to send through
        stopDiscovery();
        stopVisits();
        stopWatching();
        context.close().finally(resolve);
      });

      this.#startPushListener(page, ctx);
      this.#catchUpAllTracked(page, ctx)
        .then(() => this.#startDiscoveryLoop(page, ctx))
        .then((stop) => {
          stopDiscovery = stop;
        })
        .catch((err) => console.error("[teams-web] startup catch-up/discovery failed:", err));
    });
  }

  /** Opens a real, visible window just for you to sign in, then tears it down once you have. */
  async #runHeadedLoginSession(ctx: ConnectorContext): Promise<void> {
    const context = await launchTeamsBrowser(false);
    this.#context = context;
    const [page] = context.pages();
    await page.goto(config.teamsWeb.teamsUrl, { waitUntil: "domcontentloaded" });
    console.log("[teams-web] A visible Edge window is open — please log in there.");
    await waitUntilLoggedIn(page);
    console.log("[teams-web] Login resolved — switching back to headless.");
    ctx.reportAuth({ state: "ok" });
    await context.close();
  }

  /** Wires up real-time push-based detection for every tracked conversation, via the one Trouter connection this tab already has open. */
  #startPushListener(page: Page, ctx: ConnectorContext): void {
    attachMessagePushListener(page, (msg) => {
      try {
        this.#handlePushedMessage(msg, ctx);
      } catch (err) {
        console.error("[teams-web] failed to handle pushed message:", err);
      }
    });
  }

  /**
   * Handles one real-time new/edit/delete push (see trouter.ts) — the
   * primary detection path, with the scheduled visit loop as a slower
   * backup (in case the push connection ever drops).
   *
   * Runs synchronously against storage only, no page interaction — pushes
   * arrive independently of whatever the visit loop is doing on the shared
   * tab, and don't need to queue behind it.
   */
  #handlePushedMessage(pushed: PushedMessage, ctx: ConnectorContext): void {
    const result = ctx.ingest.incremental({
      kind: pushed.kind,
      conversationId: pushed.conversationId,
      messageId: pushed.messageId,
      sender: pushed.sender,
      content: pushed.content,
    });

    // Empty content + a real sender on a genuinely new message is push's
    // blind spot (see #pendingConversations above), not a no-op — a real
    // reaction/no-op has an empty SENDER too and is already filtered out
    // entirely upstream in trouter.ts, so this can't misfire on those.
    //
    // Separately: if push just recorded a REAL change over a message whose
    // previous revision had attachment markers, push's own content can
    // never account for those — its content is always attachment-blind —
    // so treat that the same as the empty-content case regardless of push
    // kind, most importantly for an edit.
    //
    // Separately again: a message that mixes real text with an inline
    // image (hasInlineImage, see trouter.ts) has non-empty content, so
    // neither check above catches it — but push's plain-text content still
    // can't carry the image itself, so it's the same "needs a DOM pass"
    // case.
    const droppedKnownAttachments = result.recorded && !!result.priorContent?.includes("[[ATTACHMENT|");
    const pushedInlineImage = !!pushed.sender && pushed.hasInlineImage;
    // A quote-bearing push (trouter.ts's hasQuote) has the same problem as
    // an inline image: push's flattened text is all app.js has to work
    // with, and its heuristic split can't reliably find the quote header
    // without a parseable timestamp — which a self-quote's raw HTML can
    // omit entirely (confirmed live, see hasQuote's own comment). Routed
    // through the same pending-resolution fast path so the backup DOM pass
    // (which extracts the quote structurally, timestamp or not — see
    // teams.ts) overwrites it within moments instead of up to the full
    // 5min+jitter backup interval.
    const pushedQuote = !!pushed.sender && pushed.hasQuote;

    // Guarded on the conversation actually being tracked: an untracked
    // conversation's push is dropped by ingest outright, and flagging it
    // pending would put an id in the priority queue that the visit loop
    // then has to keep discarding as stale.
    if (!ctx.ingest.trackedConversation(pushed.conversationId)) return;

    if (
      (pushed.kind === "new" && pushed.content === "" && pushed.sender) ||
      droppedKnownAttachments ||
      pushedInlineImage ||
      pushedQuote
    ) {
      this.#markPending(pushed.conversationId, pushed.messageId, ctx);
    }
  }

  #markPending(conversationId: string, messageId: string, ctx: ConnectorContext): void {
    this.#pendingConversations.add(conversationId);
    let messages = this.#pendingMessages.get(conversationId);
    if (!messages) {
      messages = new Set();
      this.#pendingMessages.set(conversationId, messages);
    }
    messages.add(messageId);
    ctx.reportMessagePending(messageId, true);
    this.#wakeVisitLoop?.();
  }

  /**
   * Clears a conversation's pending state after any real visit, successful
   * or not — a message still unresolved after this either resolved to
   * genuinely nothing or has already scrolled out of the currently-
   * rendered window, and either way retrying it forever wouldn't change
   * that. The dashboard is told to stop spinning on every flagged message
   * regardless of whether this particular pass changed anything (it may
   * already have been correct), rather than leave a spinner up forever.
   */
  #clearPending(conversationId: string, ctx: ConnectorContext): void {
    this.#pendingConversations.delete(conversationId);
    const messages = this.#pendingMessages.get(conversationId);
    if (!messages) return;
    this.#pendingMessages.delete(conversationId);
    for (const messageId of messages) ctx.reportMessagePending(messageId, false);
  }

  /** Reads every currently-visible message out of the open conversation and hands it to ingest as a snapshot. */
  async #readVisible(page: Page, conversationId: string): Promise<SnapshotMessage[]> {
    const raw = await extractVisibleMessagesStable(page);
    return raw.map((m) => ({
      messageId: m.messageId,
      sender: m.sender,
      // Attachment markers are resolved here rather than inside
      // extractVisibleMessages itself — deciding a real image's final URL
      // means checking for (and possibly saving) a local snapshot, which
      // needs storage access that the DOM-extraction layer deliberately
      // doesn't have.
      content: resolveAttachmentMarkers(m.content, m.messageId, m.attachments, m.images),
      edited: m.editedFlag,
      deleted: m.isDeleted,
    }));
  }

  /** Tracks a conversation if it isn't already (idempotent) — baseline watermark only, no backfill. */
  async #autoTrack(page: Page, discovered: DiscoveredConversation, ctx: ConnectorContext): Promise<void> {
    await openConversation(page, discovered.id);
    const raw = await extractVisibleMessagesStable(page);
    ctx.ingest.baseline(discovered, raw.at(-1)?.messageId ?? null);
  }

  /** One discovery pass: finds every conversation in the left rail and auto-tracks any that aren't tracked yet. */
  async #runDiscoveryScan(page: Page, ctx: ConnectorContext): Promise<void> {
    const discovered = await this.#withPage(async () => {
      // Confirmed live: right after startup, the left rail can still be
      // empty in the DOM for a moment even though the page itself has
      // loaded — scanning too early silently finds 0 conversations. Wait
      // for at least one real entry before trusting the scan.
      await page.locator(selectors.conversationListItem).first().waitFor({ state: "visible", timeout: 30_000 });
      return discoverConversations(page);
    });

    // Channels intentionally out of scope for now.
    const chats = discovered.filter((d) => d.kind === "chat");
    const untracked = ctx.ingest.discover(chats);
    for (const d of untracked) {
      await this.#withPage(() => this.#autoTrack(page, d, ctx));
    }
    console.log(`[discovery] scan found ${discovered.length} conversations, auto-tracked ${untracked.length} new`);
  }

  /**
   * Runs discovery once immediately, then on config.teamsWeb.discoveryIntervalMs
   * — covers "everything now" and "anything new later". Returns a function
   * to stop the recurring scan (used when switching to a headed browser for
   * re-auth, so a scan doesn't fire against a page that's about to close).
   */
  async #startDiscoveryLoop(page: Page, ctx: ConnectorContext): Promise<() => void> {
    await this.#runDiscoveryScan(page, ctx).catch((err) =>
      console.error("[teams-web] initial discovery scan failed:", err)
    );
    let stopped = false;
    const timer = setInterval(() => {
      if (stopped) return;
      this.#runDiscoveryScan(page, ctx).catch((err) => console.error("[teams-web] discovery scan failed:", err));
    }, config.teamsWeb.discoveryIntervalMs);
    return () => {
      stopped = true;
      clearInterval(timer);
    };
  }

  /**
   * Recovers messages missed while the app wasn't running — push (the
   * primary detection path) obviously couldn't have caught anything during
   * that downtime, regardless of how active a conversation normally is.
   *
   * This used to call an internal REST API (see core/html.ts's module
   * comment) instead of opening each conversation, on the theory that one
   * authenticated HTTP call would be cheaper than a DOM visit. Confirmed
   * live that it doesn't work: that endpoint 401s on every call
   * ("Authentication failed") because it needs an explicit Skype-protocol
   * auth token the current Teams web client doesn't expose anywhere
   * page.request or an in-page fetch can pick up (it lives encrypted in
   * localStorage, and — also confirmed live — real Teams navigation never
   * calls this endpoint itself anymore, so there's no live request to
   * harvest a working header from either). It was failing silently: no
   * error, just zero messages returned, every time, regardless of how long
   * the gap was — which is exactly the bug this was written to fix.
   *
   * Reuses the same open-and-extract path as the backup visit loop
   * instead — DOM-based, already proven to work, no separate auth path to
   * keep working.
   */
  async #catchUpAllTracked(page: Page, ctx: ConnectorContext): Promise<void> {
    for (const conv of ctx.ingest.trackedConversations()) {
      await this.#withPage(() => this.#visitConversation(page, conv, ctx)).catch((err) =>
        console.error(`[teams-web] catch-up failed for "${conv.title}":`, err)
      );
    }
  }

  /** One scheduled visit: opens the conversation and re-reads it (push already handles the fast path — this is just the backup). */
  async #visitConversation(page: Page, conv: TrackedConversationView, ctx: ConnectorContext): Promise<void> {
    ctx.reportSync(conv.conversationId, true);
    try {
      await openConversation(page, conv.conversationId);
      // The stable (re-confirmed) extraction, not a single-shot read —
      // this turned out to be the actual root cause of a message
      // occasionally getting attributed to the wrong conversation,
      // confirmed live even after moving to a single shared tab (which
      // ruled out every cross-tab theory chased before that):
      // openConversation's own internal settle-check covers the click
      // transition, but this used to do a *separate*, unconfirmed read
      // straight afterward, and Teams' rendering could still drift in the
      // gap between the two. Ingest's ownership guard stays on as a
      // defensive backstop regardless.
      ctx.ingest.snapshot(conv.conversationId, await this.#readVisible(page, conv.conversationId));
      this.#clearPending(conv.conversationId, ctx);
    } finally {
      ctx.reportSync(conv.conversationId, false);
    }
  }

  /**
   * Runs until stopped: repeatedly visits whichever tracked conversation is
   * most overdue for a check, sleeping briefly when nothing is due yet.
   * Every conversation is treated uniformly — oldest last-checked goes
   * first — EXCEPT #pendingConversations, checked first every tick: a
   * push-detected likely-unresolved attachment jumps straight to the front
   * of the queue rather than waiting for its normal turn, which — for a
   * busy conversation — every push touching the last-checked timestamp can
   * defer indefinitely (see that Set's own comment for why). Returns a
   * function to stop scheduling further visits.
   */
  #startVisitLoop(page: Page, ctx: ConnectorContext): () => void {
    let stopped = false;
    let scheduled: ReturnType<typeof setTimeout> | null = null;

    // Cancellable, so #wakeVisitLoop can cut short whatever wait tick() is
    // currently in (most importantly the 2s idle-poll below) instead of
    // that wait running out on its own first — the difference between a
    // pending image resolving "instantly" versus up to 2s later.
    const scheduleTick = (delayMs: number) => {
      if (scheduled) clearTimeout(scheduled);
      scheduled = setTimeout(tick, delayMs);
    };
    this.#wakeVisitLoop = () => scheduleTick(0);

    const tick = async () => {
      if (stopped) return;

      let conv: TrackedConversationView | undefined;
      const pendingId = this.#pendingConversations.values().next().value;
      if (pendingId) {
        conv = ctx.ingest.trackedConversation(pendingId);
        if (!conv) this.#pendingConversations.delete(pendingId); // stale — e.g. untracked since
      }

      if (!conv) {
        const now = Date.now();
        const due = ctx.ingest
          .trackedConversations()
          .filter((c) => isDue(c, now))
          .sort((a, b) => (a.lastCheckedAt ?? "").localeCompare(b.lastCheckedAt ?? ""));

        if (due.length === 0) {
          scheduleTick(2000);
          return;
        }
        conv = due[0];
      }

      let failed = false;
      try {
        await this.#withPage(() => this.#visitConversation(page, conv!, ctx));
      } catch (err) {
        failed = true;
        console.error(`[teams-web] visit failed for "${conv.title}":`, err);
        // Don't keep retrying a broken conversation ahead of everything
        // else every single tick — same backoff-by-falling-back-to-normal-
        // rotation treatment as any other visit failure, and a repeatedly-
        // failing conversation shouldn't leave the dashboard showing a
        // spinner forever on a message this loop can't actually reach.
        this.#clearPending(conv.conversationId, ctx);
      }
      if (stopped) return;
      // A failure (e.g. the browser itself became unavailable) would
      // otherwise retry instantly forever, busy-looping and flooding the
      // log — back off briefly instead of using setImmediate.
      scheduleTick(failed ? 3000 : 0);
    };

    tick();
    return () => {
      stopped = true;
      this.#wakeVisitLoop = null;
      if (scheduled) clearTimeout(scheduled);
    };
  }
}

// Deterministic per-conversation offset (0-30% of the interval), so
// conversations don't stay clustered in lockstep forever. Every
// conversation gets baselined within the same few minutes at startup, so
// without this they'd all become due at nearly the same moment, every
// cycle, indefinitely — a permanent thundering herd on the one shared
// tab. Deterministic (hashed from the id) rather than random so it's
// stable across restarts instead of reshuffling every time.
function jitterFor(conversationId: string, intervalMs: number): number {
  let hash = 0;
  for (let i = 0; i < conversationId.length; i++) {
    hash = (hash * 31 + conversationId.charCodeAt(i)) >>> 0;
  }
  return ((hash % 1000) / 1000) * intervalMs * 0.3;
}

function isDue(conv: TrackedConversationView, now: number): boolean {
  if (!conv.lastCheckedAt) return true;
  const interval =
    config.teamsWeb.backupCheckIntervalMs + jitterFor(conv.conversationId, config.teamsWeb.backupCheckIntervalMs);
  return now - new Date(conv.lastCheckedAt).getTime() >= interval;
}
