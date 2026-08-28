import type { Page } from "playwright";
import { config } from "../config.js";
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
import { resolveAttachmentMarkers } from "./attachments.js";
import {
  discoverConversations,
  extractVisibleMessagesStable,
  hashContent,
  openConversation,
  sendMessage as sendMessageInPage,
  sendMessageViaRest,
  toObservedMessages,
  warmChatServiceToken,
  type DiscoveredConversation,
} from "./teams.js";
import { attachMessagePushListener, type PushedMessage } from "./trouter.js";
import { selectors } from "./selectors.js";

export type ChangeListener = (conversationId: string) => void;
/** Fired around every DOM visit (open + extract) so the dashboard can show which conversation the shared tab is on right now. */
export type SyncListener = (conversationId: string, syncing: boolean) => void;
/** Fired when a specific message enters/leaves pendingResolutionMessagesByConversation, so the dashboard can show a per-message "still loading" indicator (e.g. images not resolved yet) instead of just silently displaying whatever incomplete content push gave us. */
export type MessagePendingListener = (messageId: string, pending: boolean) => void;

/**
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
 *    uniformly here (visited in turn, oldest-synced first); there used to
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
 * with only one tab open: the real cause was extractAndRecordNew doing a
 * single-shot DOM read right after navigating, instead of the
 * re-confirmed (stable) read used elsewhere — see its own comment. The
 * single-tab design is kept regardless, on its own resource-usage merits.
 */

// Serializes every operation against the single shared tab — discovery
// scans, tracking a new conversation, and each scheduled visit all queue
// through here, so nothing ever runs concurrently against it.
let pageQueue: Promise<unknown> = Promise.resolve();
function withPage<T>(fn: () => Promise<T>): Promise<T> {
  const result = pageQueue.then(fn, fn);
  pageQueue = result.then(
    () => undefined,
    () => undefined
  );
  return result;
}

function toBigIntOrNull(id: string | null | undefined): bigint | null {
  if (!id) return null;
  try {
    return BigInt(id);
  } catch {
    return null; // unparseable id: don't filter on it, better to over-record than silently drop a message
  }
}

// Conversations with at least one message that arrived via push in a
// state push can't fully represent on its own — push can't extract
// attachments at all (see handlePushedMessage below and
// renderPendingPlaceholder in app.js), so two shapes both count as
// "probably incomplete, needs a DOM pass to reconcile":
//   1. A genuinely NEW message with empty content and a real sender —
//      the signature of an attachment-only message (a real reaction has
//      an empty SENDER too, and never reaches recordObservedMessage —
//      see trouter.ts, so this can't misfire on those).
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
//      conversation's OWN follow-up push traffic kept resetting
//      touchSynced (see below) and deferring the backup visit that would
//      have caught and fixed it, for far longer than "eventually".
// The backup visit loop checks this set first, ahead of its normal
// interval-based rotation, since a REST/HTTP alternative to DOM
// extraction isn't viable here — both the chatsvc endpoint and the
// image's own `amspreviewurl` 401 without an authenticated Teams session
// neither page.request nor an in-page fetch can supply (confirmed live,
// see README). DOM extraction via the browser is the only thing that's
// actually worked; this doesn't replace it, it just stops a busy
// conversation's frequent push traffic from endlessly deferring the
// backup visit that would resolve it — touchSynced runs on every push
// (new message, edit, or delete), so an active conversation's normal
// due-interval can keep getting pushed back indefinitely, which is
// exactly the scenario most likely to have something pending in the
// first place.
const pendingResolutionConversations = new Set<string>();

// Same signal as pendingResolutionConversations, but keeping the specific
// message id(s) involved rather than just the conversation — purely for
// MessagePendingListener (the dashboard's per-message "still loading"
// indicator), which needs to know WHICH message to show a spinner on, not
// just which conversation to prioritize revisiting. Kept as a second,
// separate structure rather than folding pendingResolutionConversations
// into a Map<conversationId, Set<messageId>> so the already-working
// priority-queue routing in startVisitLoop's tick() stays untouched.
const pendingResolutionMessagesByConversation = new Map<string, Set<string>>();

// Set by startVisitLoop while it's running, so a push that just marked a
// conversation pending can cancel whatever sleep the loop is currently in
// and make it check again right away — otherwise a pending image could
// still sit for up to the loop's own idle-poll interval (2s) before
// anything picks it up, which isn't the "instant" the pending indicator
// implies. Module-level rather than threaded as a parameter since
// handlePushedMessage (called from the push listener, not the visit
// loop) is what needs to reach it.
let wakeVisitLoop: (() => void) | null = null;

/**
 * Handles one real-time new/edit/delete push (see trouter.ts) — the
 * primary detection path now, with the scheduled visit loop as a slower
 * backup (in case the push connection ever drops).
 *
 * Runs synchronously against the DB only, no page interaction — pushes
 * arrive independently of whatever the visit loop is doing on the shared
 * tab, and don't need to queue behind it.
 */
function handlePushedMessage(pushed: PushedMessage, notify: ChangeListener, onMessagePending?: MessagePendingListener): void {
  const conv = getTrackedConversation(pushed.conversationId);
  if (!conv) return; // not tracked yet — discovery will pick it up on its own schedule

  const owner = getMessageOwner(pushed.messageId);
  if (owner !== null && owner !== pushed.conversationId) return; // defensive backstop, shouldn't trigger via push

  const floor = toBigIntOrNull(conv.last_watermark_message_id);
  const id = toBigIntOrNull(pushed.messageId);

  if (pushed.kind === "new") {
    if (id !== null && floor !== null && id <= floor) return; // already known or older than our baseline
  } else if (owner !== pushed.conversationId) {
    // Edit/delete of a message we never recorded in the first place — it
    // predates when tracking began, which (by design — no backfill) means
    // we never captured its original content, so there's nothing to
    // meaningfully update.
    return;
  }

  // Fetched BEFORE recordObservedMessage below, which — if this push
  // actually changes anything — makes pushed.content the new latest
  // revision. Only meaningful in the edit/delete case (a "new" message
  // has no prior revision at all), but cheap enough to always fetch.
  const priorRevision = getLatestRevision(pushed.messageId);

  const observedAt = new Date().toISOString();
  const rev = recordObservedMessage({
    messageId: pushed.messageId,
    conversationId: pushed.conversationId,
    sender: pushed.sender,
    content: pushed.content,
    contentHash: hashContent(pushed.content),
    editedFlag: pushed.kind === "edit",
    isDeleted: pushed.kind === "delete",
    observedAt,
  });

  if (pushed.kind === "new" && id !== null && (floor === null || id > floor)) {
    advanceWatermark(pushed.conversationId, pushed.messageId, observedAt);
  }
  touchSynced(pushed.conversationId);

  // Empty content + a real sender on a genuinely new message is push's
  // blind spot (see pendingResolutionConversations above), not a no-op —
  // a real reaction/no-op has an empty SENDER too and is already filtered
  // out entirely upstream in trouter.ts, so this can't misfire on those.
  //
  // Separately: if push just recorded a REAL change (rev is truthy) over
  // a message whose previous revision had attachment markers, push's own
  // content can never account for those — its content is always
  // attachment-blind — so treat that the same as the empty-content case
  // regardless of push kind, most importantly for an edit.
  //
  // Separately again: a message that mixes real text with an inline image
  // (hasInlineImage, see trouter.ts) has non-empty content, so neither
  // check above catches it — but push's plain-text content still can't
  // carry the image itself, so it's the same "needs a DOM pass" case.
  const droppedKnownAttachments = !!rev && !!priorRevision?.content.includes("[[ATTACHMENT|");
  const pushedInlineImage = !!pushed.sender && pushed.hasInlineImage;
  if (
    (pushed.kind === "new" && pushed.content === "" && pushed.sender) ||
    droppedKnownAttachments ||
    pushedInlineImage
  ) {
    pendingResolutionConversations.add(pushed.conversationId);
    let pendingMessages = pendingResolutionMessagesByConversation.get(pushed.conversationId);
    if (!pendingMessages) {
      pendingMessages = new Set();
      pendingResolutionMessagesByConversation.set(pushed.conversationId, pendingMessages);
    }
    pendingMessages.add(pushed.messageId);
    onMessagePending?.(pushed.messageId, true);
    wakeVisitLoop?.();
  }

  if (rev) {
    touchActivity(pushed.conversationId, observedAt);
    notify(pushed.conversationId);
  }
}

/** Wires up real-time push-based detection for every tracked conversation, via the one Trouter connection this tab already has open. */
export function startPushListener(page: Page, notify: ChangeListener, onMessagePending?: MessagePendingListener): void {
  attachMessagePushListener(page, (msg) => {
    try {
      handlePushedMessage(msg, notify, onMessagePending);
    } catch (err) {
      console.error("[watcher] failed to handle pushed message:", err);
    }
  });
}

/**
 * Extracts and records every currently-visible message — not just ones
 * past the watermark — and advances the watermark for genuinely new ones.
 *
 * Used to pre-filter to only `id > floor` before recording anything at
 * all, which meant a message already at or below the watermark was never
 * looked at again by this backup path once first seen — so neither a real
 * Teams edit NOR a fix to our own extraction accuracy could ever correct
 * an already-recorded message here (push already covered real edits for
 * messages it owns; this path covered neither for anything below the
 * watermark). Confirmed live: a message captured before the quote/reply
 * extraction fix (see teams.ts/chatsvc.ts) stayed permanently fused
 * ("...URL hả anh, được anhvậy a add...", no separator) even after the fix
 * shipped and the backup loop kept revisiting that conversation — because
 * it was already below the watermark, so the old filter skipped it every
 * time. recordObservedMessage already no-ops (via its content-hash
 * compare) when nothing actually changed, so re-checking everything
 * visible on every visit is cheap and self-heals this class of bug
 * automatically, going forward, without a one-off manual backfill per
 * incident.
 *
 * `hasNewMessage` (id > floor) drives touchActivity/unread — a retroactive
 * content correction to an old message isn't new activity and shouldn't
 * bump the conversation's sort order or mark it unread. `contentUpdated`
 * (recordObservedMessage actually wrote a revision, new OR corrected)
 * drives the live-update notify, so a dashboard with that conversation
 * open refreshes to show the corrected text either way.
 *
 * Uses the stable (re-confirmed) extraction, not a single-shot read —
 * this turned out to be the actual root cause of a message occasionally
 * getting attributed to the wrong conversation, confirmed live even after
 * moving to a single shared tab (which ruled out every cross-tab theory
 * chased before that): openConversation's own internal settle-check
 * covers the click transition, but this function used to do a *separate*,
 * unconfirmed read straight afterward, and Teams' rendering could still
 * drift in the gap between the two. The ownership guard (getMessageOwner)
 * stays on as a defensive backstop regardless.
 */
async function extractAndRecordNew(
  page: Page,
  conversationId: string,
  floor: bigint | null
): Promise<{ hasNewMessage: boolean; contentUpdated: boolean; newFloor: bigint | null }> {
  const raw = await extractVisibleMessagesStable(page);
  // Attachment markers are resolved here rather than inside
  // extractVisibleMessages itself — deciding a real image's final URL
  // means checking for (and possibly saving) a local snapshot, which
  // needs DB access that the DOM-extraction layer deliberately doesn't
  // have. Mutates each raw message's content in place before it's turned
  // into an ObservedMessage.
  for (const msg of raw) {
    msg.content = resolveAttachmentMarkers(msg.content, msg.messageId, msg.attachments, msg.images);
  }
  const observed = toObservedMessages(conversationId, raw, new Date().toISOString());

  let hasNewMessage = false;
  let contentUpdated = false;
  let newFloor = floor;
  for (const msg of observed) {
    const owner = getMessageOwner(msg.messageId);
    if (owner !== null && owner !== conversationId) continue; // belongs to a different conversation — never trust it

    const id = toBigIntOrNull(msg.messageId);
    const isNew = id !== null && (floor === null || id > floor);

    const rev = recordObservedMessage(msg);
    if (rev) {
      contentUpdated = true;
      if (isNew) hasNewMessage = true;
    }

    if (id !== null && (newFloor === null || id > newFloor)) newFloor = id;
  }

  if (newFloor !== floor) advanceWatermark(conversationId, newFloor!.toString(), new Date().toISOString());
  touchSynced(conversationId);
  return { hasNewMessage, contentUpdated, newFloor };
}

/**
 * Tracks a conversation if it isn't already (idempotent) — baseline
 * watermark only, no backfill — regardless of whether it got here via
 * auto-discovery or a manual track request.
 */
async function autoTrackConversation(page: Page, discovered: DiscoveredConversation): Promise<void> {
  if (getTrackedConversation(discovered.id)) return;

  await openConversation(page, discovered.id);
  startTracking(discovered.id, discovered.title, discovered.kind);

  const raw = await extractVisibleMessagesStable(page);
  const newest = raw.at(-1);
  const now = new Date().toISOString();
  if (newest) advanceWatermark(discovered.id, newest.messageId, now);
  touchSynced(discovered.id);
}

/** One discovery pass: finds every conversation in the left rail and auto-tracks any that aren't tracked yet. */
async function runDiscoveryScan(page: Page): Promise<void> {
  const discovered = await withPage(async () => {
    // Confirmed live: right after startup, the left rail can still be
    // empty in the DOM for a moment even though the page itself has
    // loaded — scanning too early silently finds 0 conversations. Wait
    // for at least one real entry before trusting the scan.
    await page.locator(selectors.conversationListItem).first().waitFor({ state: "visible", timeout: 30_000 });
    return discoverConversations(page);
  });

  let newlyTracked = 0;
  for (const d of discovered) {
    if (d.kind !== "chat") continue; // channels intentionally out of scope for now
    if (getTrackedConversation(d.id)) continue;
    await withPage(() => autoTrackConversation(page, d));
    newlyTracked++;
  }
  console.log(`[discovery] scan found ${discovered.length} conversations, auto-tracked ${newlyTracked} new`);
}

/**
 * Runs discovery once immediately, then on config.discoveryIntervalMs —
 * covers "everything now" and "anything new later". Returns a function to
 * stop the recurring scan (used when switching to a headed browser for
 * re-auth, so a scan doesn't fire against a page that's about to close).
 */
export async function startDiscoveryLoop(page: Page): Promise<() => void> {
  await runDiscoveryScan(page).catch((err) => console.error("[watcher] initial discovery scan failed:", err));
  let stopped = false;
  const timer = setInterval(() => {
    if (stopped) return;
    runDiscoveryScan(page).catch((err) => console.error("[watcher] discovery scan failed:", err));
  }, config.discoveryIntervalMs);
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
 * This used to call an internal REST API (chatsvc.ts) instead of opening
 * each conversation, on the theory that one authenticated HTTP call would
 * be cheaper than a DOM visit. Confirmed live that it doesn't work: that
 * endpoint 401s on every call ("Authentication failed") because it needs
 * an explicit Skype-protocol auth token the current Teams web client
 * doesn't expose anywhere page.request or an in-page fetch can pick up
 * (it lives encrypted in localStorage, and — also confirmed live — real
 * Teams navigation never calls this endpoint itself anymore, so there's
 * no live request to harvest a working header from either). It was
 * failing silently: no error, just zero messages returned, every time,
 * regardless of how long the gap was — which is exactly the bug this was
 * written to fix.
 *
 * Reuses the same open-and-extract path as the backup visit loop
 * (visitConversation) instead — DOM-based, already proven to work, no
 * separate auth path to keep working. Goes through withPage since it's
 * now touching the real page, same as any other visit.
 */
export async function catchUpAllTracked(
  page: Page,
  notify: ChangeListener,
  onSyncStatus?: SyncListener,
  onMessagePending?: MessagePendingListener
): Promise<void> {
  for (const conv of listTrackedConversations()) {
    await withPage(() => visitConversation(page, conv, notify, onSyncStatus, onMessagePending)).catch((err) =>
      console.error(`[watcher] catch-up failed for "${conv.title}":`, err)
    );
  }
}

/**
 * Runs the REST fast path's token warm-up (see warmChatServiceToken in
 * teams.ts) through withPage, same as any other DOM-touching operation —
 * meant to be called once, early at startup, so a real send later doesn't
 * have to be the one paying for it.
 */
export async function warmSendPath(page: Page): Promise<void> {
  const t0 = Date.now();
  await withPage(() => warmChatServiceToken(page));
  console.log(`[send] warm-up took ${Date.now() - t0}ms`);
}

/**
 * Sends a message from the dashboard's chatbox.
 *
 * Tries the REST fast path first (see sendMessageViaRest in teams.ts) —
 * outside withPage's queue entirely, since it never touches the page's
 * DOM/navigation state and so can't collide with an in-flight visit. Only
 * on that failing (no cached token, expired, network error) does this fall
 * back to the DOM-typing path, which DOES need withPage: it can't run
 * concurrently with a scheduled visit or catch-up without risking typing
 * into whichever conversation happened to be open for a different reason.
 *
 * Deliberately does NOT record the sent message itself into the DB here,
 * on either path: Teams' own real-time push (trouter.ts) will deliver it
 * back as a genuine "new message" event within moments, same as anything
 * anyone else sends, which keeps this on the one code path that already
 * handles ownership checks, watermark advancement, and dashboard
 * notification correctly — duplicating that logic here would just be two
 * places that could drift out of sync with each other.
 */
export async function sendMessageTo(page: Page, conversationId: string, text: string): Promise<void> {
  const t0 = Date.now();
  if (await sendMessageViaRest(page, conversationId, text)) {
    console.log(`[send] REST fast path for "${conversationId}" took ${Date.now() - t0}ms`);
    return;
  }
  await withPage(() => sendMessageInPage(page, conversationId, text));
  console.log(`[send] DOM fallback for "${conversationId}" took ${Date.now() - t0}ms total`);
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
  return (hash % 1000) / 1000 * intervalMs * 0.3;
}

function intervalFor(conv: TrackedConversation): number {
  return config.pool.backupCheckIntervalMs + jitterFor(conv.conversation_id, config.pool.backupCheckIntervalMs);
}

function isDue(conv: TrackedConversation, now: number): boolean {
  if (!conv.last_synced_at) return true;
  return now - new Date(conv.last_synced_at).getTime() >= intervalFor(conv);
}

/** One scheduled visit: opens the conversation and checks for anything new (push already handles the fast path — this is just the backup). */
async function visitConversation(
  page: Page,
  conv: TrackedConversation,
  notify: ChangeListener,
  onSyncStatus?: SyncListener,
  onMessagePending?: MessagePendingListener
): Promise<void> {
  onSyncStatus?.(conv.conversation_id, true);
  try {
    await openConversation(page, conv.conversation_id);
    const floor = toBigIntOrNull(conv.last_watermark_message_id);
    const { hasNewMessage, contentUpdated } = await extractAndRecordNew(page, conv.conversation_id, floor);

    touchSynced(conv.conversation_id);
    if (hasNewMessage) {
      touchActivity(conv.conversation_id, new Date().toISOString());
    }
    if (contentUpdated) {
      notify(conv.conversation_id);
    }
    // Cleared after any real visit, successful or not — a message still
    // empty after this either resolved to genuinely nothing or has
    // already scrolled out of the currently-rendered window, and either
    // way retrying it forever wouldn't change that.
    pendingResolutionConversations.delete(conv.conversation_id);

    // Same "give up after any real attempt" treatment for the per-message
    // signal — a real DOM pass just gave a definitive, current answer for
    // every message flagged pending in this conversation, whether or not
    // this particular pass actually changed anything (it may already have
    // been correct), so tell the dashboard to stop showing a spinner on
    // all of them now rather than leave it spinning forever.
    const pendingMessages = pendingResolutionMessagesByConversation.get(conv.conversation_id);
    if (pendingMessages) {
      pendingResolutionMessagesByConversation.delete(conv.conversation_id);
      for (const messageId of pendingMessages) onMessagePending?.(messageId, false);
    }
  } finally {
    onSyncStatus?.(conv.conversation_id, false);
  }
}

/**
 * Runs forever: repeatedly visits whichever tracked conversation is most
 * overdue for a check, sleeping briefly when nothing is due yet. Every
 * conversation is treated uniformly — oldest last_synced_at goes first —
 * EXCEPT pendingResolutionConversations, checked first every tick: a
 * push-detected likely-unresolved attachment jumps straight to the front
 * of the queue rather than waiting for its normal turn, which — for a
 * busy conversation — touchSynced on every push can defer indefinitely
 * (see that Set's own comment for why). Returns a function to stop
 * scheduling further visits (used when switching to a headed browser for
 * re-auth).
 */
export function startVisitLoop(
  page: Page,
  notify: ChangeListener,
  onSyncStatus?: SyncListener,
  onMessagePending?: MessagePendingListener
): () => void {
  let stopped = false;
  let scheduled: ReturnType<typeof setTimeout> | null = null;

  // Cancellable, so wakeVisitLoop can cut short whatever wait tick() is
  // currently in (most importantly the 2s idle-poll below) instead of
  // that wait running out on its own first — the difference between a
  // pending image resolving "instantly" versus up to 2s later.
  function scheduleTick(delayMs: number) {
    if (scheduled) clearTimeout(scheduled);
    scheduled = setTimeout(tick, delayMs);
  }
  wakeVisitLoop = () => scheduleTick(0);

  async function tick() {
    if (stopped) return;

    let conv: TrackedConversation | undefined;
    const pendingId = pendingResolutionConversations.values().next().value;
    if (pendingId) {
      conv = getTrackedConversation(pendingId);
      if (!conv) pendingResolutionConversations.delete(pendingId); // stale — e.g. untracked since
    }

    if (!conv) {
      const now = Date.now();
      const due = listTrackedConversations()
        .filter((c) => isDue(c, now))
        .sort((a, b) => (a.last_synced_at ?? "").localeCompare(b.last_synced_at ?? ""));

      if (due.length === 0) {
        scheduleTick(2000);
        return;
      }
      conv = due[0];
    }

    let failed = false;
    try {
      await withPage(() => visitConversation(page, conv!, notify, onSyncStatus, onMessagePending));
    } catch (err) {
      failed = true;
      console.error(`[watcher] visit failed for "${conv.title}":`, err);
      // Don't keep retrying a broken conversation ahead of everything
      // else every single tick — same backoff-by-falling-back-to-normal-
      // rotation treatment as any other visit failure. Message-level
      // pending state gets the same treatment as visitConversation's own
      // success path — a repeatedly-failing conversation shouldn't leave
      // the dashboard showing a spinner forever on a message this loop
      // can't actually reach.
      pendingResolutionConversations.delete(conv.conversation_id);
      const pendingMessages = pendingResolutionMessagesByConversation.get(conv.conversation_id);
      if (pendingMessages) {
        pendingResolutionMessagesByConversation.delete(conv.conversation_id);
        for (const messageId of pendingMessages) onMessagePending?.(messageId, false);
      }
    }
    if (stopped) return;
    // A failure (e.g. the browser itself became unavailable) would
    // otherwise retry instantly forever, busy-looping and flooding the
    // log — back off briefly instead of using setImmediate.
    scheduleTick(failed ? 3000 : 0);
  }
  tick();
  return () => {
    stopped = true;
    wakeVisitLoop = null;
    if (scheduled) clearTimeout(scheduled);
  };
}
