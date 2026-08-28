# Teams Message Watcher

Local tool that auto-discovers and tracks every Microsoft Teams chat you have,
detects new messages, edits, and deletes in near real time, keeps a full
edit/delete history per message, and shows them in a local multi-tab dashboard
— without opening Teams itself. No Microsoft Graph API, no app registration:
it drives a real, already-logged-in Edge session on a single shared tab, and
listens to the same real-time channel Teams' own web client uses.

Response generation is intentionally out of scope for now (see design notes) —
this build only listens and records.

## If Teams' UI changes: re-verify the selectors

`src/watcher/selectors.ts` was verified live against a real Teams session
(2026-08-22) — see the comments at the top of that file for what was and
wasn't confirmed. Teams' markup can still change across client updates. If
discovery/tracking stops finding conversations or messages, that file is the
one to check and update against the current DOM in DevTools.

## Setup

```
npm install
npx playwright install msedge
npm run dev
```

On first run, a dedicated Edge window opens (profile stored in
`edge-profile/teams-watcher/`, separate from any other automation profile you use
elsewhere — see design note below). Log into Teams manually in that window. The
session persists for subsequent runs.

Then open the dashboard at http://localhost:4173 — every conversation you have
gets discovered and tracked automatically within a few minutes of startup (and
any new one shows up within `config.discoveryIntervalMs`, default 5 min).
There's no manual "track" step — tracking is 100% automatic, nothing to opt
into. Pin a conversation to keep it at the top of the list; unpinned
conversations sort by most recent activity, and any with unseen messages are
highlighted until you open them.

## Design notes

- **Dedicated Edge profile**: Chromium locks a profile's user-data-dir to one
  running browser process, so this can't share a profile with another
  long-running automation (e.g. a separate PreQA browser tool) without the two
  blocking each other. Hence its own profile.
- **Watermark, not full history**: tracking a conversation baselines to its
  newest message at that moment — no backfill to the start of the conversation.
- **Catch-up on restart, via DOM, not REST**: on startup, `catchUpAllTracked`
  (`src/watcher/index.ts`) opens every tracked conversation in turn and
  extracts whatever's currently rendered, same as the backup visit loop
  (`visitConversation`) — reusing that same proven DOM-extraction path rather
  than duplicating it. An earlier version called Teams' internal `chatsvc`
  REST API directly with `page.request`, on the theory that riding on the
  same authenticated browser context would avoid needing separate auth
  handling. Confirmed live that it doesn't: that endpoint 401s
  ("Authentication failed") on every call, silently returning zero messages
  regardless of gap length — it needs an explicit Skype-protocol auth token
  that lives encrypted in localStorage, which neither `page.request` nor an
  in-page `fetch` can pick up, and real Teams navigation doesn't call this
  endpoint itself anymore either (confirmed by capturing live traffic — no
  request to harvest a working header from). Catch-up after any outage
  silently did nothing under that design; it's DOM-based now, same as
  everything else.
- **Edit/delete history**: messages are never overwritten. Every observed
  change in content (or a flip to deleted) appends a new row to
  `message_revisions`; `messages` just tracks first/last seen. Click any
  message in the dashboard to see its full revision history.
- **One single tab for everything**: login, discovery, tracking, and every
  ongoing check all happen on exactly one Playwright `Page`. There's no real
  deep-link URL to navigate to directly — this Teams client never reflects
  the open conversation in the address bar (confirmed live: it stays at the
  app root regardless of what's open) — so every visit loads the app shell
  once and clicks through to the target conversation by id.
- **Push-based real-time detection, not DOM polling**: `src/watcher/trouter.ts`
  attaches a listener to the page's own WebSocket connection to Teams'
  "Trouter" real-time transport (`wss://*.trouter.teams.microsoft.com`) — the
  same channel the Teams web client itself uses to stay live-updated. Every
  frame is parsed for `resourceType: "NewMessage"` (new message) and
  `"MessageUpdate"` (edit, or delete when `properties.deletetime` is set).
  This is the primary detection path and needs no page visit or DOM read at
  all — a message shows up the moment Teams' own client would show it.
  `startVisitLoop` still exists underneath as a slow backup (uniform
  `config.pool.backupCheckIntervalMs`, default 5 min, across every tracked
  conversation, no priority tiers) in case a push is ever missed — e.g. a
  brief disconnect — but it's no longer the thing responsible for real-time
  behavior. An earlier design had conversations promote/demote between
  frequent-poll ("live") and slow-poll ("polling") tiers to approximate
  real-time behavior through DOM polling alone; that whole distinction is
  gone now that push actually delivers real-time updates, and polling only
  needs to be a safety net.
  A version before that kept a dedicated tab open per live conversation plus
  a small pool of polling tabs — besides the resource cost of many open
  tabs, that design was also the source of a real cross-conversation
  message-leak bug: two tabs' DOMs were confirmed live to independently
  report identical message ids for what should have been two different
  conversations, and no amount of ownership-guarding or locking fully closed
  it. With a single tab, there's structurally no second tab for content to
  ever be confused with — simpler and more correct, not just lighter.
- **Headless by default, headed only for re-auth**: `src/index.ts` runs an
  infinite loop alternating between `runHeadlessSession` (the normal case —
  no window at all) and `runHeadedLoginSession`. Every page navigation in
  the headless session is watched (`src/watcher/auth.ts`'s `watchForReauth`)
  for a redirect to `login.microsoftonline.com` (or the classic AAD login
  form elements) — Teams' session cookie eventually expires or triggers an
  MFA challenge, and Playwright can't render an interactive login form with
  no window to show it in. When that happens, the headless session stops
  its push listener/discovery/visit loop and closes; a real, visible Edge
  window then opens against the same persistent profile just for you to log
  in, closes itself once you have, and the loop goes back to headless. The
  dashboard shows a banner for the duration via the same auth-status
  WebSocket message used before. This replaced an earlier design that kept
  one browser open the whole time, headed but minimized, and used Chrome
  DevTools Protocol window controls to restore/focus it for re-auth instead
  of a full relaunch — that worked, but left a visible (if minimized) Edge
  window running at all times, which defeats the point of asking for
  headless in the first place.
- **Image attachments: downloaded and stored locally, not linked**: a real
  image attachment renders as an `<img>` whose `src` is a page-local `blob:`
  URL (dead outside the tab that created it), and the one real stable URL
  Teams gives for it (`amspreviewurl`) 401s without an authenticated Teams
  session — confirmed live via a plain `curl`. Rather than embed a link that
  usually won't load for whoever's viewing the dashboard,
  `src/watcher/attachments.ts` reads the blob's actual bytes directly out of
  the page (the browser already fetched and decoded them to render the
  thumbnail, so no additional authenticated request is needed at all),
  saves them to `data/attachments/<sha256>.<ext>`, and records the mapping
  in the `attachment_snapshots` table. The dashboard's own
  `/api/attachments/:messageId/:name` route serves the file directly — an
  image just works for any viewer, with no Teams session of its own
  required. Non-image attachments (documents, spreadsheets) are unaffected
  — those still show as a `📎 filename` chip linking to the original
  (auth-gated) Teams/SharePoint URL, since there's no meaningful way to
  inline-preview a JSON file or spreadsheet anyway.
  A directly-pasted image (Ctrl+V into the compose box) turned out to be a
  THIRD, distinct shape from either of the above — confirmed live it isn't
  a `file-attachment-grid` sibling at all, but an `<img>` rendered INSIDE
  the message body itself (wrapped in `itemtype="http://schema.skype.com/
  AMSImage"`), with no visible text — which is exactly why a pasted image
  first showed up as a completely empty message, not even an
  `[[ATTACHMENT]]` marker. Same blob-read technique applies once found; it
  just has no real filename the way an uploaded file does, so one is
  synthesized (`pasted-image-N`) and given a real extension once the
  actual content type is known from the downloaded bytes.
