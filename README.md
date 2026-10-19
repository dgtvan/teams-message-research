# Teams Message Watcher

Local tool that tracks Microsoft Teams conversations, detects new messages,
edits, and deletes in near real time, keeps a full edit/delete history per
message, and shows them in a local multi-tab dashboard — without opening Teams
itself.

How it reaches Teams is pluggable. There are two **connectors**, and exactly
one runs at a time (`WATCHER_CONNECTOR`, default `teams-web`):

| | **`teams-web`** (default) | **`teams-bot`** |
|---|---|---|
| mechanism | drives a real logged-in Edge session, listens to the same real-time channel Teams' own web client uses | a Bot Framework bot you add to groups/channels on demand |
| setup | log in once, in a dedicated Edge profile | Azure Bot registration + dev tunnel + sideloaded app |
| sees | **every** chat on your account, discovered automatically | **only** groups/channels the bot is a member of |
| history | can re-read the currently-rendered window, so it self-heals and catches up after downtime | **none** — messages sent while this app is down are lost for good |
| attachments | real image bytes, saved locally and previewed inline | links only (fetching them needs the bot's token) |
| sends as | you | the bot |

`teams-web` is the more capable one and stays the default. `teams-bot` trades
coverage and history for not driving a browser at all, and for being able to
post as a distinct identity. Both write through the same ingest into the same
DB, so switching between them keeps everything already captured.

## If Teams' UI changes: re-verify the selectors

Applies to the `teams-web` connector only — `teams-bot` reads a documented JSON
payload and doesn't care what the UI looks like.

`src/connectors/teams-web/selectors.ts` was verified live against a real Teams session
(2026-08-22) — see the comments at the top of that file for what was and
wasn't confirmed. Teams' markup can still change across client updates. If
discovery/tracking stops finding conversations or messages, that file is the
one to check and update against the current DOM in DevTools.

## Setup (`teams-web`, the default)

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
any new one shows up within `config.teamsWeb.discoveryIntervalMs`, default 5 min).
There's no manual "track" step — tracking is 100% automatic, nothing to opt
into. Pin a conversation to keep it at the top of the list; unpinned
conversations sort by most recent activity, and any with unseen messages are
highlighted until you open them.

## Setup (`teams-bot`)

Unlike `teams-web`, this connector can't work from local state alone: Teams
pushes activities *to* it, so it needs a publicly reachable HTTPS endpoint and
an Azure Bot registration. A bot cannot be added to a Teams conversation
without one.

### Try it with no Azure account at all

The whole inbound path — activity parsing, conversation tracking, revision
history, dashboard updates — can be exercised locally first, with a harness
that posts synthetic Teams activities:

```
# terminal 1 — note: no credentials set, so the endpoint runs unauthenticated
WATCHER_CONNECTOR=teams-bot WATCHER_DATA_DIR=data/bot-test WATCHER_PORT=4175 \
  WATCHER_NO_OPEN_BROWSER=1 npm run dev

# terminal 2
npm run bot-harness
```

That posts a bot-added event, plain and HTML-formatted messages, an edit, a
soft delete, a file attachment and a threaded channel reply, then tells you
what to expect in the dashboard at http://localhost:4175.

This works **only** because the Bot Framework SDK skips inbound JWT validation
when no credentials are configured. The connector warns loudly in that mode,
refuses to send, and binds to loopback only — **never put a tunnel in front of
it unauthenticated.**

### What you have to do yourself

Everything below needs your Azure/Teams account, and step 1 is the only part
that can cost money:

1. **Create an Azure Bot resource.** Azure portal → *Create a resource* → *Azure
   Bot*. Pick the **Free (F0)** pricing tier — it's sufficient for this and is
   the reason this step is called out separately. Choose *Multi Tenant* for app
   type unless you have a reason not to.
2. **Get the credentials.** On that resource: *Configuration* → copy the
   **Microsoft App ID**. Then *Manage* next to it → *Certificates & secrets* →
   *New client secret* → copy the **value** (shown once).
3. **Put them in `.env`.** `cp .env.example .env` and fill in
   `MICROSOFT_APP_ID` / `MICROSOFT_APP_PASSWORD`. `.env` is gitignored, and
   `npm run dev` loads it automatically.
4. **Enable the Teams channel.** On the Azure Bot resource: *Channels* → add
   **Microsoft Teams** → accept the terms.
5. **Start a tunnel** to the local listener and copy the public HTTPS URL:
   ```
   # either one works
   devtunnel host -p 3978 --allow-anonymous
   ngrok http 3978
   ```
6. **Set the messaging endpoint.** Back on *Configuration*, set it to
   `https://<your-tunnel-host>/api/messages` and save. This changes every time
   a free tunnel restarts, so expect to redo it.
7. **Build and sideload the Teams app:**
   ```
   npm run build-teams-app -- --bot-id <your-microsoft-app-id>
   ```
   That writes `teams-app/message-watcher.zip` (manifest + generated icons).
   Upload it in Teams → *Apps* → *Manage your apps* → *Upload a custom app*.
   Sideloading has to be permitted by your tenant's app setup policy; if the
   upload option isn't there, that's why.
8. **Run it and invite the bot:**
   ```
   WATCHER_CONNECTOR=teams-bot npm run dev
   ```
   Then add the bot to a group chat or channel. It starts recording from that
   moment — there's no backfill, because Bot Framework gives it no way to read
   history.

### Verifying each hop

Worth doing in order, since a failure at step 6 looks identical to a failure at
step 4 from inside Teams:

```
curl http://127.0.0.1:3978/api/messages            # is the connector listening?
curl https://<your-tunnel-host>/api/messages       # does the tunnel reach it?
```

Both should return `{"ok":true,...}`, with `"authenticated":true` once
credentials are set.

## Layout: connectors, core, server

The app is three layers, and the boundary between the first two is a real
interface rather than a convention:

```
src/connectors/*  →  src/core  →  src/db
                          ↓
                     src/server  →  public/   (dashboard)
```

- **`src/connectors/*` — how we talk to Teams, in both directions.** A
  *connector* owns a two-way link: observing messages coming out (new, edited,
  deleted) and sending messages back in. It also owns its own transport,
  authentication and scheduling. `teams-web` is the one that exists today
  (Playwright driving a real logged-in Edge profile) and is deliberately the
  **only** place in the repo that imports Playwright, holds Teams DOM
  selectors, or knows what Trouter is.
- **`src/core` — what a message means once we've seen it.** Ingest
  (watermarks, revision dedupe, ownership guarding, unread/activity
  bookkeeping), attachment snapshots and the `[[ATTACHMENT|…]]` / `[[QUOTE|…]]`
  marker format the dashboard renders. Connector-agnostic: nothing here knows
  or cares how an observation was made.
- **`src/server` + `public/`** — read the DB, serve the dashboard, broadcast
  live updates. Also connector-agnostic.

The contract between the first two is `Connector` / `ConnectorContext` /
`Ingest` in [`src/core/types.ts`](src/core/types.ts). A connector gets handed a
context, reports what it observes, and is asked to send; everything else is its
own business.

**Adding another connector** is therefore: a new directory under
`src/connectors/`, one line in the registry in `src/connectors/index.ts`, and
its own `config` block. Nothing in `core`, `db`, `server` or the dashboard
changes. Pick which one runs with `WATCHER_CONNECTOR` (default `teams-web`);
exactly one runs at a time, never both concurrently.

`teams-bot` was in fact built that way — it added `src/connectors/teams-bot/`,
one registry line and one config block, and touched nothing in `core`, `db` or
`server`. The one thing it *did* change is noted under `Connector.send` below.

Two things worth knowing when writing one:

- **Two ways to report messages, with different trust levels.**
  `ingest.incremental()` is for "this one message just changed" (a push, a
  change notification) and is treated as strictly additive — it will not
  overwrite anything at or below the conversation's watermark, because such a
  source is usually a lossier view of a message than a full re-read is.
  `ingest.snapshot()` is for "here is the current state of this conversation"
  and is authoritative and self-healing: it re-records everything it's given,
  so an improvement to a connector's own extraction retroactively corrects
  already-captured messages with no manual backfill. A connector may use
  either or both — `teams-web` uses both (push for speed, DOM re-reads as a
  backup and to fill in what push can't carry).
- **Message ids are expected to sort as integers.** The watermark ("is this
  message new?") compares them as `BigInt`, which is what Teams' own numeric
  message ids do — both connectors get such ids, `teams-web` from the DOM and
  `teams-bot` from `activity.id`. A connector whose ids don't have that
  property would need that comparison generalized.
- **`Connector.send` records its own message only if nothing else will.**
  `teams-web` doesn't, because Teams echoes a sent message back through its
  push path; `teams-bot` must, because Bot Framework never delivers a bot's own
  messages back to it. Check which applies before assuming either.

## Design notes

- **Dedicated Edge profile**: Chromium locks a profile's user-data-dir to one
  running browser process, so this can't share a profile with another
  long-running automation (e.g. a separate PreQA browser tool) without the two
  blocking each other. Hence its own profile.
- **Watermark, not full history**: tracking a conversation baselines to its
  newest message at that moment — no backfill to the start of the conversation.
- **Catch-up on restart, via DOM, not REST**: on startup, `catchUpAllTracked`
  (`src/connectors/teams-web/index.ts`) opens every tracked conversation in turn and
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
- **Push-based real-time detection, not DOM polling**: `src/connectors/teams-web/trouter.ts`
  attaches a listener to the page's own WebSocket connection to Teams'
  "Trouter" real-time transport (`wss://*.trouter.teams.microsoft.com`) — the
  same channel the Teams web client itself uses to stay live-updated. Every
  frame is parsed for `resourceType: "NewMessage"` (new message) and
  `"MessageUpdate"` (edit, or delete when `properties.deletetime` is set).
  This is the primary detection path and needs no page visit or DOM read at
  all — a message shows up the moment Teams' own client would show it.
  `startVisitLoop` still exists underneath as a slow backup (uniform
  `config.teamsWeb.backupCheckIntervalMs`, default 5 min, across every tracked
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
  the headless session is watched (`src/connectors/teams-web/auth.ts`'s `watchForReauth`)
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
  `src/core/attachments.ts` reads the blob's actual bytes directly out of
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
