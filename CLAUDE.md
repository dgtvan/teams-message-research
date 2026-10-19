# Running/debugging this project

- **Don't pop a browser tab during debugging**: `npm run dev` auto-opens the
  dashboard in the OS default browser on startup (`openInBrowser` in
  `src/server/index.ts`) — fine for the user starting it normally, but
  disruptive if you're restarting it repeatedly while debugging (each
  restart pops a fresh tab in whatever browser the user is actually working
  in). Set `WATCHER_NO_OPEN_BROWSER=1` to suppress it:
  ```
  WATCHER_NO_OPEN_BROWSER=1 npm run dev
  ```

- **Prefer an isolated debugging session over sharing the user's live one.**
  The user may be running their own instance at the same time you're
  debugging — don't collide with it, and don't debug against their real
  data if you don't have to. Clone it and use a different port:
  ```
  npm run clone-debug-data
  WATCHER_DATA_DIR=data/debug WATCHER_PORT=4174 WATCHER_NO_OPEN_BROWSER=1 npm run dev
  ```
  `clone-debug-data` (`scripts/clone-debug-data.mjs`) snapshots
  `data/teams.db*` and `data/attachments/` into `data/debug/` — re-run it
  whenever you want a fresher snapshot, it's cheap. `WATCHER_DATA_DIR`
  points `dbPath`/`attachmentsDir` at that clone instead of the live data
  (see `src/config.ts`); `WATCHER_PORT` picks a different port so the two
  dashboard servers don't fight over 4173. This is enough isolation for
  anything that only needs the dashboard/API/DB — rendering changes,
  extraction-logic changes tested by re-reading already-captured messages,
  etc. — which covers most debugging.

  **This does NOT isolate the browser** — `edgeProfileDir` (the one
  logged-in Teams session) is a single shared resource on purpose, see the
  comment in `src/config.ts` for why there's deliberately no
  `WATCHER_PROFILE_DIR` env var for a second profile. But live DOM/browser
  access doesn't need a second profile either — it needs a second TAB in
  the browser the watcher already has open, which is a completely
  different thing from a second browser PROCESS and doesn't touch the
  lock at all. Restart the watcher once with a CDP debug port open:
  ```
  WATCHER_NO_OPEN_BROWSER=1 WATCHER_CDP_PORT=9333 npm run dev
  ```
  then from a separate script, attach to that SAME running instance
  instead of launching a new one:
  ```js
  import { chromium } from "playwright";
  const browser = await chromium.connectOverCDP("http://localhost:9333");
  const context = browser.contexts()[0]; // the watcher's own context
  const page = await context.newPage();  // a new tab, alongside its own — no lock conflict
  // ... use page, then page.close() when done ...
  ```
  Do NOT call `browser.close()` at the end — for a CDP-attached connection
  that's supposed to just disconnect rather than close the real browser,
  but getting that wrong would kill the user's actual live session, and
  it's not worth the risk of being wrong about it. Just `page.close()` and
  let the script exit; the remote browser doesn't belong to your script's
  process, so it isn't affected either way. This is how the "instant
  image resolution" timing tests actually got measured live, without ever
  pausing the user's own running session — confirmed working, including
  confirming the watcher's own tab and its push listener/visit loop kept
  running completely undisturbed throughout.

  `WATCHER_CDP_PORT` is unset by default (no open debug port on a normal
  run) — only set it for a debugging session, and only on an instance you
  started for that purpose.

- **The `teams-bot` connector needs none of the browser caution above.**
  `WATCHER_CONNECTOR=teams-bot` drives no browser at all, so there's no
  profile lock to collide with and no reason to avoid restarting it. It
  binds a local port (`WATCHER_BOT_PORT`, default 3978) instead, so that
  is the thing two sessions can fight over. With no
  `MICROSOFT_APP_ID`/`MICROSOFT_APP_PASSWORD` set it runs unauthenticated
  on loopback, which is what makes it fully testable offline:
  ```
  WATCHER_CONNECTOR=teams-bot WATCHER_DATA_DIR=data/bot-test WATCHER_PORT=4175 WATCHER_NO_OPEN_BROWSER=1 npm run dev
  npm run bot-harness   # posts synthetic Teams activities at it (separate terminal)
  ```
  Prefer this for anything that isn't specifically about the browser path
  — it exercises ingest, revisions, tracking and the dashboard end to end
  with no Teams account involved. Never expose it through a tunnel while
  unauthenticated.

- **Check what's already running before starting anything**
  (`tasklist //FI "IMAGENAME eq node.exe"` on Windows). Even with the
  isolation above, starting a browser-touching session while the user's is
  up will fail on the profile lock, not a bug in the app.

- **Stop your own session when you're done with it** — don't leave a
  debugging instance running across turns. Leaving one running caused real
  confusion once already: the user tried to start their own copy later and
  couldn't, because a session left running from a prior conversation was
  still holding the port/profile lock.

- **Self-healing**: the backup visit loop re-extracts every currently
  visible message on every visit, not just new ones (see
  the visit loop in `src/connectors/teams-web/index.ts`), and
  `recordObservedMessage` only writes a new revision if content actually
  changed. So a fix to extraction logic doesn't need a manual backfill —
  restarting the watcher and letting it revisit conversations is usually
  enough to retroactively correct already-captured messages, as long as
  they're still within Teams' currently-rendered (non-scrolled) window.
