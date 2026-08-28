import { config } from "./config.js";
import { launchTeamsBrowser } from "./watcher/browser.js";
import {
  catchUpAllTracked,
  sendMessageTo,
  startDiscoveryLoop,
  startPushListener,
  startVisitLoop,
} from "./watcher/index.js";
import { isOnLoginPage, waitUntilLoggedIn, watchForReauth } from "./watcher/auth.js";
import { startServer } from "./server/index.js";

type NotifyUpdate = (conversationId: string) => void;
type NotifyAuth = (status: "required" | "resolved") => void;
type NotifySync = (conversationId: string, syncing: boolean) => void;
type SetSendHandler = (handler: ((conversationId: string, text: string) => Promise<void>) | null) => void;

// Tracked so a shutdown signal can close whichever browser context is
// currently open instead of leaving an orphaned Edge process holding the
// profile lock. Covers Ctrl+C (SIGINT) and a graceful stop from a process
// manager/supervisor (SIGTERM) — a hard kill of this process instead
// (Task Manager "End task", taskkill /F) skips this entirely, but on
// Windows Playwright launches the browser inside a Job Object tied to
// this process's lifetime, so the whole Edge tree still gets torn down
// automatically even then.
let currentContext: { close: () => Promise<void> } | null = null;
async function shutdown() {
  await currentContext?.close().catch(() => {});
  process.exit(0);
}
process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);

/**
 * Runs one headless session: launches, and if already logged in (the
 * common case — the persistent profile carries the session across
 * restarts), wires up push detection + catch-up + discovery + the backup
 * visit loop and runs until Teams redirects to a login page. Resolves
 * once the session has torn itself down, either because a login page
 * showed up (needs a headed browser you can't get headless) or the
 * session was already logged out when it launched.
 */
async function runHeadlessSession(
  notifyUpdate: NotifyUpdate,
  notifyAuth: NotifyAuth,
  notifySync: NotifySync,
  setSendHandler: SetSendHandler
): Promise<void> {
  const context = await launchTeamsBrowser(true);
  currentContext = context;
  const [page] = context.pages();
  await page.goto(config.teamsUrl, { waitUntil: "domcontentloaded" });

  if (await isOnLoginPage(page)) {
    await context.close();
    return;
  }

  console.log("[watcher] Headless session active.");
  setSendHandler((conversationId, text) => sendMessageTo(page, conversationId, text));
  await new Promise<void>((resolve) => {
    let stopDiscovery = () => {};
    const stopVisits = startVisitLoop(page, notifyUpdate, notifySync);

    const stopWatching = watchForReauth(page, () => {
      console.log("[watcher] Teams needs you to log in again — switching to a visible browser.");
      notifyAuth("required");
      setSendHandler(null); // this page is about to close — nothing left to send through
      stopDiscovery();
      stopVisits();
      stopWatching();
      context.close().finally(resolve);
    });

    startPushListener(page, notifyUpdate);
    catchUpAllTracked(page, notifyUpdate, notifySync)
      .then(() => startDiscoveryLoop(page))
      .then((stop) => {
        stopDiscovery = stop;
      })
      .catch((err) => console.error("[watcher] startup catch-up/discovery failed:", err));
  });
}

/** Opens a real, visible window just for you to sign in, then tears it down once you have. */
async function runHeadedLoginSession(notifyAuth: NotifyAuth): Promise<void> {
  const context = await launchTeamsBrowser(false);
  currentContext = context;
  const [page] = context.pages();
  await page.goto(config.teamsUrl, { waitUntil: "domcontentloaded" });
  console.log("[watcher] A visible Edge window is open — please log in there.");
  await waitUntilLoggedIn(page);
  console.log("[watcher] Login resolved — switching back to headless.");
  notifyAuth("resolved");
  await context.close();
}

async function main() {
  const server = startServer();
  for (;;) {
    await runHeadlessSession(
      server.notifyConversationUpdate,
      server.notifyAuth,
      server.notifySyncStatus,
      server.setSendHandler
    );
    await runHeadedLoginSession(server.notifyAuth);
  }
}

main().catch((err) => {
  console.error("[fatal]", err);
  process.exit(1);
});
