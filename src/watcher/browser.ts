import { chromium, type BrowserContext } from "playwright";
import fs from "node:fs";
import { config } from "../config.js";

/**
 * Launches a persistent Edge context under the dedicated "teams-watcher"
 * profile. First run: log into Teams manually in the window that opens
 * (see README) — the session persists in edge-profile/teams-watcher for
 * subsequent runs.
 *
 * `headless` is per-launch, not a fixed setting: normal operation runs
 * headless (nothing to show), but Playwright can't render an interactive
 * login form with no window at all, so index.ts relaunches this headed
 * whenever Teams drops back to a login page, then headless again once
 * you're signed back in. Same profile dir either way, so the session
 * carries over across the switch.
 */
export async function launchTeamsBrowser(headless: boolean): Promise<BrowserContext> {
  fs.mkdirSync(config.edgeProfileDir, { recursive: true });

  const context = await chromium.launchPersistentContext(config.edgeProfileDir, {
    channel: config.edgeChannel,
    headless,
    viewport: { width: 1400, height: 900 },
    // See config.ts's debugCdpPort comment — unset on a normal run, so
    // this doesn't add an args: [] entry (and therefore no open debug
    // port) unless a debugging session explicitly asked for one.
    args: config.debugCdpPort ? [`--remote-debugging-port=${config.debugCdpPort}`] : [],
  });

  return context;
}
