import type { Frame, Page } from "playwright";

/**
 * Detects when Teams has dropped back to an AAD login/interactive-auth
 * page (session expired, MFA challenge, etc.). The watcher runs headless
 * by default, which has no window to show you a login form with — so
 * detecting this is what tells index.ts to tear the headless browser down
 * and relaunch headed just for you to sign in.
 */

const LOGIN_HOSTNAME_PATTERN = /login\.microsoftonline\.com|login\.live\.com/i;

// Long-lived AAD form element ids/names — these have been stable across
// years of login.microsoftonline.com revisions, so they're a solid
// backup signal alongside the hostname check.
const LOGIN_DOM_SELECTORS = [
  "#i0116", // email/UPN input
  "#i0118", // password input
  "#idSIButton9", // "Sign in" / "Next" submit
  'input[name="loginfmt"]',
].join(", ");

export async function isOnLoginPage(page: Page): Promise<boolean> {
  try {
    const url = new URL(page.url());
    if (LOGIN_HOSTNAME_PATTERN.test(url.hostname)) return true;
  } catch {
    // page.url() can be "about:blank" etc. during navigation; ignore.
  }
  try {
    const el = await page.$(LOGIN_DOM_SELECTORS);
    return el !== null;
  } catch {
    return false;
  }
}

/** Blocks until the page is no longer on a login page — used while a headed window is up for you to sign in. */
export async function waitUntilLoggedIn(page: Page, pollMs = 2000): Promise<void> {
  while (await isOnLoginPage(page)) {
    await page.waitForTimeout(pollMs);
  }
}

/**
 * Watches a page's navigations for a redirect to a login page, firing
 * onReauthNeeded once (a headless session only ever needs to react to
 * this once — index.ts tears the whole session down in response). Returns
 * a function to stop watching.
 */
export function watchForReauth(page: Page, onReauthNeeded: () => void): () => void {
  let fired = false;
  const handler = (frame: Frame) => {
    if (frame !== page.mainFrame() || fired) return;
    isOnLoginPage(page)
      .then((onLogin) => {
        if (onLogin && !fired) {
          fired = true;
          onReauthNeeded();
        }
      })
      .catch(() => {});
  };
  page.on("framenavigated", handler);
  return () => page.off("framenavigated", handler);
}
