// Frees the dashboard's port before starting, by killing whatever process
// is already listening on it. Wired in as npm's predev/prestart hook (see
// package.json) so a stale watcher instance left running from a prior
// session — confirmed to happen: stopping a background task doesn't always
// kill the underlying npm/tsx process tree on Windows — no longer silently
// blocks the next `npm run dev`/`npm start`. That stale process also blocks
// Teams login, since edgeProfileDir (config.ts) is a single shared browser
// profile that only the stale process tree is holding onto.
//
// Kills the whole process TREE, not just the listening PID: `npm run dev`
// is nested three deep (npm-cli.js -> tsx/dist/cli.mjs -> node w/ the tsx
// loader) and Playwright launches the browser as a further child of that —
// killing only the top PID can leave siblings/children (and the browser)
// running.
//
// Best-effort only: any failure here just logs and falls through, since
// this runs as a pre-hook and must never be the reason `npm run dev` itself
// refuses to start.
import { execSync } from "node:child_process";

const port = Number(process.env.WATCHER_PORT) || 4173;

function findWindowsPids() {
  const out = execSync(
    `powershell -NoProfile -Command "(Get-NetTCPConnection -LocalPort ${port} -State Listen -ErrorAction SilentlyContinue).OwningProcess | Sort-Object -Unique"`,
    { encoding: "utf8" }
  );
  return out
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map(Number);
}

function findPosixPids() {
  try {
    const out = execSync(`lsof -ti tcp:${port}`, { encoding: "utf8" });
    return out
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
      .map(Number);
  } catch {
    return []; // lsof exits non-zero when nothing matches the port
  }
}

function killWindows(pid) {
  execSync(`taskkill /PID ${pid} /F /T`, { stdio: "ignore" });
}

try {
  const pids = process.platform === "win32" ? findWindowsPids() : findPosixPids();

  if (!pids.length) {
    console.log(`[free-port] Port ${port} is already free.`);
  } else {
    for (const pid of pids) {
      try {
        if (process.platform === "win32") killWindows(pid);
        else process.kill(pid, "SIGKILL");
        console.log(`[free-port] Killed stale process ${pid} that was holding port ${port}.`);
      } catch (err) {
        console.warn(`[free-port] Could not kill process ${pid}: ${err.message}`);
      }
    }
  }
} catch (err) {
  console.warn(`[free-port] Skipped (couldn't check port ${port}): ${err.message}`);
}
