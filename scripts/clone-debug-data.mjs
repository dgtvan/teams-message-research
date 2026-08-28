// Snapshots the real data/ directory into data/debug/, for a debugging
// session to point at via WATCHER_DATA_DIR (see README's "Debugging
// without disrupting a live session") instead of touching the live DB
// while the user might be running their own instance against it.
//
// Copies teams.db-wal/-shm alongside teams.db (not just the main file) —
// the DB runs in WAL mode, so recent writes can still be sitting in the
// -wal file rather than checkpointed into teams.db yet; copying only the
// main file could silently snapshot a stale state.
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const srcDir = path.join(projectRoot, "data");
const destDir = path.join(projectRoot, "data", "debug");

if (!fs.existsSync(path.join(srcDir, "teams.db"))) {
  console.error(`No database found at ${path.join(srcDir, "teams.db")} — nothing to clone yet.`);
  process.exit(1);
}

fs.rmSync(destDir, { recursive: true, force: true });
fs.mkdirSync(destDir, { recursive: true });

for (const dbFile of ["teams.db", "teams.db-wal", "teams.db-shm"]) {
  const src = path.join(srcDir, dbFile);
  if (fs.existsSync(src)) fs.copyFileSync(src, path.join(destDir, dbFile));
}

const attachmentsSrc = path.join(srcDir, "attachments");
if (fs.existsSync(attachmentsSrc)) {
  fs.cpSync(attachmentsSrc, path.join(destDir, "attachments"), { recursive: true });
}

console.log(`Cloned into ${destDir}`);
console.log(`Run a debugging session against it with:`);
console.log(`  WATCHER_DATA_DIR=${path.relative(projectRoot, destDir)} WATCHER_PORT=4174 WATCHER_NO_OPEN_BROWSER=1 npm run dev`);
