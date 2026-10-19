import { DatabaseSync } from "node:sqlite";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { config } from "../config.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

export interface TrackedConversation {
  conversation_id: string;
  title: string;
  kind: "chat" | "channel";
  tracking_started_at: string;
  last_watermark_message_id: string | null;
  last_watermark_at: string | null;
  last_synced_at: string | null;
  last_activity_at: string | null;
  pinned: number;
  last_read_at: string | null;
}

export interface MessageRevision {
  revision_id: number;
  message_id: string;
  content: string;
  content_hash: string;
  edited_flag: number;
  is_deleted: number;
  captured_at: string;
}

export interface AttachmentSnapshot {
  message_id: string;
  name: string;
  content_hash: string;
  content_type: string;
  byte_size: number;
  captured_at: string;
}

export interface ObservedMessage {
  messageId: string;
  conversationId: string;
  sender: string | null;
  content: string;
  contentHash: string;
  editedFlag: boolean;
  isDeleted: boolean;
  observedAt: string;
}

fs.mkdirSync(path.dirname(config.dbPath), { recursive: true });
const db = new DatabaseSync(config.dbPath);
db.exec("PRAGMA journal_mode = WAL");
db.exec(fs.readFileSync(path.join(__dirname, "schema.sql"), "utf-8"));

// Lightweight migration for columns added after a DB already exists —
// schema.sql's CREATE TABLE IF NOT EXISTS only applies to a fresh DB.
function ensureColumn(table: string, column: string, ddl: string) {
  const cols = db.prepare(`PRAGMA table_info(${table})`).all() as Array<{ name: string }>;
  if (!cols.some((c) => c.name === column)) {
    db.exec(`ALTER TABLE ${table} ADD COLUMN ${ddl}`);
  }
}
ensureColumn("tracked_conversations", "pinned", "pinned INTEGER NOT NULL DEFAULT 0");
ensureColumn("tracked_conversations", "last_read_at", "last_read_at TEXT");

export function listTrackedConversations(): TrackedConversation[] {
  return db.prepare("SELECT * FROM tracked_conversations").all() as unknown as TrackedConversation[];
}

export function getTrackedConversation(conversationId: string): TrackedConversation | undefined {
  return db
    .prepare("SELECT * FROM tracked_conversations WHERE conversation_id = ?")
    .get(conversationId) as unknown as TrackedConversation | undefined;
}

/** Which conversation a message id is already attributed to, if any — see recordObservedMessage's ownership guard. */
export function getMessageOwner(messageId: string): string | null {
  const row = db.prepare("SELECT conversation_id FROM messages WHERE message_id = ?").get(messageId) as
    | { conversation_id: string }
    | undefined;
  return row?.conversation_id ?? null;
}

export function startTracking(conversationId: string, title: string, kind: "chat" | "channel") {
  const now = new Date().toISOString();
  db.prepare(
    `INSERT INTO tracked_conversations (conversation_id, title, kind, tracking_started_at)
     VALUES (?, ?, ?, ?)
     ON CONFLICT(conversation_id) DO NOTHING`
  ).run(conversationId, title, kind, now);
}

/** Informational only — "most recently active" for the dashboard, not a scheduling input. */
export function touchActivity(conversationId: string, at: string) {
  db.prepare(`UPDATE tracked_conversations SET last_activity_at = ? WHERE conversation_id = ?`).run(
    at,
    conversationId
  );
}

export function advanceWatermark(conversationId: string, messageId: string, at: string) {
  db.prepare(
    `UPDATE tracked_conversations
     SET last_watermark_message_id = ?, last_watermark_at = ?, last_synced_at = ?
     WHERE conversation_id = ?`
  ).run(messageId, at, new Date().toISOString(), conversationId);
}

export function touchSynced(conversationId: string) {
  db.prepare(`UPDATE tracked_conversations SET last_synced_at = ? WHERE conversation_id = ?`).run(
    new Date().toISOString(),
    conversationId
  );
}

export function setPinned(conversationId: string, pinned: boolean) {
  db.prepare(`UPDATE tracked_conversations SET pinned = ? WHERE conversation_id = ?`).run(
    pinned ? 1 : 0,
    conversationId
  );
}

/** Marks a conversation as read as of `at` — the dashboard calls this when you open it. */
export function markRead(conversationId: string, at: string) {
  db.prepare(`UPDATE tracked_conversations SET last_read_at = ? WHERE conversation_id = ?`).run(at, conversationId);
}

/**
 * Records an observed message snapshot. Inserts a new revision only if the
 * content (or deleted state) differs from the latest known revision for
 * that message_id — this is what preserves edit/delete history instead of
 * overwriting it.
 *
 * One deliberate, narrow exception: filling in a message's still-empty
 * first-ever revision (the push-only "pending" placeholder — see the
 * isFillingInPendingPlaceholder check below) updates that row in place
 * instead of inserting a second one, since that placeholder was never a
 * real state the message was actually in. Every other case is still
 * strictly append-only.
 *
 * Returns the inserted (or, for that one exception, updated) revision, or
 * null if nothing changed.
 */
export function recordObservedMessage(msg: ObservedMessage): MessageRevision | null {
  const existing = db
    .prepare("SELECT * FROM messages WHERE message_id = ?")
    .get(msg.messageId) as { message_id: string; conversation_id: string } | undefined;

  // A message belongs to exactly one conversation, permanently. If this id
  // is already attributed elsewhere, recording it here would corrupt that
  // invariant — reject rather than trust it, regardless of why the
  // mismatch happened (this was hit repeatedly in practice: Teams'
  // conversation-switch rendering is async in ways no amount of
  // client-side wait/stability heuristics fully closed).
  if (existing && existing.conversation_id !== msg.conversationId) {
    console.warn(
      `[db] refusing to record message ${msg.messageId} under ${msg.conversationId} — already owned by ${existing.conversation_id}`
    );
    return null;
  }

  if (!existing) {
    db.prepare(
      `INSERT INTO messages (message_id, conversation_id, sender, first_seen_at, last_seen_at, is_deleted)
       VALUES (?, ?, ?, ?, ?, ?)`
    ).run(msg.messageId, msg.conversationId, msg.sender, msg.observedAt, msg.observedAt, msg.isDeleted ? 1 : 0);
  } else {
    db.prepare(
      `UPDATE messages SET last_seen_at = ?, is_deleted = ? WHERE message_id = ?`
    ).run(msg.observedAt, msg.isDeleted ? 1 : 0, msg.messageId);
  }

  const revisionCount = (
    db.prepare(`SELECT COUNT(*) AS n FROM message_revisions WHERE message_id = ?`).get(msg.messageId) as {
      n: number;
    }
  ).n;
  const latestRevision =
    revisionCount > 0
      ? (db
          .prepare(`SELECT * FROM message_revisions WHERE message_id = ? ORDER BY revision_id DESC LIMIT 1`)
          .get(msg.messageId) as unknown as MessageRevision)
      : undefined;

  const changed =
    !latestRevision ||
    latestRevision.content_hash !== msg.contentHash ||
    !!latestRevision.is_deleted !== msg.isDeleted;

  if (!changed) return null;

  // Push can't extract attachments/mentions/quotes at all (see
  // connectors/teams-web/trouter.ts) — a brand-new message's very first revision is
  // sometimes just an empty placeholder, recorded so the dashboard has
  // something to attach a "fetching…" indicator to while push's own
  // blind spot gets resolved by the next DOM visit (see renderBody in
  // app.js). That placeholder was never a real state Teams actually
  // showed the message in, so filling it in once resolved should
  // complete that ONE revision rather than leave a "was empty" ghost
  // entry behind in the history modal forever — confirmed live: exactly
  // that showed up as two revisions for one message, one of them just
  // "Fetching message content…".
  //
  // Scoped tightly on purpose — only when this is the message's ONE AND
  // ONLY revision so far, and only when that revision was literally
  // empty — so this can't quietly swallow a genuine edit later in a
  // message's real history (e.g. a message that already had content,
  // then got edited down to something else, still gets a real second
  // revision as it always has).
  const isFillingInPendingPlaceholder =
    revisionCount === 1 && latestRevision?.content === "" && msg.content !== "";

  if (isFillingInPendingPlaceholder && latestRevision) {
    db.prepare(
      `UPDATE message_revisions SET content = ?, content_hash = ?, edited_flag = ?, is_deleted = ?, captured_at = ? WHERE revision_id = ?`
    ).run(msg.content, msg.contentHash, msg.editedFlag ? 1 : 0, msg.isDeleted ? 1 : 0, msg.observedAt, latestRevision.revision_id);

    return db
      .prepare("SELECT * FROM message_revisions WHERE revision_id = ?")
      .get(latestRevision.revision_id) as unknown as MessageRevision;
  }

  const info = db
    .prepare(
      `INSERT INTO message_revisions (message_id, content, content_hash, edited_flag, is_deleted, captured_at)
       VALUES (?, ?, ?, ?, ?, ?)`
    )
    .run(
      msg.messageId,
      msg.content,
      msg.contentHash,
      msg.editedFlag ? 1 : 0,
      msg.isDeleted ? 1 : 0,
      msg.observedAt
    );

  return db
    .prepare("SELECT * FROM message_revisions WHERE revision_id = ?")
    .get(info.lastInsertRowid) as unknown as MessageRevision;
}

export function getAttachmentSnapshot(messageId: string, name: string): AttachmentSnapshot | undefined {
  return db
    .prepare("SELECT * FROM attachment_snapshots WHERE message_id = ? AND name = ?")
    .get(messageId, name) as unknown as AttachmentSnapshot | undefined;
}

/** Idempotent — a re-visit that finds the same (message_id, name) again just no-ops. */
export function saveAttachmentSnapshot(snap: AttachmentSnapshot): void {
  db.prepare(
    `INSERT INTO attachment_snapshots (message_id, name, content_hash, content_type, byte_size, captured_at)
     VALUES (?, ?, ?, ?, ?, ?)
     ON CONFLICT(message_id, name) DO NOTHING`
  ).run(snap.message_id, snap.name, snap.content_hash, snap.content_type, snap.byte_size, snap.captured_at);
}

export function getMessagesForConversation(conversationId: string) {
  return db
    .prepare(
      `SELECT m.message_id, m.sender, m.is_deleted, m.first_seen_at, m.last_seen_at
       FROM messages m WHERE m.conversation_id = ?
       ORDER BY m.first_seen_at ASC`
    )
    .all(conversationId) as Array<{
    message_id: string;
    sender: string | null;
    is_deleted: number;
    first_seen_at: string;
    last_seen_at: string;
  }>;
}

export function getRevisionsForMessage(messageId: string): MessageRevision[] {
  return db
    .prepare(`SELECT * FROM message_revisions WHERE message_id = ? ORDER BY revision_id ASC`)
    .all(messageId) as unknown as MessageRevision[];
}

export function getLatestRevision(messageId: string): MessageRevision | undefined {
  return db
    .prepare(`SELECT * FROM message_revisions WHERE message_id = ? ORDER BY revision_id DESC LIMIT 1`)
    .get(messageId) as unknown as MessageRevision | undefined;
}

export function getLatestRevisionsForConversation(conversationId: string) {
  return db
    .prepare(
      `SELECT m.message_id, m.sender, m.is_deleted, m.first_seen_at,
              r.content, r.content_hash, r.edited_flag, r.captured_at, r.revision_id
       FROM messages m
       JOIN message_revisions r ON r.revision_id = (
         SELECT revision_id FROM message_revisions
         WHERE message_id = m.message_id
         ORDER BY revision_id DESC LIMIT 1
       )
       WHERE m.conversation_id = ?
       ORDER BY m.first_seen_at ASC`
    )
    .all(conversationId);
}

export default db;
