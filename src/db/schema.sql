-- Conversations the watcher tracks. Every conversation discovered in the
-- left rail ends up here automatically (see the discovery loop) — nothing
-- is recorded for one until its row exists, but rows are no longer an
-- opt-in step you have to trigger yourself.
CREATE TABLE IF NOT EXISTS tracked_conversations (
  conversation_id   TEXT PRIMARY KEY,
  title             TEXT NOT NULL,
  kind              TEXT NOT NULL,              -- 'chat' | 'channel'
  tracking_started_at TEXT NOT NULL,
  last_watermark_message_id TEXT,               -- newest message_id we've fully processed
  last_watermark_at TEXT,
  last_synced_at    TEXT,
  last_activity_at  TEXT,                       -- last time a genuinely new message was observed (informational — "most recently active")
  pinned            INTEGER NOT NULL DEFAULT 0,
  last_read_at      TEXT                         -- last time you opened this conversation in the dashboard
);

-- One row per distinct message the watcher has ever seen in a tracked
-- conversation. Content itself lives in message_revisions, never here,
-- so an edit/delete never overwrites prior evidence.
CREATE TABLE IF NOT EXISTS messages (
  message_id      TEXT PRIMARY KEY,
  conversation_id TEXT NOT NULL REFERENCES tracked_conversations(conversation_id),
  sender          TEXT,
  first_seen_at   TEXT NOT NULL,
  last_seen_at    TEXT NOT NULL,
  is_deleted      INTEGER NOT NULL DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_messages_conversation
  ON messages(conversation_id);

-- Append-only log: a new row every time observed content for a message_id
-- changes (edit) or it flips to deleted. Never updated in place, with one
-- deliberate exception in recordObservedMessage (db/index.ts): filling in
-- a message's still-empty first revision (push's "pending" placeholder —
-- see connectors/teams-web/trouter.ts) completes that row rather than appending a
-- second one, since it was never a real state the message was actually
-- in from Teams' perspective.
CREATE TABLE IF NOT EXISTS message_revisions (
  revision_id   INTEGER PRIMARY KEY AUTOINCREMENT,
  message_id    TEXT NOT NULL REFERENCES messages(message_id),
  content       TEXT NOT NULL,
  content_hash  TEXT NOT NULL,
  edited_flag   INTEGER NOT NULL DEFAULT 0,     -- Teams' own "Edited" label, if present
  is_deleted    INTEGER NOT NULL DEFAULT 0,
  captured_at   TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_revisions_message
  ON message_revisions(message_id, revision_id);

-- A locally-downloaded copy of an image attachment's actual bytes (see
-- core/attachments.ts). Attachments otherwise aren't structured data at
-- all — just `[[ATTACHMENT|name|url]]` marker text inside a revision's
-- content (see app.js) — but an image needs real storage, not a marker,
-- because the remote URL Teams gives us requires an authenticated Teams
-- session to load and can't just be embedded and left to resolve itself
-- (confirmed live: 401s without it). content_hash is also the on-disk
-- filename under data/attachments/, so byte-identical images (e.g. the
-- same screenshot forwarded into multiple messages) are only ever stored
-- once regardless of how many (message_id, name) rows reference it.
CREATE TABLE IF NOT EXISTS attachment_snapshots (
  message_id    TEXT NOT NULL REFERENCES messages(message_id),
  name          TEXT NOT NULL,
  content_hash  TEXT NOT NULL,
  content_type  TEXT NOT NULL,
  byte_size     INTEGER NOT NULL,
  captured_at   TEXT NOT NULL,
  PRIMARY KEY (message_id, name)
);
