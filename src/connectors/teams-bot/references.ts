import fs from "node:fs";
import type { ConversationReference } from "botbuilder";
import path from "node:path";

/**
 * Where the bot keeps the `ConversationReference` it needs in order to send
 * a message into a conversation it isn't currently being spoken to in.
 *
 * Bot Framework can't address a conversation by id alone — a proactive send
 * needs the service URL, the bot's own identity within that conversation
 * and the conversation descriptor, all of which only arrive attached to an
 * inbound activity. So each inbound activity's reference is saved, and
 * sending looks one up.
 *
 * Without persistence, a restart would leave the dashboard unable to send
 * into any conversation until that conversation happened to produce another
 * message — which for a quiet channel could be days.
 *
 * Deliberately a connector-owned JSON file rather than a table in the
 * shared DB: this is one connector's implementation detail (teams-web has
 * no equivalent and never will), and the schema in src/db is meant to stay
 * about messages, not about how any particular connector addresses Teams.
 * Small enough — one smallish object per conversation the bot is in — that
 * a rewrite-on-change file is the right tool.
 */

/**
 * Bot Framework's own reference type, kept as-is rather than re-declared
 * with just the fields we use: a parallel hand-written shape has to stay
 * structurally assignable to this one, and it already drifted twice (the
 * SDK requires real strings where an inbound activity may carry nothing).
 * Partial because an inbound activity doesn't always populate everything,
 * and the SDK accepts a partial reference.
 */
export type StoredReference = Partial<ConversationReference>;

export class ReferenceStore {
  #filePath: string;
  #refs = new Map<string, StoredReference>();

  constructor(filePath: string) {
    this.#filePath = filePath;
    this.#load();
  }

  #load(): void {
    try {
      const raw = JSON.parse(fs.readFileSync(this.#filePath, "utf-8")) as Record<string, StoredReference>;
      for (const [id, ref] of Object.entries(raw)) this.#refs.set(id, ref);
      console.log(`[teams-bot] loaded ${this.#refs.size} conversation reference(s)`);
    } catch (err) {
      // A missing file is the normal first-run case. A corrupt one is worth
      // saying out loud, but not worth refusing to start over — the store
      // rebuilds itself from inbound activities, so the cost is only that
      // sending is unavailable in quiet conversations until they speak.
      if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
        console.warn(`[teams-bot] couldn't read ${this.#filePath}, starting empty:`, err);
      }
    }
  }

  #persist(): void {
    const out = Object.fromEntries(this.#refs);
    fs.mkdirSync(path.dirname(this.#filePath), { recursive: true });
    // Written via a temp file + rename so a crash mid-write can't leave a
    // truncated file that the next start would have to discard wholesale.
    const tmp = `${this.#filePath}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(out, null, 2));
    fs.renameSync(tmp, this.#filePath);
  }

  get(conversationId: string): StoredReference | undefined {
    return this.#refs.get(conversationId);
  }

  /** Saves a reference, writing to disk only when something actually changed — every inbound activity calls this. */
  remember(conversationId: string, ref: StoredReference): void {
    const existing = this.#refs.get(conversationId);
    if (existing && JSON.stringify(existing) === JSON.stringify(ref)) return;
    this.#refs.set(conversationId, ref);
    try {
      this.#persist();
    } catch (err) {
      // Losing a reference costs the ability to send into that conversation
      // after a restart; it doesn't affect observing messages at all, so
      // it's not worth failing the activity over.
      console.warn("[teams-bot] couldn't persist conversation references:", err);
    }
  }
}
