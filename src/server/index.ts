import express from "express";
import { createServer } from "node:http";
import { WebSocketServer, WebSocket } from "ws";
import { exec } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { config } from "../config.js";
import {
  listTrackedConversations,
  getLatestRevisionsForConversation,
  getRevisionsForMessage,
  getAttachmentSnapshot,
  setPinned,
  markRead,
} from "../db/index.js";
import { localFilePath } from "../watcher/attachments.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** Opens the dashboard in the OS default browser so you never have to remember the port — skipped entirely when config.server.autoOpenBrowser is false (WATCHER_NO_OPEN_BROWSER=1). */
function openInBrowser(url: string) {
  const cmd =
    process.platform === "win32"
      ? `start "" "${url}"`
      : process.platform === "darwin"
        ? `open "${url}"`
        : `xdg-open "${url}"`;
  exec(cmd, (err) => {
    if (err) console.warn("[server] couldn't auto-open the browser:", err.message);
  });
}

type SendHandler = (conversationId: string, text: string) => Promise<void>;

export function startServer() {
  const app = express();
  app.use(express.json());
  app.use(express.static(path.join(__dirname, "..", "..", "public")));

  const httpServer = createServer(app);
  const wss = new WebSocketServer({ server: httpServer });

  const broadcastRaw = (payload: unknown) => {
    const data = JSON.stringify(payload);
    for (const client of wss.clients) {
      if (client.readyState === WebSocket.OPEN) client.send(data);
    }
  };

  const notifyConversationUpdate = (conversationId: string) =>
    broadcastRaw({ type: "update", conversationId });

  const notifyAuth = (status: "required" | "resolved") => broadcastRaw({ type: "auth", status });

  const notifySyncStatus = (conversationId: string, syncing: boolean) =>
    broadcastRaw({ type: "sync", conversationId, syncing });

  // The watcher's browser page is created/torn down across re-logins (see
  // index.ts's runHeadlessSession/runHeadedLoginSession loop), independent
  // of this server, which starts once and stays up — so rather than the
  // server holding a `page` reference directly, index.ts registers
  // whichever send function is currently valid (or null while there's no
  // active headless session, e.g. mid-relogin) and the send endpoint below
  // just calls whatever's currently registered.
  let sendHandler: SendHandler | null = null;
  const setSendHandler = (handler: SendHandler | null) => {
    sendHandler = handler;
  };

  app.get("/api/conversations", (_req, res) => {
    res.json(listTrackedConversations());
  });

  app.post("/api/conversations/:id/pin", (req, res) => {
    const { pinned } = req.body ?? {};
    setPinned(req.params.id, !!pinned);
    res.json({ ok: true });
  });

  app.post("/api/conversations/:id/read", (req, res) => {
    markRead(req.params.id, new Date().toISOString());
    res.json({ ok: true });
  });

  app.get("/api/conversations/:id/messages", (req, res) => {
    res.json(getLatestRevisionsForConversation(req.params.id));
  });

  app.get("/api/messages/:messageId/revisions", (req, res) => {
    res.json(getRevisionsForMessage(req.params.messageId));
  });

  // Serves a locally-downloaded image attachment (see watcher/attachments.ts)
  // — the actual bytes, not a redirect to Teams' own URL, since that one
  // needs an authenticated Teams session this dashboard doesn't have.
  app.get("/api/attachments/:messageId/:name", (req, res) => {
    const snap = getAttachmentSnapshot(req.params.messageId, req.params.name);
    if (!snap) {
      res.status(404).end();
      return;
    }
    res.type(snap.content_type);
    res.sendFile(localFilePath(snap), (err) => {
      if (err && !res.headersSent) res.status(404).end();
    });
  });

  // Doesn't write anything to the DB itself — the message comes back
  // through Teams' own real-time push moments after sending, same as any
  // message from anyone else, and gets recorded through that one path
  // (see sendMessageTo's comment in watcher/index.ts).
  app.post("/api/conversations/:id/send", async (req, res) => {
    const { text } = req.body ?? {};
    if (typeof text !== "string" || !text.trim()) {
      res.status(400).json({ error: "text is required" });
      return;
    }
    if (!sendHandler) {
      res.status(503).json({ error: "Watcher isn't connected right now — check the Edge window." });
      return;
    }
    try {
      await sendHandler(req.params.id, text);
      res.json({ ok: true });
    } catch (err) {
      res.status(500).json({ error: err instanceof Error ? err.message : String(err) });
    }
  });

  httpServer.listen(config.server.port, () => {
    const url = `http://localhost:${config.server.port}`;
    console.log(`[server] dashboard at ${url}`);
    if (config.server.autoOpenBrowser) openInBrowser(url);
  });

  return { notifyConversationUpdate, notifyAuth, notifySyncStatus, setSendHandler };
}
