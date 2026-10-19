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
import { localFilePath } from "../core/attachments.js";

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

  // Per-message equivalent of notifySyncStatus, for a message whose latest
  // content the connector knows to be incomplete (e.g. an edit whose
  // attachments its real-time path couldn't carry) and is working on,
  // rather than "which conversation is the connector on right now".
  // Why a given message is incomplete is entirely the connector's
  // business — see Connector/ConnectorContext in core/types.ts.
  // No REST snapshot endpoint for this, same as sync
  // status above — purely event-driven, so a dashboard that loads mid-
  // pending won't show the spinner until the next state change, which
  // matches how the conversation-level sync indicator already works.
  const notifyMessagePending = (messageId: string, pending: boolean) =>
    broadcastRaw({ type: "messagePending", messageId, pending });

  // This server starts once and stays up; the connector underneath it does
  // not necessarily stay connected (the teams-web one tears its browser
  // down and back up across re-logins, and a connector may not be able to
  // send at all — Connector.send is optional). So rather than the server
  // holding a connector reference directly, index.ts registers whichever
  // send function is currently valid, or null when there is none, and the
  // send endpoint below just calls whatever's currently registered.
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

  // Serves a locally-downloaded image attachment (see core/attachments.ts)
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

  // Doesn't write anything to the DB itself — Teams echoes the message
  // back through the connector's own observation path moments after
  // sending, same as any message from anyone else, and gets recorded
  // through that one path (see Connector.send in core/types.ts).
  app.post("/api/conversations/:id/send", async (req, res) => {
    const { text } = req.body ?? {};
    if (typeof text !== "string" || !text.trim()) {
      res.status(400).json({ error: "text is required" });
      return;
    }
    if (!sendHandler) {
      res.status(503).json({ error: "The connector can't send right now — check the watcher's console." });
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

  return { notifyConversationUpdate, notifyAuth, notifySyncStatus, notifyMessagePending, setSendHandler };
}
