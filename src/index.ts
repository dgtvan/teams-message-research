import { config } from "./config.js";
import { createConnector } from "./connectors/index.js";
import { createIngest } from "./core/ingest.js";
import type { ConnectorContext } from "./core/types.js";
import { startServer } from "./server/index.js";

/**
 * Composition root: starts the dashboard server, builds the ingest
 * pipeline, picks a connector and wires the two together. That's all.
 *
 * Everything about HOW messages are observed or sent — browsers, logins,
 * polling, sockets — belongs to the connector (src/connectors/*) and is
 * deliberately invisible here, so adding a second one is a matter of
 * writing it plus one line in the connector registry, with nothing in
 * storage, the server or the dashboard to touch.
 */
async function main() {
  const server = startServer();

  const ingest = createIngest({
    onConversationUpdate: server.notifyConversationUpdate,
  });

  const connector = createConnector(config.connector);
  console.log(`[watcher] using the "${connector.name}" connector`);

  const ctx: ConnectorContext = {
    ingest,
    reportAuth: (auth) => {
      if (auth.state === "needs-user-action") console.log(`[watcher] ${auth.message}`);
      server.notifyAuth(auth.state === "ok" ? "resolved" : "required");
    },
    reportSync: server.notifySyncStatus,
    reportMessagePending: server.notifyMessagePending,
  };

  // Only registered when the connector can actually send — the endpoint
  // reports "sending isn't available" rather than failing obscurely for a
  // read-only connector.
  server.setSendHandler(connector.send ? (id, text) => connector.send!(id, text) : null);

  // Tracked so a shutdown signal can tear the connector down instead of
  // leaving, say, an orphaned Edge process holding a profile lock. Covers
  // Ctrl+C (SIGINT) and a graceful stop from a process manager/supervisor
  // (SIGTERM) — a hard kill of this process instead (Task Manager "End
  // task", taskkill /F) skips this entirely, but on Windows Playwright
  // launches the browser inside a Job Object tied to this process's
  // lifetime, so the whole Edge tree still gets torn down automatically
  // even then.
  const shutdown = async () => {
    await connector.stop().catch(() => {});
    process.exit(0);
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);

  await connector.run(ctx);
}

main().catch((err) => {
  console.error("[fatal]", err);
  process.exit(1);
});
