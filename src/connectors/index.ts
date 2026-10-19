import type { Connector } from "../core/types.js";
import { createTeamsBotConnector } from "./teams-bot/index.js";
import { createTeamsWebConnector } from "./teams-web/index.js";

/**
 * Every connector this app knows how to run, by name — the one place that
 * has to change when a new one is added (alongside its own directory).
 *
 * Exactly one runs at a time, chosen by WATCHER_CONNECTOR. They are
 * alternatives, not layers: both write through the same ingest into the same
 * DB, so switching between them keeps all previously captured history, but
 * they are never run concurrently.
 *
 * Factories, not instances: constructing a connector shouldn't launch a
 * browser or bind a port, so only the selected one is ever built.
 */
const CONNECTORS: Record<string, () => Connector> = {
  // Drives a real logged-in Edge profile. Sees every chat on the account,
  // can re-read history, sends as you.
  "teams-web": createTeamsWebConnector,
  // A Bot Framework bot invited into specific groups/channels. Sees only
  // where it's a member, has no history at all, sends as the bot.
  "teams-bot": createTeamsBotConnector,
};

export function createConnector(name: string): Connector {
  const factory = CONNECTORS[name];
  if (!factory) {
    throw new Error(
      `Unknown connector "${name}". Available: ${Object.keys(CONNECTORS).join(", ")} (set WATCHER_CONNECTOR).`
    );
  }
  return factory();
}
