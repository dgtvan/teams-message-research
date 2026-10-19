import { createServer, type Server } from "node:http";
import express from "express";
import {
  CloudAdapter,
  ConfigurationBotFrameworkAuthentication,
  ConfigurationServiceClientCredentialFactory,
  type TurnContext,
} from "botbuilder";
import { config } from "../../config.js";
import { resolveAttachmentMarkers } from "../../core/attachments.js";
import type { Connector, ConnectorContext } from "../../core/types.js";
import {
  classifyActivity,
  describeConversation,
  isBotAddedToConversation,
  isBotRemovedFromConversation,
  normalizeConversationId,
  type TeamsActivity,
} from "./activities.js";
import { ReferenceStore, type StoredReference } from "./references.js";

/**
 * The "teams-bot" connector: observes Teams through a Bot Framework bot
 * that's been added to specific groups/channels, and sends back into them
 * as the bot.
 *
 * The shape is the mirror image of the teams-web connector, and the
 * differences are worth understanding before relying on it:
 *
 * - **Inbound, not outbound.** teams-web drives a browser and pulls; this
 *   runs an HTTP listener and Teams pushes activities at it. That means it
 *   needs a publicly reachable HTTPS endpoint (a dev tunnel in front of the
 *   local port) and an Azure Bot registration — a bot can't be added to a
 *   Teams conversation without one. See README's bot setup checklist.
 *
 * - **No discovery scan.** A bot only learns a conversation exists by being
 *   put in it, so there's no equivalent of walking the left rail. Tracking
 *   begins on the conversationUpdate that adds the bot, and its baseline is
 *   empty (no history — see below), so everything from that moment on is new.
 *
 * - **No history, therefore no catch-up and no self-healing.** Bot Framework
 *   has no "read this conversation's messages" call, so this connector can
 *   only ever report what it was pushed and never implements
 *   ingest.snapshot(). Two real consequences: messages sent while this app
 *   is down are lost permanently (Bot Framework retries briefly, then drops
 *   them), and a later improvement to this file can't retroactively correct
 *   already-stored messages the way a teams-web DOM re-read does. Reading
 *   history would need Microsoft Graph with resource-specific consent
 *   (ChannelMessage.Read.Group / ChatMessage.Read.Chat), which is a separate
 *   mechanism from the activity stream and isn't wired up here.
 *
 * - **It does record its own sends.** Unlike teams-web, which relies on
 *   Teams echoing a sent message back through its own push path, Bot
 *   Framework never delivers the bot's own messages back to it. So send()
 *   records what it sent, or it would never appear in the dashboard at all.
 *
 * - **Attachments are links, not bytes.** See extractAttachments in
 *   ./activities.ts.
 */
export function createTeamsBotConnector(): Connector {
  return new TeamsBotConnector();
}

class TeamsBotConnector implements Connector {
  readonly name = "teams-bot";

  #server: Server | null = null;
  #adapter: CloudAdapter | null = null;
  #refs = new ReferenceStore(config.teamsBot.referenceFile);
  #ctx: ConnectorContext | null = null;
  #stopped = false;

  async run(ctx: ConnectorContext): Promise<void> {
    this.#ctx = ctx;
    const { appId, appPassword, appType, appTenantId, port, endpointPath } = config.teamsBot;

    // With no credentials configured, the Bot Framework SDK skips inbound
    // JWT validation entirely — which is exactly what the local harness
    // (scripts/bot-activity-harness.mjs) relies on, and exactly what must
    // never be exposed to a real network. Said out loud on startup rather
    // than left to be discovered.
    if (!appId || !appPassword) {
      const message =
        "No bot credentials set (MICROSOFT_APP_ID / MICROSOFT_APP_PASSWORD) — running UNAUTHENTICATED. " +
        "Fine for the local harness; do NOT put a tunnel in front of this. Sending is disabled.";
      console.warn(`[teams-bot] ${message}`);
      ctx.reportAuth({ state: "needs-user-action", message });
    } else {
      ctx.reportAuth({ state: "ok" });
    }

    const auth = new ConfigurationBotFrameworkAuthentication(
      {},
      new ConfigurationServiceClientCredentialFactory({
        MicrosoftAppId: appId,
        MicrosoftAppPassword: appPassword,
        MicrosoftAppType: appType,
        MicrosoftAppTenantId: appTenantId,
      })
    );

    const adapter = new CloudAdapter(auth);
    this.#adapter = adapter;

    // Errors are logged and swallowed rather than propagated: a single
    // malformed or unexpected activity must not take the listener down,
    // because there's no redelivery to fall back on.
    adapter.onTurnError = async (_context, error) => {
      console.error("[teams-bot] error handling an activity:", error);
    };

    const app = express();
    app.use(express.json());

    // A plain GET is handy for confirming a tunnel actually reaches this
    // process before involving Teams at all.
    app.get(endpointPath, (_req, res) => {
      res.json({ ok: true, connector: this.name, authenticated: !!(appId && appPassword) });
    });

    app.post(endpointPath, async (req, res) => {
      await adapter.process(req, res, (turnContext) => this.#handleTurn(turnContext, ctx));
    });

    const server = createServer(app);
    this.#server = server;

    await new Promise<void>((resolve, reject) => {
      server.once("error", reject);
      // Bound to loopback on purpose: reaching this from Teams is the
      // tunnel's job, and binding all interfaces would put an endpoint that
      // may be running unauthenticated (see above) on the local network.
      server.listen(port, "127.0.0.1", () => {
        console.log(`[teams-bot] listening on http://127.0.0.1:${port}${endpointPath}`);
        resolve();
      });
    });

    // Nothing more to drive — Teams pushes from here on. Resolves only when
    // stop() closes the listener, matching run()'s contract that it runs
    // until stopped.
    await new Promise<void>((resolve) => {
      if (this.#stopped) return resolve();
      server.once("close", resolve);
    });
  }

  async stop(): Promise<void> {
    this.#stopped = true;
    const server = this.#server;
    this.#server = null;
    if (!server) return;
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }

  /**
   * Sends as the bot, into a conversation the bot is in.
   *
   * Proactive (continueConversationAsync) rather than a reply on an open
   * turn, because a dashboard send has no inbound activity behind it — it
   * arrives out of band, whenever you happen to type something.
   *
   * Records what it sent, unlike the teams-web connector: Bot Framework
   * does not deliver a bot's own messages back to it, so nothing else ever
   * would.
   */
  async send(conversationId: string, text: string): Promise<void> {
    const adapter = this.#adapter;
    const ctx = this.#ctx;
    if (!adapter || !ctx) throw new Error("The bot connector isn't running.");

    const { appId } = config.teamsBot;
    if (!appId) {
      throw new Error("The bot has no credentials configured, so it can't send. Set MICROSOFT_APP_ID/PASSWORD.");
    }

    const ref = this.#refs.get(conversationId);
    if (!ref) {
      throw new Error(
        "The bot hasn't seen this conversation yet, so it has no way to address it. " +
          "It can only send where it's a member and has received at least one activity."
      );
    }

    let sentMessageId: string | null = null;
    let botName: string | null = null;
    await adapter.continueConversationAsync(appId, ref, async (turnContext) => {
      const response = await turnContext.sendActivity({ type: "message", text });
      sentMessageId = response?.id ?? null;
      botName = turnContext.activity?.from?.name ?? ref.bot?.name ?? null;
    });

    if (!sentMessageId) {
      // The send itself succeeded, but with no id there's nothing to key a
      // message row on. Better to leave the dashboard missing one message
      // than to invent an id that a later edit/delete could never match.
      console.warn("[teams-bot] send returned no message id — not recording it locally");
      return;
    }

    ctx.ingest.incremental({
      kind: "new",
      conversationId,
      messageId: sentMessageId,
      sender: botName ?? "Bot",
      content: text,
    });
  }

  async #handleTurn(turnContext: TurnContext, ctx: ConnectorContext): Promise<void> {
    const activity = turnContext.activity as TeamsActivity;
    const rawConversationId = activity.conversation?.id;
    if (!rawConversationId) return;
    const conversationId = normalizeConversationId(rawConversationId);

    // Saved for every inbound activity, not just the ones we record: a
    // typing indicator or reaction is a perfectly good source of the
    // addressing details a later dashboard send needs. Guards internally on
    // the fields actually being present.
    this.#rememberReference(conversationId, turnContext);

    if (isBotRemovedFromConversation(activity)) {
      console.log(`[teams-bot] removed from "${conversationId}" — keeping what was already captured`);
      return;
    }

    // Tracking starts when the bot is added, and also lazily on any message
    // from a conversation we don't know yet — the add event can be missed
    // (it's sent once, with no redelivery if this app happens to be down for
    // it), and without this fallback such a conversation would stay
    // permanently invisible despite the bot sitting in it.
    const described = describeConversation(activity);
    if (described && (isBotAddedToConversation(activity) || !ctx.ingest.trackedConversation(conversationId))) {
      const untracked = ctx.ingest.discover([described]);
      for (const conv of untracked) {
        // No baseline message id: a bot has no history to baseline against,
        // so everything from now on counts as new.
        ctx.ingest.baseline(conv, null);
        console.log(`[teams-bot] now tracking "${conv.title}" (${conv.kind})`);
      }
    }

    const observed = classifyActivity(activity);
    if (!observed) return;

    // Markers are resolved here rather than in ./activities.ts for the same
    // reason the teams-web connector does it at this level: it needs storage
    // access that the parsing layer deliberately doesn't have. No image
    // bytes are passed — see extractAttachments.
    const content = resolveAttachmentMarkers(
      observed.text,
      observed.message.messageId,
      observed.attachments,
      []
    );

    ctx.ingest.incremental({ ...observed.message, content });
  }

  #rememberReference(conversationId: string, turnContext: TurnContext): void {
    const activity = turnContext.activity;
    if (!activity.serviceUrl || !activity.channelId || !activity.conversation?.id) return;
    const ref: StoredReference = {
      serviceUrl: activity.serviceUrl,
      channelId: activity.channelId,
      conversation: {
        // The RAW id, suffix and all — this is what Bot Framework must be
        // handed back to address the conversation, even though everything on
        // our side keys off the normalized one. A channel reply's thread
        // suffix would make the bot reply into that specific thread; the
        // base channel id posts a new top-level message, which is what a
        // dashboard send means.
        id: conversationId,
        conversationType: activity.conversation.conversationType,
        isGroup: activity.conversation.isGroup,
        tenantId: activity.conversation.tenantId,
        name: activity.conversation.name,
      },
      ...(activity.recipient?.id
        ? { bot: { id: activity.recipient.id, name: activity.recipient.name ?? "" } }
        : {}),
    };
    this.#refs.remember(conversationId, ref);
  }
}
