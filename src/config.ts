import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const projectRoot = path.resolve(__dirname, "..");

// A debugging session (Claude, or anyone poking at this alongside the
// user's own running instance) shouldn't have to share a port or a live
// database with it — WATCHER_DATA_DIR points dbPath/attachmentsDir at an
// isolated copy instead (see README's "Debugging without disrupting a
// live session" for how to clone one), and WATCHER_PORT picks a different
// port so the two servers don't collide. Both default to the normal
// values, so a plain `npm run dev` is completely unaffected.
const dataDir = process.env.WATCHER_DATA_DIR
  ? path.resolve(process.env.WATCHER_DATA_DIR)
  : path.join(projectRoot, "data");

export const config = {
  // Which connector actually observes messages (see src/connectors/index.ts
  // for the registry of names). Everything outside src/connectors is
  // written against the Connector interface in core/types.ts, so this is
  // the only switch needed to run a different one.
  connector: process.env.WATCHER_CONNECTOR ?? "teams-web",

  dbPath: path.join(dataDir, "teams.db"),

  // Downloaded image attachment bytes (see core/attachments.ts) — kept as
  // plain files on disk rather than DB blobs, named by content hash so
  // serving them is a direct file read with no query needed beyond
  // looking up which hash a (message, name) pair maps to.
  attachmentsDir: path.join(dataDir, "attachments"),

  server: {
    port: Number(process.env.WATCHER_PORT) || 4173,
    // Set WATCHER_NO_OPEN_BROWSER=1 to skip auto-opening the dashboard in
    // the OS default browser on startup — meant for a debugging session
    // that restarts the watcher repeatedly (each restart would otherwise
    // pop a fresh tab in whatever browser you're actually working in,
    // unrelated to the dedicated Edge automation profile the teams-web
    // connector drives separately).
    autoOpenBrowser: process.env.WATCHER_NO_OPEN_BROWSER !== "1",
  },

  /**
   * Settings belonging to the "teams-web" connector specifically (Playwright
   * driving a real logged-in Edge profile). Deliberately nested rather
   * than sitting at the top level: none of it means anything to another
   * connector, and keeping it scoped is what makes "is this shared, or is
   * this one connector's implementation detail?" answerable at a glance.
   */
  teamsWeb: {
    teamsUrl: "https://teams.microsoft.com/v2/",

    // Dedicated Edge profile for this tool. Deliberately NOT the shared
    // "van-assistant" profile: Chromium locks a profile's user-data-dir to
    // a single running browser process, so a long-lived watcher here would
    // block any other tool (e.g. the PreQA skill) from using that profile,
    // and vice versa. Log into Teams manually under this profile once
    // (see README) before running the watcher.
    //
    // That same lock means WATCHER_DATA_DIR/WATCHER_PORT alone don't fully
    // isolate a debugging session — this profile is a single shared
    // resource (the one real, logged-in Teams browser session), so a
    // SECOND browser PROCESS pointed at it while the user's own instance is
    // running will fail to launch at all, not just collide on data. There's
    // no WATCHER_PROFILE_DIR env var for this on purpose: a second profile
    // needs its own separate Teams login to be useful, which is the user's
    // call to set up, not something to default into silently.
    //
    // What live investigation actually needs isn't a second process at all
    // though — it's a second TAB in the browser the watcher already has
    // open, which doesn't touch the profile lock (it's the same process).
    // WATCHER_CDP_PORT exposes Chromium's remote-debugging port for exactly
    // that: a separate script can `chromium.connectOverCDP(...)` into the
    // already-running instance and open/close its own page there, fully
    // alongside the watcher's own tab and without disturbing it. Unset by
    // default (no open debug port on a normal run) — only meant to be set
    // for a debugging session.
    edgeProfileDir: path.join(projectRoot, "edge-profile", "teams-watcher"),
    edgeChannel: "msedge" as const,
    debugCdpPort: process.env.WATCHER_CDP_PORT ? Number(process.env.WATCHER_CDP_PORT) : null,

    // Single-tab backup visit scheduler: real-time push (trouter.ts) is
    // the primary way new/edited/deleted messages get caught, near-
    // instantly, for every tracked conversation at once — this is just a
    // slower safety net in case the push connection ever drops. Every
    // conversation is treated uniformly (no more "live"/"polling"
    // distinction — that only ever changed how often this backup ran, and
    // stopped meaning anything once push started covering every
    // conversation equally).
    backupCheckIntervalMs: 5 * 60 * 1000,

    // How often the scheduler re-scans the left rail for conversations
    // that don't have a tracked_conversations row yet, and auto-tracks
    // them (baseline watermark only, no backfill — same as tracking any
    // other conversation). Covers both "every existing conversation" at
    // startup and any brand-new one that shows up later.
    discoveryIntervalMs: 5 * 60 * 1000,
  },

  /**
   * Settings belonging to the "teams-bot" connector specifically (a Bot
   * Framework bot added to groups/channels on demand).
   *
   * Unlike teams-web, this one CANNOT work from local state alone: Teams
   * pushes activities to it, so it needs an Azure Bot registration and a
   * publicly reachable HTTPS endpoint tunnelled to `port` below. See
   * README's "Running the teams-bot connector" checklist.
   */
  teamsBot: {
    // Local port for the activity listener, bound to loopback only — the
    // tunnel is what makes it reachable. 3978 is the Bot Framework
    // convention, kept so the tunnel/emulator defaults line up.
    port: Number(process.env.WATCHER_BOT_PORT) || 3978,
    endpointPath: "/api/messages",

    // Standard Bot Framework credential env var names, deliberately NOT
    // renamed to WATCHER_*: they're what the Azure portal, the Bot Framework
    // Emulator and every piece of Microsoft documentation calls them, and
    // inventing local aliases would just make those instructions wrong.
    //
    // With appId/appPassword unset, the SDK skips inbound JWT validation
    // altogether — which is what makes scripts/bot-activity-harness.mjs
    // work with no Azure setup, and why the connector refuses to send and
    // warns loudly in that mode. Never put a tunnel in front of it unset.
    appId: process.env.MICROSOFT_APP_ID ?? "",
    appPassword: process.env.MICROSOFT_APP_PASSWORD ?? "",
    appType: process.env.MICROSOFT_APP_TYPE ?? "MultiTenant",
    appTenantId: process.env.MICROSOFT_APP_TENANT_ID ?? "",

    // Bot Framework can't address a conversation by id alone; a proactive
    // send needs a ConversationReference that only ever arrives attached to
    // an inbound activity. Persisted so a restart doesn't leave the
    // dashboard unable to send into a quiet conversation until it next
    // produces a message. Lives in the data dir so WATCHER_DATA_DIR
    // isolates it along with everything else.
    referenceFile: path.join(dataDir, "teams-bot-references.json"),
  },
} as const;
