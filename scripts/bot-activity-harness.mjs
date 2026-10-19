/**
 * Posts synthetic Teams activities at the teams-bot connector's local
 * endpoint, so the whole inbound path — activity parsing, conversation
 * tracking, ingest, revision history, dashboard broadcast — can be
 * exercised without an Azure Bot registration, a tunnel, or a real Teams
 * group.
 *
 * This works only because the Bot Framework SDK skips inbound JWT
 * validation when no MICROSOFT_APP_ID/PASSWORD is configured (the same
 * behaviour the Bot Framework Emulator relies on). That is also exactly why
 * the connector refuses to send and warns loudly in that mode, and why the
 * listener binds to loopback only.
 *
 * Usage — with the watcher running under WATCHER_CONNECTOR=teams-bot:
 *
 *   node scripts/bot-activity-harness.mjs                 # full scenario
 *   node scripts/bot-activity-harness.mjs --port 3978
 *   node scripts/bot-activity-harness.mjs --conversation 19:my-test@thread.v2
 *
 * The activity shapes below mirror what Teams actually sends, including the
 * quirks the connector has to cope with: edits and deletes reusing the
 * original message id, channel replies carrying a `;messageid=` thread
 * suffix, and formatted text arriving as HTML.
 */

const args = process.argv.slice(2);
const argValue = (name, fallback) => {
  const i = args.indexOf(`--${name}`);
  return i !== -1 && args[i + 1] ? args[i + 1] : fallback;
};

const PORT = Number(argValue("port", process.env.WATCHER_BOT_PORT || 3978));
const ENDPOINT = `http://127.0.0.1:${PORT}/api/messages`;
const CONVERSATION_ID = argValue("conversation", "19:harness-group@thread.v2");
const BOT_ID = "28:00000000-0000-0000-0000-00000000bot0";
const SERVICE_URL = "https://smba.trafficmanager.net/harness/";

// Teams message ids are stringified tick counts, and the watermark compares
// them as BigInt — so they have to be numeric and increasing to be treated
// as new. Derived from the clock so repeated harness runs never collide.
let nextId = Date.now();
const newMessageId = () => String(nextId++);

function baseActivity(overrides) {
  return {
    serviceUrl: SERVICE_URL,
    channelId: "msteams",
    timestamp: new Date().toISOString(),
    from: { id: "29:harness-user", name: "Harness User", aadObjectId: "00000000-0000-0000-0000-0000000user" },
    recipient: { id: BOT_ID, name: "Watcher Bot" },
    conversation: { id: CONVERSATION_ID, conversationType: "groupChat", isGroup: true, name: "Harness Group" },
    ...overrides,
  };
}

async function post(label, activity) {
  const res = await fetch(ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(activity),
  });
  const body = await res.text();
  const ok = res.status >= 200 && res.status < 300;
  console.log(`${ok ? "  ok " : " FAIL"}  ${label}  → ${res.status}${body ? ` ${body.slice(0, 120)}` : ""}`);
  if (!ok) process.exitCode = 1;
  return ok;
}

async function main() {
  console.log(`Posting synthetic activities to ${ENDPOINT}`);
  console.log(`Conversation: ${CONVERSATION_ID}\n`);

  // Confirm something is actually listening before firing a whole scenario
  // at it, so a stopped watcher reads as one clear message rather than six
  // connection errors.
  try {
    const probe = await fetch(ENDPOINT, { method: "GET" });
    const info = await probe.json().catch(() => ({}));
    if (info.authenticated) {
      console.error(
        "The connector has credentials configured, so it will reject these unsigned activities.\n" +
          "Unset MICROSOFT_APP_ID/MICROSOFT_APP_PASSWORD to use the harness."
      );
      process.exit(1);
    }
    console.log("Endpoint is up and running unauthenticated — proceeding.\n");
  } catch {
    console.error(
      `Nothing is listening on ${ENDPOINT}.\n` +
        "Start the watcher first, e.g.:\n" +
        "  WATCHER_CONNECTOR=teams-bot WATCHER_DATA_DIR=data/debug WATCHER_PORT=4174 \\\n" +
        "    WATCHER_NO_OPEN_BROWSER=1 npm run dev"
    );
    process.exit(1);
  }

  // 1. The bot gets added — the bot's entire equivalent of discovery.
  await post(
    "conversationUpdate: bot added to the group",
    baseActivity({ type: "conversationUpdate", membersAdded: [{ id: BOT_ID }] })
  );

  // 2. A plain message.
  const plainId = newMessageId();
  await post("message: plain text", baseActivity({ type: "message", id: plainId, text: "hello from the harness" }));

  // 3. A formatted message — Teams sends real HTML for anything with
  //    formatting in it, which the connector runs through htmlToText.
  await post(
    "message: HTML formatting",
    baseActivity({
      type: "message",
      id: newMessageId(),
      textFormat: "xml",
      text: "<div>first line</div><div>second <strong>bold</strong> line</div>",
    })
  );

  // 4. An edit of message 2 — same id, which is what makes ingest attach a
  //    new revision instead of creating a second message.
  await post(
    "messageUpdate: edit of the plain message (same id)",
    baseActivity({
      type: "messageUpdate",
      id: plainId,
      text: "hello from the harness (edited)",
      channelData: { eventType: "editMessage" },
    })
  );

  // 5. A message with a file attachment, to exercise marker building.
  await post(
    "message: with a file attachment",
    baseActivity({
      type: "message",
      id: newMessageId(),
      text: "here's a file",
      attachments: [
        {
          contentType: "application/vnd.microsoft.teams.file.download.info",
          name: "notes.txt",
          content: { downloadUrl: "https://example.invalid/notes.txt", fileType: "txt" },
        },
      ],
    })
  );

  // 6. A soft delete of message 2.
  await post(
    "messageDelete: soft delete of the plain message",
    baseActivity({
      type: "messageDelete",
      id: plainId,
      channelData: { eventType: "softDeleteMessage" },
    })
  );

  // 7. A channel reply, whose conversation id carries a thread suffix the
  //    connector strips so the channel stays one conversation.
  const channelId = "19:harness-channel@thread.tacv2";
  await post(
    "message: channel reply (thread suffix should be stripped)",
    baseActivity({
      type: "message",
      id: newMessageId(),
      text: "a threaded reply in a channel",
      conversation: { id: `${channelId};messageid=1700000000000`, conversationType: "channel" },
      channelData: { team: { id: "19:team@thread.tacv2", name: "Harness Team" }, channel: { id: channelId, name: "General" } },
    })
  );

  // 8. Something we deliberately ignore, to confirm it's a no-op not a crash.
  await post("typing: should be ignored", baseActivity({ type: "typing" }));

  console.log(
    `\nDone. Check the dashboard, and expect:\n` +
      `  - a "Harness Group" conversation with 3 messages (one edited then deleted, so 3 revisions on it)\n` +
      `  - a "Harness Team / General" conversation with 1 message, keyed on ${channelId} with no thread suffix`
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
