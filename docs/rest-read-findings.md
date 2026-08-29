# Finding: reading messages via REST also works

Investigated 2026-08-29, while debugging why sends were slow. Not acted on — this is a
record of the finding for whenever the backup-visit-loop/DOM-extraction path is next up
for work. Nothing in this repo currently uses this.

## What was confirmed live

Using a bearer token captured off a real chatsvc request (see `sendMessageViaRest` and
`warmChatServiceToken` in `src/watcher/teams.ts` for how sending already does this), a
plain `GET` against the same API family used for sending returns real data:

```
GET https://teams.microsoft.com/api/chatsvc/au/v1/users/ME/conversations/48%3Anotes/messages?view=msnp24Equivalent&pageSize=20
Authorization: Bearer eyJ...
→ 200, {"messages":[{ "sequenceId":276, "content":"<p>...</p>", "from":"...", "composetime":"...", ... }]}
```

Also confirmed working:
- `GET .../conversations/48%3Anotes/messages` (no query params) — same shape
- `GET .../conversations/48%3Anotes` — conversation metadata: `lastMessage`,
  `consumptionhorizon`, `lastSequenceId`, roster/thread properties

`"au"` is this account's own region prefix (same caveat as the send path — not a
general constant).

## Why this contradicts the README's existing "REST reading is dead" note

`src/watcher/chatsvc.ts` and the README both document an earlier attempt
(`fetchMessagesSince`) that called this same `chatsvc` API and got 401 "Authentication
failed" on every call, concluding REST reading doesn't work — because neither
`page.request` nor an in-page `fetch` automatically attaches the Bearer token this API
needs.

That conclusion was correct about the *technique tried*, not about the underlying
possibility: it never attempted to explicitly capture and attach a token itself, the
same gap that turned out to block sending too until solved. Once a token is captured
the same way (see `ensureChatTokenCapture`), GET works fine — same as POST does for
sending.

## Why this is a bigger deal than the send-path fix

The DOM-based backup visit loop and `catchUpAllTracked` — most of the startup latency
and ongoing per-conversation overhead in this app — exist specifically because DOM
extraction was believed to be the only viable path. If REST reading is a real
replacement, it could remove:
- `openConversation`'s click/wait-for-header/wait-for-stable-list sequence, and its
  8s/15s/5s timeout budget per conversation
- The `withPage` single-tab queue's biggest source of contention (every visit currently
  serializes through one shared tab)
- Selector fragility — no dependency on `src/watcher/selectors.ts` staying in sync with
  Teams' DOM

## What's NOT yet known — needs investigation before any of this is built

- **Attachments/images**: the DOM path reads image bytes directly out of the rendered
  `<img>` element's blob (see the long comment on `attachmentPreviewUrlAttr` in
  `selectors.ts`) specifically because the attachment's real URL 401s without an
  authenticated *browser* session — unclear whether the same captured bearer token
  would authenticate that URL directly, or whether attachments need a different token
  scope/mechanism entirely.
- **Incremental sync**: `sequenceId` and `consumptionhorizon` in the response look like
  exactly the cursor `fetchMessagesSince` was trying to use for a real "since X" sync
  instead of DOM's snapshot-only re-extraction. Not tested — only fetched the latest
  page so far, not a ranged/incremental query.
- **Edits and deletes**: DOM extraction's revision-tracking (`recordObservedMessage`)
  and push (`trouter.ts`) both currently handle these; unclear whether/how a GET
  response represents an edited or deleted message differently from a normal one.
- **Token scope for reading across ALL conversations, not just one already open**: only
  tested against "48:notes" specifically. Should confirm the same token works for an
  arbitrary conversation id without needing to have DOM-opened it first.
- **Whether this generalizes past chat threads to channels** — this account's
  discovery only ever exercised chat conversations (see `selectors.ts`'s own note that
  channels were never verified even for the DOM path).

## Bottom line

Real and worth pursuing, but a materially bigger scope than "read messages instead of
scraping the DOM" — touches attachment handling, sync semantics, and edit/delete
handling that the current DOM+push architecture already solves. Should be scoped as
its own piece of work, not folded into an unrelated change.
