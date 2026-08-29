let activeConversationId = null;
// The message currently shown in the detail panel (see showMessageDetail),
// or null when it's closed. Kept so a live update can refresh the open
// panel in place, and so loadMessages can re-apply the .selected
// highlight after it rebuilds the whole list from scratch.
let selectedMessageId = null;
// Purely transient (never persisted — the backend only ever emits this live,
// there's no DB column for it), so it's plain in-memory state here, rebuilt
// into each tab's class list on every loadTabs() re-render.
const syncingIds = new Set();
// Same idea, at message granularity: message ids the backend has flagged as
// "the latest revision we recorded might be incomplete" (see
// pendingResolutionMessagesByConversation in watcher/index.ts) — most
// commonly an edit that dropped attachments because push can't carry them,
// waiting on the backup DOM visit to reconcile. Drives renderImagesLoadingNote
// in renderBody below.
const pendingMessageIds = new Set();

const tabList = document.getElementById("tabList");
const emptyState = document.getElementById("emptyState");
const conversationView = document.getElementById("conversationView");
const conversationTitle = document.getElementById("conversationTitle");
const messageList = document.getElementById("messageList");
const resizer = document.getElementById("resizer");
const detailPanel = document.getElementById("detailPanel");
const detailBody = document.getElementById("detailBody");

async function fetchJSON(url, opts) {
  const res = await fetch(url, opts);
  if (!res.ok) throw new Error(`${url} -> ${res.status}`);
  return res.json();
}

async function loadTabs() {
  const conversations = await fetchJSON("/api/conversations");
  // Pinned conversations first; within each group, most recently active first.
  conversations.sort((a, b) => {
    if (a.pinned !== b.pinned) return b.pinned - a.pinned;
    if (a.last_activity_at || b.last_activity_at) {
      return (b.last_activity_at ?? "").localeCompare(a.last_activity_at ?? "");
    }
    return a.title.localeCompare(b.title);
  });
  tabList.innerHTML = "";
  for (const conv of conversations) {
    const unread = isUnread(conv);
    const syncing = syncingIds.has(conv.conversation_id);
    const el = document.createElement("div");
    el.dataset.conversationId = conv.conversation_id;
    el.className =
      "tab" +
      (conv.conversation_id === activeConversationId ? " active" : "") +
      (unread ? " unread" : "") +
      (syncing ? " syncing" : "");
    el.innerHTML = `
      <div class="tab-row">
        <div class="tab-title"><span class="sync-spinner" title="Syncing…"></span>${unread ? '<span class="unread-dot"></span>' : ""}${escapeHtml(conv.title)}</div>
        <button class="pin-btn${conv.pinned ? " pinned" : ""}" title="${conv.pinned ? "Unpin" : "Pin"}">${conv.pinned ? "★" : "☆"}</button>
      </div>
      <div class="tab-kind">${conv.kind}</div>
    `;
    el.onclick = () => selectConversation(conv.conversation_id, conv.title);
    el.querySelector(".pin-btn").onclick = async (e) => {
      e.stopPropagation();
      await fetchJSON(`/api/conversations/${encodeURIComponent(conv.conversation_id)}/pin`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ pinned: !conv.pinned }),
      });
      loadTabs();
    };
    tabList.appendChild(el);
  }
}

function isUnread(conv) {
  return !!conv.last_activity_at && (!conv.last_read_at || conv.last_activity_at > conv.last_read_at);
}

async function markRead(id) {
  try {
    await fetchJSON(`/api/conversations/${encodeURIComponent(id)}/read`, { method: "POST" });
  } catch (err) {
    console.error("mark read failed:", err);
  }
}

async function selectConversation(id, title) {
  activeConversationId = id;
  emptyState.hidden = true;
  conversationView.hidden = false;
  // Shown as soon as a conversation is open, not just once a message is
  // selected — otherwise there's nothing to drag until you've clicked a
  // message once, which is the awkward part this is fixing.
  detailPanel.hidden = false;
  resizer.hidden = false;
  conversationTitle.textContent = title;
  // A draft in progress for the PREVIOUS conversation shouldn't silently
  // carry over and get sent to this one instead.
  composeInput.value = "";
  composeInput.style.height = "auto";
  composeError.hidden = true;
  closeDetailPanel(); // a message from the PREVIOUS conversation's open panel wouldn't make sense here
  await markRead(id);
  await loadTabs();
  await loadMessages(id);
}

// Kept between calls so showMessageDetail's header (sender, timestamp,
// deleted state) doesn't need its own fetch — loadMessages already has it.
let currentMessages = [];

async function loadMessages(id) {
  const messages = await fetchJSON(`/api/conversations/${encodeURIComponent(id)}/messages`);
  currentMessages = messages;
  messageList.innerHTML = "";
  let stillSelected = false;
  for (const m of messages) {
    const el = document.createElement("div");
    el.dataset.messageId = m.message_id;
    const isSelected = m.message_id === selectedMessageId;
    if (isSelected) stillSelected = true;
    el.className = "message" + (m.is_deleted ? " deleted" : "") + (isSelected ? " selected" : "");
    el.innerHTML = `
      <div class="message-meta">
        <span class="message-sender" style="color:${colorForSender(m.sender)}">${escapeHtml(m.sender ?? "Unknown")}</span>
        <span class="message-time">${new Date(m.captured_at).toLocaleString()}</span>
        ${m.edited_flag ? '<span class="edited-tag">(edited)</span>' : ""}
      </div>
      ${renderBody(m.content, !!m.is_deleted, { pending: pendingMessageIds.has(m.message_id) })}
    `;
    // Anywhere on the message — text or an image inside it — opens the
    // detail panel. Images render with no <a> wrapper of their own now
    // (see renderAttachments), so nothing here steals the click into
    // also opening a new tab, which was the actual complaint: one click
    // used to fire two competing actions.
    el.onclick = () => showMessageDetail(m.message_id);
    messageList.appendChild(el);
  }
  messageList.scrollTop = messageList.scrollHeight;

  // The open panel (if any) belongs to a message that just got a fresh
  // revision — most commonly a pending image finishing resolution — so
  // refresh it in place rather than leaving it showing stale content.
  // Gone from the list entirely (rare — e.g. conversation switched under
  // it) closes it instead of refreshing something that no longer exists.
  if (selectedMessageId) {
    if (stillSelected) showMessageDetail(selectedMessageId);
    else closeDetailPanel();
  }
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str ?? "";
  return div.innerHTML;
}

// Deterministic per-sender color (same hashing approach as the backend's
// jitterFor in watcher/index.ts) so the same person is always the same
// color everywhere — message list, quote previews, revision history — and
// becomes recognizable at a glance in a multi-person log, not just by
// re-reading the name every time. Picked for contrast against the dark
// (#1e1e1e) background, distinct enough from each other and from the
// #0078d4 accent blue used for UI chrome.
const SENDER_COLORS = ["#4fc3f7", "#81c784", "#ffb74d", "#e57373", "#ba68c8", "#4dd0c4", "#f06292", "#9fa8da"];
function colorForSender(name) {
  if (!name) return "#999";
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = (hash * 31 + name.charCodeAt(i)) >>> 0;
  return SENDER_COLORS[hash % SENDER_COLORS.length];
}

// Teams renders a "replying to" preview as its own block above the actual
// reply text. The extraction now preserves that block boundary as a line
// break in the stored plain text instead of mashing both blocks together
// with zero separator (see teams.ts's extractVisibleMessages and
// chatsvc.ts's htmlToText — that was the original bug: a stored message
// had its quoted text and reply text fused mid-word, e.g. "rồivụ").
//
// The exact DOM/HTML structure Teams uses for that quote sub-block was
// never confirmed live — the watcher was actively running against the
// real account while this was investigated, so a second automated
// session couldn't safely attach to inspect it without disrupting live
// monitoring, and guessing a selector risked silently matching nothing.
// So instead of a schema column populated at capture time, the quote is
// re-detected here, at render time, straight off the flattened text
// pattern already confirmed in captured data: a quoted header line reads
// "<Quoted Author><timestamp><optional icon><quoted text>", glued
// together with little or no whitespace, followed by a line break and
// then the real reply text.
//
// Deliberately conservative: requires a capitalized, name-shaped phrase
// (1-5 title-case words) immediately before the timestamp, and requires
// non-empty text after the first line. An ordinary multi-line message
// that happens to start with a name and a date could still misfire as a
// "quote box" — a cosmetic false positive, not a data problem — so this
// stays a heuristic, not a guarantee.
const NAME_RE = /^[\p{Lu}][\p{L}'.-]*(?:\s+[\p{Lu}][\p{L}'.-]*){0,4}$/u;
const TIMESTAMP_RE = /^\d{1,2}\/\d{1,2}\/\d{4},?\s+\d{1,2}:\d{2}(?::\d{2})?\s*[AP]M$/;
const QUOTE_HEADER_RE = new RegExp(
  `^(${NAME_RE.source.slice(1, -1)})\\s*(${TIMESTAMP_RE.source.slice(1, -1)})\\s*(.*)$`,
  "u"
);
const stripLeadingIcon = (s) => s.replace(/^[\u{1F000}-\u{1FFFF}☀-➿\s]+/u, "").trim();

function splitQuotedReply(content) {
  const lines = (content ?? "").split("\n");
  if (lines.length < 2) return null;

  // DOM-extracted messages (extractVisibleMessages' innerText, as opposed
  // to the push path's htmlToText) put the quoted author and timestamp on
  // their OWN separate lines rather than fused onto one — Teams' quote
  // sub-elements are apparently separate enough in the live layout that
  // innerText breaks between every one of them, not just around the whole
  // quote block. Confirmed live: a self-healed re-capture of a previously
  // fused message came back as "Tung Nguyen Ba" / "8/27/2026 2:43 PM" /
  // "<quoted text>" / "" / "<reply text>" — three lines for the header
  // alone, then a blank line before the real reply. Handled as its own
  // case rather than trying to unify with the single-line push format
  // below, since the two really are shaped differently at the source.
  if (NAME_RE.test(lines[0].trim()) && lines[1] && TIMESTAMP_RE.test(lines[1].trim())) {
    const blankIdx = lines.findIndex((l, i) => i >= 2 && l.trim() === "");
    if (blankIdx !== -1) {
      const quotedText = stripLeadingIcon(lines.slice(2, blankIdx).join("\n"));
      const replyText = lines.slice(blankIdx + 1).join("\n").trim();
      if (replyText) {
        return { quotedAuthor: lines[0].trim(), quotedTimestamp: lines[1].trim(), quotedText, replyText };
      }
    }
  }

  // Push-payload messages (chatsvc.ts's htmlToText): author, timestamp, and
  // quoted text all fused onto the first line with no internal separator,
  // reply text starting on the next line.
  const match = QUOTE_HEADER_RE.exec(lines[0]);
  if (!match) return null;
  const [, quotedAuthor, quotedTimestamp, quotedTextFirstLine] = match;
  const replyText = lines.slice(1).join("\n").trim();
  if (!replyText) return null; // nothing after the header — don't force a split
  return {
    quotedAuthor: quotedAuthor.trim(),
    quotedTimestamp,
    quotedText: stripLeadingIcon(quotedTextFirstLine),
    replyText,
  };
}

// Mentions, attachments, and (structural) quotes aren't structured DB
// columns — they're appended/prepended to the plain-text content as
// `[[...]]` marker lines by the extraction (teams.ts), which pulled them
// directly off real DOM elements rather than guessing, so no schema
// migration was needed to add any of them. Pulled off here, regardless of
// which line they're on, before anything else touches the text, so they
// never leak into displayed content.
const ATTACHMENT_LINE_RE = /^\[\[ATTACHMENT\|([^|]*)\|([^\]]*)\]\]$/;
const MENTIONS_LINE_RE = /^\[\[MENTIONS\|(.+)\]\]$/;
// Greedy (.*) deliberately grabs everything up to the LAST "]]" on the
// line, so a quoted text containing a literal "]]" doesn't truncate it —
// only author/timestamp (names/dates, never containing "|" in practice)
// are pipe-delimited.
const QUOTE_LINE_RE = /^\[\[QUOTE\|([^|]*)\|([^|]*)\|(.*)\]\]$/;

// `segments` (alongside the pre-existing `text`/`attachments`) is an
// ORDERED list of { type: "text", value } | { type: "image", att } pieces,
// rebuilt in the same order the lines appeared in `content` — which for an
// image is now typically wherever teams.ts's attachments.ts resolved its
// [[PENDING_ATTACHMENT|name]] placeholder to, i.e. the actual position the
// image was pasted at, not always the very end. renderBody uses this to
// show an inline image right where it was pasted (matching Teams' own
// layout) instead of every image being bundled into a chip row after all
// the text, which is what `attachments`+renderAttachments alone still
// gives you — kept for the callers (the detail panel, `text`-only
// contexts) that want the old "just the text, images separately" split.
function extractTrailers(content) {
  const bodyLines = [];
  const attachments = [];
  const segments = [];
  let textBuffer = [];
  let mentionNames = [];
  let structuredQuote = null;
  const flushTextBuffer = () => {
    if (textBuffer.length) {
      segments.push({ type: "text", value: textBuffer.join("\n") });
      textBuffer = [];
    }
  };
  for (const line of (content ?? "").split("\n")) {
    const trimmed = line.trim();
    const attMatch = ATTACHMENT_LINE_RE.exec(trimmed);
    const mentionMatch = MENTIONS_LINE_RE.exec(trimmed);
    const quoteMatch = QUOTE_LINE_RE.exec(trimmed);
    if (attMatch) {
      const att = { name: attMatch[1], url: attMatch[2] };
      attachments.push(att);
      // Only an IMAGE gets its own segment, positioned here — a non-image
      // file attachment has no meaningful position of its own (Teams
      // renders those as a grid regardless of surrounding text), so it's
      // left out of `segments` entirely and stays purely in `attachments`
      // for callers to render as a trailing chip row, same as before.
      if (IMAGE_EXT_RE.test(att.name) && att.url) {
        flushTextBuffer();
        segments.push({ type: "image", att });
      }
    } else if (mentionMatch) mentionNames = mentionMatch[1].split("|").filter(Boolean);
    else if (quoteMatch) {
      structuredQuote = { quotedAuthor: quoteMatch[1], quotedTimestamp: quoteMatch[2], quotedText: quoteMatch[3] };
    } else {
      bodyLines.push(line);
      textBuffer.push(line);
    }
  }
  flushTextBuffer();
  return { text: bodyLines.join("\n").trim(), attachments, segments, mentionNames, structuredQuote };
}

// Renders each mentioned name (already known exactly, extracted from a
// real DOM mention element — not guessed) as a colored "@Name" pill,
// matching that person's color everywhere else in the UI. Applied AFTER
// escapeHtml, so this replaces on already-safe text — the mention names
// themselves came through escapeHtml too, so a name can't inject markup.
function highlightMentions(escapedText, mentionNames) {
  let result = escapedText;
  for (const rawName of mentionNames) {
    const name = escapeHtml(rawName);
    if (!name) continue;
    const pattern = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`(^|\\s)(${pattern})(?=\\s|$)`, "g");
    result = result.replace(
      re,
      (_m, pre, matched) => `${pre}<span class="mention-pill" style="color:${colorForSender(rawName)}">@${matched}</span>`
    );
  }
  return result;
}

const IMAGE_EXT_RE = /\.(png|jpe?g|gif|webp|bmp|svg)$/i;

// asLinks=false (the message LIST, including a segment rendered inline by
// renderBody) renders the chip as an inert <div> — no <a>, no
// target=_blank — so a click anywhere on the message, image included,
// does exactly one thing: select it and open the detail panel. That used
// to be two competing actions on one click (open the history popup AND
// open the image in a new tab), which was the actual complaint.
// asLinks=true is used only for non-image files inside the now-open
// detail panel, where a click is an explicit, unambiguous "open this
// file" action with nothing else listening on the same click.
//
// Split out of renderAttachments below so renderBody can also call it
// per-segment, at an image's own position in the text, rather than only
// ever through renderAttachments' all-images-in-one-row rendering.
function renderImageChip(a, asLinks) {
  const name = escapeHtml(a.name);
  const url = escapeHtml(a.url);
  const icon = `<span class="attachment-icon">📎</span><span class="attachment-name">${name}</span>`;
  const tag = asLinks && a.url ? "a" : "div";
  const linkAttrs = tag === "a" ? ` href="${url}" target="_blank" rel="noopener"` : "";
  // Normally a.url is our own /api/attachments/... route serving a
  // locally-downloaded copy of the image's actual bytes (see
  // watcher/attachments.ts) — no auth needed, always loads. The onerror
  // fallback exists for the rarer case where no local snapshot exists
  // (e.g. the in-page blob read failed) and a.url is a remote Teams link
  // needing an authenticated session this dashboard's viewer may not
  // have — falls back to a plain link chip rather than a broken-image icon.
  return `<${tag}${linkAttrs} class="attachment-chip attachment-image">
    <img src="${url}" alt="${name}" onerror="this.parentElement.classList.add('broken'); this.remove()" />
    ${icon}
  </${tag}>`;
}

function renderAttachments(attachments, asLinks) {
  if (!attachments.length) return "";
  const chips = attachments.map((a) => {
    if (IMAGE_EXT_RE.test(a.name) && a.url) return renderImageChip(a, asLinks);
    const name = escapeHtml(a.name);
    const url = escapeHtml(a.url);
    const icon = `<span class="attachment-icon">📎</span><span class="attachment-name">${name}</span>`;
    const tag = asLinks && a.url ? "a" : "div";
    const linkAttrs = tag === "a" ? ` href="${url}" target="_blank" rel="noopener"` : "";
    return `<${tag}${linkAttrs} class="attachment-chip">${icon}</${tag}>`;
  });
  return `<div class="attachment-list">${chips.join("")}</div>`;
}

// Real-time push (trouter.ts) delivers a message the instant Teams would
// show it, but it can't extract attachments/images at all — that only
// happens when the DOM-based visit loop next opens the conversation, up
// to config.pool.backupCheckIntervalMs later (confirmed live: a message's
// first revision is empty content from push, its second — sometimes
// minutes later — has the resolved [[ATTACHMENT]] marker). Without this,
// a message with an image just looks blank/broken for that whole window,
// which is exactly what got reported: "the image suddenly appeared,
// thought it was a bug." Content with truly nothing in it (no quote, no
// text, no attachment) and a real sender is now a reliable "not resolved
// yet" signal — a genuine reaction/no-op push (empty content AND empty
// sender) is already filtered out entirely in trouter.ts and never
// reaches here at all.
function renderPendingPlaceholder() {
  return `<div class="message-pending"><span class="pending-spinner"></span>Fetching message content…</div>`;
}

// The narrower sibling of renderPendingPlaceholder above: this message
// already has SOME real content (text, at least) but the backend has
// explicitly flagged it as pending (see pendingMessageIds) — most often an
// edit that dropped attachments because push can't carry them. Reuses the
// same .message-pending/.pending-spinner styling for visual consistency,
// just with different wording and shown ALONGSIDE the real content instead
// of replacing it, since — unlike the fully-empty case above — there's
// already something real to show.
function renderImagesLoadingNote() {
  return `<div class="message-pending"><span class="pending-spinner"></span>Loading images…</div>`;
}

// includeAttachments=false is used by the detail panel, which renders
// images (a dedicated gallery) and non-image files (their own section)
// itself rather than inline with the text — see showMessageDetail.
function renderBody(content, isDeleted, { includeAttachments = true, pending = false } = {}) {
  const { text, attachments, segments, mentionNames, structuredQuote } = extractTrailers(content);
  // Quote branch below doesn't get the inline-at-position treatment (see
  // teams.ts/attachments.ts — a quoted reply's own pasted image isn't
  // given a placeholder there), so it still wants every attachment
  // trailing, same as before.
  const attachmentsHtml = includeAttachments ? renderAttachments(attachments, false) : "";

  // structuredQuote comes from teams.ts's DOM extraction, pulled directly
  // off Teams' own quoted-reply-card element — `text` here is ALREADY
  // just the reply (the quote was split off at capture time), so no
  // further splitting is needed. splitQuotedReply is a fallback for
  // content that never went through that path: messages captured before
  // this fix, or via the push payload (chatsvc.ts), which isn't
  // restructured the same way.
  const quoted = structuredQuote ?? splitQuotedReply(text);
  const replyText = structuredQuote ? text : quoted?.replyText;

  if (!quoted) {
    if (!text && !attachments.length && !isDeleted) return renderPendingPlaceholder();
    if (!includeAttachments) {
      // The detail panel: plain text only — it renders its own dedicated
      // image gallery and file section separately (see showMessageDetail),
      // so an image here would just show twice.
      const bodyHtml = text ? `<div class="message-body">${highlightMentions(escapeHtml(text), mentionNames)}</div>` : "";
      return bodyHtml + (pending ? renderImagesLoadingNote() : "");
    }
    // Inline images render AT THEIR ORIGINAL POSITION in the text —
    // `segments` (see extractTrailers) already carries them in the right
    // order, since teams.ts/attachments.ts resolve a pasted image's
    // placeholder in place rather than always trailing the text. A
    // non-image file attachment has no position of its own (Teams itself
    // renders those as a grid regardless of surrounding text), so those
    // still render as a trailing chip row, same as before this existed.
    const bodyHtml = segments.length
      ? `<div class="message-body">${segments
          .map((seg) =>
            seg.type === "image" ? renderImageChip(seg.att, false) : highlightMentions(escapeHtml(seg.value), mentionNames)
          )
          .join("")}</div>`
      : "";
    const fileAttachments = attachments.filter((a) => !(IMAGE_EXT_RE.test(a.name) && a.url));
    return bodyHtml + renderAttachments(fileAttachments, false) + (pending ? renderImagesLoadingNote() : "");
  }
  return `
    <div class="quote-block">
      <div class="quote-author">
        <span class="quote-icon">↪</span>
        <span style="color:${colorForSender(quoted.quotedAuthor)}">${escapeHtml(quoted.quotedAuthor)}</span>
        <span class="quote-timestamp">${escapeHtml(quoted.quotedTimestamp)}</span>
      </div>
      ${quoted.quotedText ? `<div class="quote-text">${highlightMentions(escapeHtml(quoted.quotedText), mentionNames)}</div>` : ""}
    </div>
    ${replyText ? `<div class="message-body">${highlightMentions(escapeHtml(replyText), mentionNames)}</div>` : ""}
    ${attachmentsHtml}
    ${pending ? renderImagesLoadingNote() : ""}
  `;
}

// --- Detail panel: replaces the old revision-history modal. Clicking a
// message (anywhere on it) opens this instead of a popup; the image
// gallery below is the other half of the same complaint — an image used
// to open a new tab AND the modal on the same click.

// Builds just the two image sections (filmstrip + big viewer) — returns
// "" when the message has no images, so callers can drop it in
// unconditionally.
function renderImageGallerySection(images) {
  if (!images.length) return "";
  const thumbs = images
    .map(
      (img, i) =>
        `<img class="detail-thumb${i === 0 ? " selected" : ""}" data-idx="${i}" src="${escapeHtml(img.url)}" alt="${escapeHtml(img.name)}" />`
    )
    .join("");
  return `
    <div class="detail-images">
      <div class="detail-thumbs">${thumbs}</div>
      <div class="detail-viewer">
        <img id="detailViewerImg" src="${escapeHtml(images[0].url)}" alt="${escapeHtml(images[0].name)}" />
      </div>
    </div>
  `;
}

// Wires the gallery's interactivity after it's in the DOM — done
// separately from rendering (rather than inline onclick= attributes)
// because it needs a live reference to the `images` array, not just
// strings safe to embed in an HTML attribute.
function wireImageGallery(container, images) {
  const viewerImg = container.querySelector("#detailViewerImg");
  if (!viewerImg) return;
  const thumbs = container.querySelectorAll(".detail-thumb");
  thumbs.forEach((thumb) => {
    thumb.onclick = () => {
      const img = images[Number(thumb.dataset.idx)];
      viewerImg.src = img.url;
      viewerImg.alt = img.name;
      viewerImg.classList.remove("zoomed"); // switching images resets any zoom on the previous one
      thumbs.forEach((t) => t.classList.toggle("selected", t === thumb));
    };
  });
  // Click-to-zoom: swaps the viewer image to its real size inside a
  // scrollable viewport instead of a fixed max — that's what makes it
  // actually zoomable (scroll around it) rather than just "shown bigger".
  viewerImg.onclick = () => viewerImg.classList.toggle("zoomed");
}

function closeDetailPanel() {
  selectedMessageId = null;
  // The panel itself stays visible (see selectConversation) so the
  // resizer stays draggable — this just clears its content back to a
  // placeholder instead of hiding the whole panel.
  detailBody.innerHTML = `<div class="empty-state">Select a message to view its details.</div>`;
  document.querySelectorAll(".message.selected").forEach((el) => el.classList.remove("selected"));
}
document.getElementById("detailClose").onclick = closeDetailPanel;

async function showMessageDetail(messageId) {
  selectedMessageId = messageId;
  for (const el of messageList.children) {
    el.classList.toggle("selected", el.dataset.messageId === messageId);
  }

  const revisions = await fetchJSON(`/api/messages/${encodeURIComponent(messageId)}/revisions`);
  if (selectedMessageId !== messageId) return; // a later click already moved the panel on

  const latest = revisions[revisions.length - 1];
  const listEntry = currentMessages.find((m) => m.message_id === messageId);
  if (!latest || !listEntry) {
    detailBody.innerHTML = `<div class="empty-state">No data for this message.</div>`;
    return;
  }

  const { attachments } = extractTrailers(latest.content);
  const images = attachments.filter((a) => IMAGE_EXT_RE.test(a.name) && a.url);
  const files = attachments.filter((a) => !IMAGE_EXT_RE.test(a.name));

  const historyHtml = revisions
    .slice()
    .reverse()
    .map(
      (r) => `
        <div class="revision-item">
          <div class="revision-meta">${new Date(r.captured_at).toLocaleString()} ${r.is_deleted ? "(deleted)" : r.edited_flag ? "(edited)" : ""}</div>
          ${renderBody(r.content, !!r.is_deleted)}
        </div>
      `
    )
    .join("");

  const attachmentCount = images.length + files.length;
  const attachmentsHtml = attachmentCount
    ? `${renderImageGallerySection(images)}${files.length ? renderAttachments(files, true) : ""}`
    : `<div class="empty-state">No attachments on this message.</div>`;

  detailBody.innerHTML = `
    <div class="detail-sections">
      ${renderDetailSection("attachments", "Attachments", attachmentCount, attachmentsHtml)}
      ${renderDetailSection("history", "History", revisions.length, historyHtml)}
    </div>
  `;

  wireImageGallery(detailBody, images);
  wireDetailSection(detailBody, "attachments");
  wireDetailSection(detailBody, "history");
}

// --- Generic collapsible/resizable detail-panel sections ("Attachments",
// "History", and whatever gets added next). Collapsed by default; expanded
// state and manual height are remembered across message switches (a whole
// panel rebuild — not a plain DOM show/hide — since showMessageDetail
// replaces detailBody.innerHTML wholesale each time) so re-selecting a
// message or a live-update refresh doesn't quietly re-close or re-size
// something the user just set.
//
// Sizing is accordion-style (VS Code's sidebar sections are the reference
// point) BY DEFAULT: collapsed means just the header, and the whole group
// of sections is docked to the bottom of the panel (see .detail-sections'
// justify-content:flex-end in style.css — that docks the GROUP, so several
// collapsed sections stack on one another rather than each fending for
// itself); freshly expanded means that section claims all remaining space
// down to the bottom of the panel (.detail-section.expanded { flex: 1 1
// auto }). On top of that default, a section is still manually resizable
// via the drag handle at its top edge — dragging shrinks it away from
// "fill everything" to an explicit height. That explicit height (once
// set) is persisted per section key; expanded/collapsed is only
// remembered for the lifetime of the page, matching the old history-only
// behavior.
//
// To add another section: pick a unique key, drop a
// renderDetailSection(key, title, count, bodyHtml) call inside the
// `.detail-sections` wrapper above, and add a matching
// wireDetailSection(detailBody, key) call alongside the two above.
const sectionState = new Map(); // key -> { expanded, heightPx }
function getSectionState(key) {
  let state = sectionState.get(key);
  if (!state) {
    let heightPx = null;
    try {
      heightPx = Number(localStorage.getItem(`detailSectionHeight:${key}`)) || null;
    } catch {
      // Private browsing / storage disabled — just falls back to "fill" every time.
    }
    state = { expanded: false, heightPx };
    sectionState.set(key, state);
  }
  return state;
}

function renderDetailSection(key, title, count, bodyHtml) {
  const state = getSectionState(key);
  return `
    <div class="detail-section${state.expanded ? " expanded" : ""}${state.expanded && state.heightPx ? " custom-height" : ""}" data-section-key="${key}">
      <div class="detail-section-resizer"${state.expanded ? "" : " hidden"} title="Drag to resize"></div>
      <button type="button" class="detail-section-toggle${state.expanded ? " expanded" : ""}">
        <span class="detail-section-chevron">▶</span>${title} (${count})
      </button>
      <div class="detail-section-body"${state.expanded ? "" : " hidden"} style="${state.heightPx ? `height:${state.heightPx}px` : ""}">${bodyHtml}</div>
    </div>
  `;
}

// Drag state lives at module scope, and its mousemove/mouseup listeners
// are attached exactly ONCE below (not inside wireDetailSection) — that
// function re-runs on every single message click, since showMessageDetail
// replaces detailBody.innerHTML each time, and attaching fresh
// window-level listeners on every one of those calls without ever
// removing the old ones would leak a new listener per click forever. The
// resizer bar and section are fresh elements each render, so wiring their
// own mousedown/class directly each time is fine.
let sectionResizeKey = null;
let sectionResizeBody = null;
let sectionResizeBar = null;
let sectionResizeStartY = 0;
let sectionResizeStartHeight = 0;

window.addEventListener("mousemove", (e) => {
  if (!sectionResizeBody) return;
  // The handle sits on the section's TOP edge — dragging UP (clientY
  // decreasing) has to GROW the box (pulling the boundary up claims more
  // room from the sections above) and dragging DOWN has to shrink it.
  const height = Math.min(
    Math.max(sectionResizeStartHeight - (e.clientY - sectionResizeStartY), 80),
    window.innerHeight * 0.7
  );
  sectionResizeBody.style.height = `${height}px`;
});
window.addEventListener("mouseup", () => {
  if (!sectionResizeBody) return;
  const state = getSectionState(sectionResizeKey);
  state.heightPx = sectionResizeBody.offsetHeight;
  try {
    localStorage.setItem(`detailSectionHeight:${sectionResizeKey}`, String(state.heightPx));
  } catch {
    // Fine to just not persist it.
  }
  sectionResizeBar?.classList.remove("dragging");
  sectionResizeKey = null;
  sectionResizeBody = null;
  sectionResizeBar = null;
  document.body.style.userSelect = "";
});

function wireDetailSection(container, key) {
  const section = container.querySelector(`.detail-section[data-section-key="${key}"]`);
  const toggle = section?.querySelector(".detail-section-toggle");
  const body = section?.querySelector(".detail-section-body");
  const bar = section?.querySelector(".detail-section-resizer");
  if (!section || !toggle || !body) return;

  const state = getSectionState(key);

  toggle.onclick = () => {
    state.expanded = !state.expanded;
    section.classList.toggle("expanded", state.expanded);
    section.classList.toggle("custom-height", state.expanded && !!state.heightPx);
    toggle.classList.toggle("expanded", state.expanded);
    body.hidden = !state.expanded;
    if (bar) bar.hidden = !state.expanded; // resizing a collapsed section makes no sense
  };

  if (bar) {
    bar.onmousedown = (e) => {
      sectionResizeKey = key;
      sectionResizeBody = body;
      sectionResizeBar = bar;
      sectionResizeStartY = e.clientY;
      sectionResizeStartHeight = body.offsetHeight;
      // Switch off the default "fill everything" flex right away, not on
      // mouseup — otherwise .detail-section stays flex:1 (fill) for the
      // whole drag, which fights every height we set on mousemove (its
      // flex-grow just snaps the body back to fill the full box every
      // frame) and the resize would only visibly "jump" once at the end.
      section.classList.add("custom-height");
      bar.classList.add("dragging");
      document.body.style.userSelect = "none";
      e.preventDefault(); // avoid text-selection drag artifacts while dragging
    };
  }
}

// --- Draggable divider between the message list and the detail panel.
// Width is clamped to the same min/max the panel's CSS already declares
// (min-width/max-width still apply as a hard floor/ceiling even if this
// ever computed something outside that range) and persisted so it stays
// put across reloads.
let resizingPanel = false;
resizer.addEventListener("mousedown", () => {
  resizingPanel = true;
  resizer.classList.add("dragging");
  document.body.style.userSelect = "none";
});
window.addEventListener("mousemove", (e) => {
  if (!resizingPanel) return;
  const width = Math.min(Math.max(window.innerWidth - e.clientX, 260), window.innerWidth * 0.7);
  detailPanel.style.width = `${width}px`;
});
window.addEventListener("mouseup", () => {
  if (!resizingPanel) return;
  resizingPanel = false;
  resizer.classList.remove("dragging");
  document.body.style.userSelect = "";
  try {
    localStorage.setItem("detailPanelWidth", detailPanel.style.width);
  } catch {
    // Private browsing / storage disabled — the width just won't persist.
  }
});
try {
  const savedWidth = localStorage.getItem("detailPanelWidth");
  if (savedWidth) detailPanel.style.width = savedWidth;
} catch {
  // Same as above — fine to just use the CSS default.
}

// --- Compose bar: send a real message through the watcher's own Teams
// session, exactly as if typed there. Sending is real and can't be undone
// once it goes through, so this only ever fires on an explicit user
// action (click, or Enter in the box) — nothing here sends on its own. ---
const composeInput = document.getElementById("composeInput");
const composeSend = document.getElementById("composeSend");
const composeError = document.getElementById("composeError");

function autoGrowCompose() {
  composeInput.style.height = "auto";
  composeInput.style.height = `${composeInput.scrollHeight}px`;
}
composeInput.addEventListener("input", autoGrowCompose);

async function sendCurrentMessage() {
  const text = composeInput.value;
  if (!text.trim() || !activeConversationId) return;

  composeInput.disabled = true;
  composeSend.disabled = true;
  composeError.hidden = true;

  try {
    const res = await fetch(`/api/conversations/${encodeURIComponent(activeConversationId)}/send`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ text }),
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(body.error || `Send failed (${res.status})`);

    composeInput.value = "";
    autoGrowCompose();
    // Push should also deliver this back within moments, but loading now
    // too means the sent message shows up immediately rather than waiting
    // on that round trip.
    await loadMessages(activeConversationId);
  } catch (err) {
    composeError.textContent = err instanceof Error ? err.message : String(err);
    composeError.hidden = false;
  } finally {
    composeInput.disabled = false;
    composeSend.disabled = false;
    composeInput.focus();
  }
}

composeSend.onclick = sendCurrentMessage;
composeInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    sendCurrentMessage();
  }
});

// --- Live updates ---
const authBanner = document.getElementById("authBanner");

// Connection dot reflects the WebSocket's actual state in real time — the
// most direct signal of whether the backend process is even reachable
// (as opposed to "auth required", which means the backend is up but Teams
// needs a re-login).
const connDot = document.getElementById("connDot");
const connText = document.getElementById("connText");
function setConnStatus(state, label) {
  connDot.className = "conn-dot " + state;
  connText.textContent = label;
}

function connectWebSocket() {
  setConnStatus("connecting", "Connecting…");
  const ws = new WebSocket(`ws://${location.host}`);
  ws.onopen = () => setConnStatus("connected", "Live");
  ws.onmessage = (event) => {
    const msg = JSON.parse(event.data);
    if (msg.type === "update") {
      if (msg.conversationId === activeConversationId) {
        loadMessages(activeConversationId);
        markRead(activeConversationId).then(loadTabs); // stays "read" while you're actively viewing it
      } else {
        loadTabs();
      }
    } else if (msg.type === "auth") {
      authBanner.hidden = msg.status !== "required";
    } else if (msg.type === "sync") {
      // Toggled directly on the existing tab element rather than going
      // through loadTabs() — this fires on every visit start/end (every few
      // seconds while the backup loop runs), and a full re-fetch+rebuild for
      // each one would be wasteful and would fight with the sort order
      // reshuffling under the user's cursor.
      if (msg.syncing) syncingIds.add(msg.conversationId);
      else syncingIds.delete(msg.conversationId);
      const el = tabList.querySelector(`[data-conversation-id="${msg.conversationId}"]`);
      if (el) el.classList.toggle("syncing", msg.syncing);
    } else if (msg.type === "messagePending") {
      if (msg.pending) pendingMessageIds.add(msg.messageId);
      else pendingMessageIds.delete(msg.messageId);
      // A full reload rather than patching the one row in place — same
      // approach loadMessages already takes for a genuine "update", and
      // this fires rarely enough (message edits/new-attachment messages,
      // not every push) that the simplicity is worth more than avoiding a
      // cheap local re-fetch. Only when the message is actually showing
      // right now, same as the "update" branch above.
      if (currentMessages.some((m) => m.message_id === msg.messageId)) {
        loadMessages(activeConversationId);
      }
    }
  };
  ws.onclose = () => {
    setConnStatus("disconnected", "Backend offline");
    setTimeout(connectWebSocket, 2000);
  };
}

loadTabs();
connectWebSocket();
