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

function extractTrailers(content) {
  const bodyLines = [];
  const attachments = [];
  let mentionNames = [];
  let structuredQuote = null;
  for (const line of (content ?? "").split("\n")) {
    const trimmed = line.trim();
    const attMatch = ATTACHMENT_LINE_RE.exec(trimmed);
    const mentionMatch = MENTIONS_LINE_RE.exec(trimmed);
    const quoteMatch = QUOTE_LINE_RE.exec(trimmed);
    if (attMatch) attachments.push({ name: attMatch[1], url: attMatch[2] });
    else if (mentionMatch) mentionNames = mentionMatch[1].split("|").filter(Boolean);
    else if (quoteMatch) {
      structuredQuote = { quotedAuthor: quoteMatch[1], quotedTimestamp: quoteMatch[2], quotedText: quoteMatch[3] };
    } else bodyLines.push(line);
  }
  return { text: bodyLines.join("\n").trim(), attachments, mentionNames, structuredQuote };
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

// asLinks=false (the message LIST) renders every chip as an inert <div> —
// no <a>, no target=_blank — so a click anywhere on the message, image
// included, does exactly one thing: select it and open the detail panel.
// That used to be two competing actions on one click (open the history
// popup AND open the image in a new tab), which was the actual complaint.
// asLinks=true is used only for non-image files inside the now-open
// detail panel, where a click is an explicit, unambiguous "open this
// file" action with nothing else listening on the same click.
function renderAttachments(attachments, asLinks) {
  if (!attachments.length) return "";
  const chips = attachments.map((a) => {
    const name = escapeHtml(a.name);
    const url = escapeHtml(a.url);
    const icon = `<span class="attachment-icon">📎</span><span class="attachment-name">${name}</span>`;
    const isImage = IMAGE_EXT_RE.test(a.name) && a.url;
    const tag = asLinks && a.url ? "a" : "div";
    const linkAttrs = tag === "a" ? ` href="${url}" target="_blank" rel="noopener"` : "";
    if (isImage) {
      // Normally a.url is our own /api/attachments/... route serving a
      // locally-downloaded copy of the image's actual bytes (see
      // watcher/attachments.ts) — no auth needed, always loads. The
      // onerror fallback exists for the rarer case where no local
      // snapshot exists (e.g. the in-page blob read failed) and a.url is
      // a remote Teams link needing an authenticated session this
      // dashboard's viewer may not have — falls back to a plain link chip
      // rather than a broken-image icon.
      return `<${tag}${linkAttrs} class="attachment-chip attachment-image">
        <img src="${url}" alt="${name}" onerror="this.parentElement.classList.add('broken'); this.remove()" />
        ${icon}
      </${tag}>`;
    }
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
  const { text, attachments, mentionNames, structuredQuote } = extractTrailers(content);
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
    const bodyHtml = text ? `<div class="message-body">${highlightMentions(escapeHtml(text), mentionNames)}</div>` : "";
    return bodyHtml + attachmentsHtml + (pending ? renderImagesLoadingNote() : "");
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
  detailPanel.hidden = true;
  resizer.hidden = true;
  document.querySelectorAll(".message.selected").forEach((el) => el.classList.remove("selected"));
}
document.getElementById("detailClose").onclick = closeDetailPanel;

async function showMessageDetail(messageId) {
  selectedMessageId = messageId;
  detailPanel.hidden = false;
  resizer.hidden = false;
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

  detailBody.innerHTML = `
    <div class="detail-meta">
      <span class="message-sender" style="color:${colorForSender(listEntry.sender)}">${escapeHtml(listEntry.sender ?? "Unknown")}</span>
      <span class="message-time">${new Date(latest.captured_at).toLocaleString()}</span>
      ${latest.edited_flag ? '<span class="edited-tag">(edited)</span>' : ""}
    </div>
    ${renderBody(latest.content, !!latest.is_deleted, { includeAttachments: false, pending: pendingMessageIds.has(messageId) })}
    ${renderImageGallerySection(images)}
    ${files.length ? renderAttachments(files, true) : ""}
    <div class="detail-history${historyExpanded ? " expanded" : ""}${historyExpanded && historyHeightPx ? " custom-height" : ""}">
      <div class="detail-history-resizer"${historyExpanded ? "" : " hidden"} title="Drag to resize"></div>
      <button type="button" class="detail-history-toggle${historyExpanded ? " expanded" : ""}">
        <span class="detail-history-chevron">▶</span>History (${revisions.length})
      </button>
      <div class="revision-list"${historyExpanded ? "" : " hidden"} style="${historyHeightPx ? `height:${historyHeightPx}px` : ""}">${historyHtml}</div>
    </div>
  `;

  wireImageGallery(detailBody, images);
  wireHistorySection(detailBody);
}

// Collapsed by default, expanded state remembered across message switches
// (a whole panel rebuild — not a plain DOM show/hide — since
// showMessageDetail replaces detailBody.innerHTML wholesale each time) so
// re-selecting a message or a live-update refresh doesn't quietly re-close
// something you just opened.
//
// Sizing is accordion-style (VS Code's sidebar sections are the reference
// point) BY DEFAULT: collapsed means just the header, pushed to the bottom
// of the panel (see .detail-history's margin-top:auto in style.css);
// freshly expanded means the section claims all remaining space down to
// the bottom of the panel (.detail-history.expanded { flex: 1 1 auto }).
// On top of that default, the section is still manually resizable via the
// drag handle at its top edge — dragging shrinks it away from "fill
// everything" to an explicit height, while margin-top:auto keeps it
// docked to the bottom either way, so shrinking opens up space ABOVE it
// rather than leaving a gap below. That explicit height (once set) is
// remembered across messages the same way historyExpanded is.
let historyExpanded = false;
let historyHeightPx = null;
try {
  historyHeightPx = Number(localStorage.getItem("detailHistoryHeight")) || null;
} catch {
  // Private browsing / storage disabled — just falls back to "fill" every time.
}

// Drag state lives at module scope, and its mousemove/mouseup listeners
// are attached exactly ONCE below (not inside wireHistorySection) — that
// function re-runs on every single message click, since showMessageDetail
// replaces detailBody.innerHTML each time, and attaching fresh
// window-level listeners on every one of those calls without ever
// removing the old ones would leak a new listener per click forever. The
// resizer bar and section are fresh elements each render, so wiring their
// own mousedown/class directly each time is fine.
let historyResizeSection = null;
let historyResizeBody = null;
let historyResizeBar = null;
let historyResizeStartY = 0;
let historyResizeStartHeight = 0;

window.addEventListener("mousemove", (e) => {
  if (!historyResizeBody) return;
  // The handle sits on the section's TOP edge — dragging UP (clientY
  // decreasing) has to GROW the box (pulling the boundary up claims more
  // room from the sections above) and dragging DOWN has to shrink it.
  const height = Math.min(
    Math.max(historyResizeStartHeight - (e.clientY - historyResizeStartY), 80),
    window.innerHeight * 0.7
  );
  historyResizeBody.style.height = `${height}px`;
});
window.addEventListener("mouseup", () => {
  if (!historyResizeBody) return;
  historyHeightPx = historyResizeBody.offsetHeight;
  try {
    localStorage.setItem("detailHistoryHeight", String(historyHeightPx));
  } catch {
    // Fine to just not persist it.
  }
  historyResizeBar?.classList.remove("dragging");
  historyResizeSection = null;
  historyResizeBody = null;
  historyResizeBar = null;
  document.body.style.userSelect = "";
});

function wireHistorySection(container) {
  const section = container.querySelector(".detail-history");
  const toggle = container.querySelector(".detail-history-toggle");
  const body = container.querySelector(".revision-list");
  const bar = container.querySelector(".detail-history-resizer");
  if (!section || !toggle || !body) return;

  toggle.onclick = () => {
    historyExpanded = !historyExpanded;
    section.classList.toggle("expanded", historyExpanded);
    section.classList.toggle("custom-height", historyExpanded && !!historyHeightPx);
    toggle.classList.toggle("expanded", historyExpanded);
    body.hidden = !historyExpanded;
    if (bar) bar.hidden = !historyExpanded; // resizing a collapsed section makes no sense
  };

  if (bar) {
    bar.onmousedown = (e) => {
      historyResizeSection = section;
      historyResizeBody = body;
      historyResizeBar = bar;
      historyResizeStartY = e.clientY;
      historyResizeStartHeight = body.offsetHeight;
      // Switch off the default "fill everything" flex right away, not on
      // mouseup — otherwise .detail-history stays flex:1 (fill) for the
      // whole drag, which fights every height we set on mousemove (its
      // flex-grow just snaps the list back to fill the full box every
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
