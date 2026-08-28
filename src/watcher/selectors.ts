/**
 * SELECTORS — verified against a live modern Teams web client (v2,
 * teams.microsoft.com), Fluent UI-based, 2026-08. Confirmed via the app's
 * own /api/debug/find and /api/debug/query endpoints against a real
 * session (see README's "Verify the selectors" section for how to
 * re-verify if Teams' markup changes again).
 *
 * Two things NOT verified live (no example available at the time):
 * - messageEditedMarkerIdPrefix: inferred from the aria-labelledby
 *   reference pattern (author-<mid>, timestamp-<mid> were both confirmed
 *   to follow this convention), but no edited message was on hand to
 *   confirm edited-<mid> itself appears/contains text.
 * - deletedMessageTextPatterns: no deleted message was on hand; body text
 *   is matched against known Teams phrasing as a heuristic backstop.
 *
 * Author/timestamp/edited elements are NOT DOM descendants of their
 * message (Teams groups consecutive messages from one sender under a
 * single shared header), so they're looked up by id (author-<mid>,
 * timestamp-<mid>, edited-<mid>) rather than queried as descendants.
 */
export const selectors = {
  // Left rail: chat/channel entries. Channels were NOT verified — this
  // account's left rail only had chats open during verification. If
  // channel tracking doesn't discover anything, check whether channel
  // items use a different data-testid/data-item-type here.
  conversationListItem: '[role="treeitem"][data-testid="list-item"]',

  // Scrollable container for the open thread's messages.
  messagePane: '[data-tid="message-pane-list-runway"]',

  // One message. Its id is in the data-mid attribute (a numeric string).
  //
  // Scoped to descend from messagePane deliberately: confirmed live that
  // an UNSCOPED `[data-tid="chat-pane-message"]` query matches elsewhere
  // on the page too — almost certainly a floating notification/preview
  // toast Teams renders using the same message markup, elsewhere in the
  // DOM, for whichever conversation just received a message, regardless
  // of which conversation is actually open. Querying page-wide picked
  // that up and fed it through as if it belonged to the open conversation
  // — the real cause of messages showing up under the wrong conversation
  // (not the render-timing race that was fixed alongside this).
  messageListItem: '[data-tid="message-pane-list-runway"] [data-tid="chat-pane-message"]',
  messageBody: '[data-message-content]',

  // A file/image attachment renders as a SIBLING of messageBody, not a
  // descendant of it — messageBody is empty for an attachment-only
  // message. Each attachment item carries `numberoffiles` regardless of
  // its actual value, which is a convenient marker to select each one
  // individually. Two confirmed-live shapes for what's inside an item,
  // and they're NOT interchangeable:
  // - A document (confirmed live: a JSON file uploaded to SharePoint) has
  //   a nested content-card-custom-title element whose aria-label packs
  //   both filename and URL as "<name>\n<url>".
  // - An actual image (confirmed live: sent a real PNG via the compose
  //   box's attach button and captured the resulting DOM directly) has NO
  //   content-card-custom-title at all — instead a real `<img src=...>`
  //   thumbnail. Its `src` is a page-local `blob:` URL (only resolves
  //   inside the exact tab that created it, useless anywhere else), and
  //   `attachmentPreviewUrlAttr` on the item gives a real, stable HTTPS
  //   URL — but that one 401s without an authenticated Teams session
  //   (confirmed live via a plain curl), which this dashboard's own
  //   viewer generally won't have. Neither URL is what actually makes
  //   images work: the blob's *bytes* are read directly out of the page
  //   (see teams.ts) and saved locally (see attachments.ts) — that's the
  //   part that sidesteps the auth problem entirely, since the browser
  //   already fetched and decoded them just to render the thumbnail.
  //   attachmentPreviewUrlAttr is kept only as the attachment's `url`
  //   field for cases where no image bytes were captured (e.g. the
  //   in-page fetch of the blob itself failed) — a link that needs Teams
  //   auth to open is still better than no reference to the file at all.
  // Extraction tries attachmentPreviewUrlAttr FIRST (present only for a
  // real image), then falls back to the document-card shape, then to the
  // item's own aria-label (just the filename, no URL) if neither matched.
  attachmentGrid: '[data-tid="file-attachment-grid"]',
  attachmentItem: "[numberoffiles]",
  attachmentTitle: '[data-testid="content-card-custom-title"]',
  attachmentPreviewUrlAttr: "amspreviewurl",

  // A quoted/replied-to message preview, confirmed live via a self-quote
  // in "48:notes" ("Van Dang Thai" quoting their own earlier "x" + an
  // attachment, replying "1"). Structural, not a text-pattern guess like
  // the original heuristic this replaced: quotedReplyTimestamp and
  // quotedReplyContent are both explicit data-tid elements. There's no
  // dedicated data-tid for the author specifically — it's the timestamp
  // element's previousElementSibling within the same flex row (confirmed
  // in the same capture: `<span>Van Dang Thai</span><span
  // data-tid="quoted-reply-timestamp">...`). Also confirmed live: the
  // timestamp span can be present in the DOM but invisible (didn't show up
  // in innerText at all for this self-quote, cause unconfirmed — possibly
  // a width/breakpoint thing specific to your-own-message styling) — worth
  // knowing if `quotedReplyTimestamp` ever comes back empty despite the
  // element existing. If a quoted message itself had an image/file, the
  // preview here shows only a generic 📄/📷 placeholder GLYPH baked into
  // quotedReplyContent's own text — Teams does not expose the real
  // attachment (name/URL) from within the quoting message's DOM at all;
  // that only exists on the original message's own row, if it was
  // captured separately.
  quotedReplyCard: '[data-tid="quoted-reply-card"]',
  quotedReplyTimestamp: '[data-tid="quoted-reply-timestamp"]',
  quotedReplyContent: '[data-tid="quoted-reply-preview-content"]',

  // A person mention: Teams wraps each WORD of a multi-word display name
  // in its own sibling element sharing the same data-person-mri (confirmed
  // live: "Tung Nguyen Ba" is three separate elements, joined by literal
  // &nbsp; text nodes between them, not one element with the full name) —
  // extraction groups consecutive siblings sharing the same mri back into
  // one name. An @Everyone mention carries the same attribute on a single,
  // unsplit element, so the same grouping logic handles it for free.
  mentionAttr: "data-person-mri",

  // Message compose box (confirmed live: a real contenteditable div, one
  // per open conversation) — used to send a message from the dashboard.
  composeBox: '[data-tid="ckeditor"]',

  messageAuthorIdPrefix: "author-",
  messageTimestampIdPrefix: "timestamp-",
  messageEditedMarkerIdPrefix: "edited-", // unverified, see note above

  deletedMessageTextPatterns: [/this message (was|has been) deleted/i],
} as const;
