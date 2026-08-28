import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { config } from "../config.js";
import { getAttachmentSnapshot, saveAttachmentSnapshot, type AttachmentSnapshot } from "../db/index.js";

fs.mkdirSync(config.attachmentsDir, { recursive: true });

const EXT_BY_CONTENT_TYPE: Record<string, string> = {
  "image/png": ".png",
  "image/jpeg": ".jpg",
  "image/gif": ".gif",
  "image/webp": ".webp",
  "image/bmp": ".bmp",
  "image/svg+xml": ".svg",
};

function extFor(contentType: string): string {
  return EXT_BY_CONTENT_TYPE[contentType] ?? "";
}

export function localFilePath(snap: Pick<AttachmentSnapshot, "content_hash" | "content_type">): string {
  return path.join(config.attachmentsDir, `${snap.content_hash}${extFor(snap.content_type)}`);
}

/**
 * An in-page-captured image, as a data URL ("data:image/png;base64,...") —
 * see teams.ts's blob-to-dataURL conversion. Reading it this way (straight
 * out of the blob the browser already fetched and decoded to render the
 * thumbnail) sidesteps needing our own authenticated request to Teams'
 * media gateway entirely — confirmed live that a plain request to that
 * URL 401s without a Teams session, and Playwright's page.request/an
 * in-page fetch with credentials don't carry whatever auth the page's own
 * JS used internally (same problem as chatsvc.ts's dead REST path).
 */
export interface CapturedImage {
  name: string;
  dataUrl: string;
}

const DATA_URL_RE = /^data:([^;]+);base64,(.+)$/;

/**
 * Builds the final `[[ATTACHMENT|displayName|url]]` marker for one
 * attachment, saving its bytes (if any were captured) to a local snapshot
 * first — split out of resolveAttachmentMarkers below so it can be reused
 * for both an in-place replace (a pasted image with a known position — see
 * PENDING_ATTACHMENT below) and a trailing append (everything else).
 *
 * Safe to call on every visit: saving is idempotent (ON CONFLICT DO
 * NOTHING at the DB layer, and the file write is skipped outright if a
 * file for that content hash already exists), so re-capturing the same
 * image on a later revisit is a cheap no-op, not a repeat download.
 */
function buildMarker(messageId: string, att: { name: string; url: string }, dataUrl: string | undefined): string {
  if (!dataUrl) return `[[ATTACHMENT|${att.name}|${att.url}]]`;

  const existing = getAttachmentSnapshot(messageId, att.name);
  let snap: AttachmentSnapshot;
  if (existing) {
    snap = existing;
  } else {
    const match = DATA_URL_RE.exec(dataUrl);
    if (!match) {
      // Malformed data URL — fall back to the (unloadable) remote url
      // rather than lose the attachment reference entirely.
      return `[[ATTACHMENT|${att.name}|${att.url}]]`;
    }
    const [, contentType, base64] = match;
    const bytes = Buffer.from(base64, "base64");
    const contentHash = crypto.createHash("sha256").update(bytes).digest("hex");
    snap = {
      message_id: messageId,
      name: att.name,
      content_hash: contentHash,
      content_type: contentType,
      byte_size: bytes.length,
      captured_at: new Date().toISOString(),
    };
    const filePath = localFilePath(snap);
    if (!fs.existsSync(filePath)) fs.writeFileSync(filePath, bytes);
    saveAttachmentSnapshot(snap);
  }

  // The DISPLAYED name needs a real extension for the frontend's image
  // detection (IMAGE_EXT_RE in app.js) to recognize it as previewable —
  // a directly-pasted image (see teams.ts) has no real filename to begin
  // with, unlike an uploaded file, so one's synthesized here from the
  // actual detected content type. The STORAGE key (DB row + URL path
  // segment below) deliberately stays att.name as-is regardless, so a
  // repeat visit's getAttachmentSnapshot lookup above keeps finding the
  // same row instead of re-saving under a shifting key every time.
  const displayName = /\.[a-z0-9]+$/i.test(att.name) ? att.name : `${att.name}${extFor(snap.content_type)}`;
  return `[[ATTACHMENT|${displayName}|/api/attachments/${encodeURIComponent(messageId)}/${encodeURIComponent(att.name)}]]`;
}

/**
 * Resolves every attachment into its final marker and folds it back into
 * `content` — either IN PLACE, replacing the `[[PENDING_ATTACHMENT|name]]`
 * placeholder teams.ts left at the image's actual position in the text
 * (the plain-text pasted-image case, so it renders where it was pasted
 * instead of always trailing all the text — see app.js's renderBody), or,
 * for anything that has no such placeholder (uploaded files/images from
 * the file-attachment grid, which Teams itself never interleaves with
 * text either; a quoted reply's own pasted image, not yet given the same
 * positional treatment — see teams.ts), appended at the end same as
 * before this existed.
 */
export function resolveAttachmentMarkers(
  content: string,
  messageId: string,
  attachments: { name: string; url: string }[],
  images: CapturedImage[]
): string {
  const imageByName = new Map(images.map((img) => [img.name, img.dataUrl]));

  let result = content;
  let appended = "";
  for (const att of attachments) {
    const marker = buildMarker(messageId, att, imageByName.get(att.name));
    const placeholder = `[[PENDING_ATTACHMENT|${att.name}]]`;
    if (result.includes(placeholder)) {
      result = result.split(placeholder).join(marker);
    } else {
      appended += `\n${marker}`;
    }
  }
  return result + appended;
}
