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
 * Saves any newly-captured images for this message that aren't already
 * snapshotted, then returns the `[[ATTACHMENT|name|url]]` marker line for
 * EVERY attachment (image or not) so callers don't need to track which
 * ones were images separately — non-image attachments just pass their
 * remote url through unchanged, exactly as before this existed.
 *
 * Safe to call on every visit: saving is idempotent (ON CONFLICT DO
 * NOTHING at the DB layer, and the file write is skipped outright if a
 * file for that content hash already exists), so re-capturing the same
 * image on a later revisit is a cheap no-op, not a repeat download.
 */
export function resolveAttachmentMarkers(
  messageId: string,
  attachments: { name: string; url: string }[],
  images: CapturedImage[]
): string {
  const imageByName = new Map(images.map((img) => [img.name, img.dataUrl]));

  let markers = "";
  for (const att of attachments) {
    const dataUrl = imageByName.get(att.name);
    if (!dataUrl) {
      markers += `\n[[ATTACHMENT|${att.name}|${att.url}]]`;
      continue;
    }

    const existing = getAttachmentSnapshot(messageId, att.name);
    let snap: AttachmentSnapshot;
    if (existing) {
      snap = existing;
    } else {
      const match = DATA_URL_RE.exec(dataUrl);
      if (!match) {
        // Malformed data URL — fall back to the (unloadable) remote url
        // rather than lose the attachment reference entirely.
        markers += `\n[[ATTACHMENT|${att.name}|${att.url}]]`;
        continue;
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
    markers += `\n[[ATTACHMENT|${displayName}|/api/attachments/${encodeURIComponent(messageId)}/${encodeURIComponent(att.name)}]]`;
  }
  return markers;
}
