/**
 * Builds the sideloadable Teams app package for the teams-bot connector.
 *
 * Substitutes your bot's app id into teams-app/manifest.template.json,
 * generates the two icons Teams requires (it rejects a package without
 * them), and writes teams-app/message-watcher.zip ready to upload via
 * Teams → Apps → Manage your apps → Upload a custom app.
 *
 *   node scripts/build-teams-app.mjs --bot-id <your-azure-bot-app-id>
 *   MICROSOFT_APP_ID=<id> node scripts/build-teams-app.mjs
 *
 * Deliberately dependency-free: PNG and ZIP are both written by hand below
 * rather than pulling in an image library and an archiver for a script that
 * runs approximately twice in this project's life.
 */

import fs from "node:fs";
import path from "node:path";
import zlib from "node:zlib";

const args = process.argv.slice(2);
const argValue = (name) => {
  const i = args.indexOf(`--${name}`);
  return i !== -1 ? args[i + 1] : undefined;
};

const BOT_ID = argValue("bot-id") ?? process.env.MICROSOFT_APP_ID ?? "";
const OUT_DIR = path.resolve("teams-app");

const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
if (!GUID_RE.test(BOT_ID)) {
  console.error(
    `Need your Azure Bot's Microsoft App ID (a GUID).\n` +
      (BOT_ID ? `Got: ${BOT_ID}\n` : "") +
      `\n  node scripts/build-teams-app.mjs --bot-id 00000000-0000-0000-0000-000000000000\n` +
      `\nFind it in the Azure portal under your Azure Bot resource → Configuration → Microsoft App ID.`
  );
  process.exit(1);
}

// ---------------------------------------------------------------- CRC32
// Shared by PNG chunks and ZIP entries.
const CRC_TABLE = (() => {
  const table = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c;
  }
  return table;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ------------------------------------------------------------------ PNG
function pngChunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const typeAndData = Buffer.concat([Buffer.from(type, "ascii"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typeAndData));
  return Buffer.concat([len, typeAndData, crc]);
}

/**
 * Minimal RGBA PNG. `pixel(x, y)` returns [r, g, b, a].
 *
 * Colour type 6 (RGBA), bit depth 8, no interlacing, filter 0 on every
 * scanline — the simplest thing every decoder handles.
 */
function writePng(size, pixel) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  let o = 0;
  for (let y = 0; y < size; y++) {
    raw[o++] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixel(x, y);
      raw[o++] = r;
      raw[o++] = g;
      raw[o++] = b;
      raw[o++] = a;
    }
  }

  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; // bit depth
  ihdr[9] = 6; // RGBA
  ihdr[10] = 0; // deflate
  ihdr[11] = 0; // adaptive filtering
  ihdr[12] = 0; // no interlace

  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk("IHDR", ihdr),
    pngChunk("IDAT", zlib.deflateSync(raw, { level: 9 })),
    pngChunk("IEND", Buffer.alloc(0)),
  ]);
}

// A rounded square with a speech bubble notch, in the manifest's accent
// colour. Teams wants: color.png 192x192 full-colour, outline.png 32x32
// single-colour on transparent.
const ACCENT = [47, 91, 140];

function roundedSquareAlpha(x, y, size, inset, radius) {
  const lo = inset;
  const hi = size - 1 - inset;
  if (x < lo || x > hi || y < lo || y > hi) return 0;
  const dx = Math.max(lo + radius - x, x - (hi - radius), 0);
  const dy = Math.max(lo + radius - y, y - (hi - radius), 0);
  return Math.hypot(dx, dy) <= radius ? 255 : 0;
}

const colorIcon = writePng(192, (x, y) => {
  const a = roundedSquareAlpha(x, y, 192, 16, 36);
  if (!a) return [0, 0, 0, 0];
  // A lighter inner bubble so the icon reads as "messages" at small sizes.
  const inner = roundedSquareAlpha(x, y, 192, 52, 20);
  return inner ? [255, 255, 255, 255] : [...ACCENT, 255];
});

const outlineIcon = writePng(32, (x, y) => {
  const outer = roundedSquareAlpha(x, y, 32, 3, 7);
  const hole = roundedSquareAlpha(x, y, 32, 7, 5);
  return outer && !hole ? [255, 255, 255, 255] : [0, 0, 0, 0];
});

// ------------------------------------------------------------------ ZIP
// Stored (uncompressed) entries only — a manifest and two small icons.
function buildZip(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;

  for (const { name, data } of entries) {
    const nameBuf = Buffer.from(name, "utf-8");
    const crc = crc32(data);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); // local file header signature
    local.writeUInt16LE(20, 4); // version needed
    local.writeUInt16LE(0, 6); // flags
    local.writeUInt16LE(0, 8); // method: stored
    local.writeUInt16LE(0, 10); // mod time
    local.writeUInt16LE(0x2100, 12); // mod date (20xx-01-01; exact value is irrelevant here)
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28); // extra field length
    locals.push(local, nameBuf, data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0); // central directory signature
    central.writeUInt16LE(20, 4); // version made by
    central.writeUInt16LE(20, 6); // version needed
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x2100, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    centrals.push(central, nameBuf);

    offset += local.length + nameBuf.length + data.length;
  }

  const centralBuf = Buffer.concat(centrals);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); // end of central directory
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([...locals, centralBuf, end]);
}

// ----------------------------------------------------------------- main
const template = fs.readFileSync(path.join(OUT_DIR, "manifest.template.json"), "utf-8");
const manifest = template.replaceAll("${BOT_ID}", BOT_ID);
JSON.parse(manifest); // fail loudly here rather than inside Teams' uploader

fs.writeFileSync(path.join(OUT_DIR, "color.png"), colorIcon);
fs.writeFileSync(path.join(OUT_DIR, "outline.png"), outlineIcon);

const zipPath = path.join(OUT_DIR, "message-watcher.zip");
fs.writeFileSync(
  zipPath,
  buildZip([
    { name: "manifest.json", data: Buffer.from(manifest, "utf-8") },
    { name: "color.png", data: colorIcon },
    { name: "outline.png", data: outlineIcon },
  ])
);

console.log(`Built ${zipPath}`);
console.log(`  bot id:  ${BOT_ID}`);
console.log(`  icons:   color.png (192x192), outline.png (32x32)`);
console.log(`\nUpload it in Teams → Apps → Manage your apps → Upload a custom app.`);
console.log(`Then add the bot to a group or channel to start recording it.`);
