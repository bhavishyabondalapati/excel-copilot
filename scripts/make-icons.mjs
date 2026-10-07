// Generates the add-in's PNG icons (16/32/64/80 px) with no image libraries:
// a green rounded square with a white 2x2 grid. Run: npm run icons
import { writeFileSync } from "node:fs";
import { deflateSync } from "node:zlib";

const crcTable = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
const crc32 = (buf) => {
  let c = 0xffffffff;
  for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
};
const chunk = (type, data) => {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const td = Buffer.concat([Buffer.from(type), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(td));
  return Buffer.concat([len, td, crc]);
};

function png(size) {
  const raw = Buffer.alloc(size * (size * 4 + 1));
  const r = size * 0.2; // corner radius
  for (let y = 0; y < size; y++) {
    raw[y * (size * 4 + 1)] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const i = y * (size * 4 + 1) + 1 + x * 4;
      const cx = Math.min(Math.max(x + 0.5, r), size - r);
      const cy = Math.min(Math.max(y + 0.5, r), size - r);
      const inside = Math.hypot(x + 0.5 - cx, y + 0.5 - cy) <= r;
      // White frame + cross = a little 2x2 spreadsheet grid.
      const u = (x + 0.5) / size;
      const v = (y + 0.5) / size;
      const grid = (Math.abs(u - 0.5) < 0.04 || Math.abs(v - 0.5) < 0.04) && u > 0.2 && u < 0.8 && v > 0.2 && v < 0.8;
      const frame = u > 0.2 && u < 0.8 && v > 0.2 && v < 0.8 && (u < 0.26 || u > 0.74 || v < 0.26 || v > 0.74);
      const white = grid || frame;
      const [R, G, B] = white ? [255, 255, 255] : [16, 124, 65];
      raw[i] = R; raw[i + 1] = G; raw[i + 2] = B; raw[i + 3] = inside ? 255 : 0;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8; ihdr[9] = 6; // 8-bit RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk("IHDR", ihdr),
    chunk("IDAT", deflateSync(raw)),
    chunk("IEND", Buffer.alloc(0)),
  ]);
}

for (const size of [16, 32, 64, 80, 128]) writeFileSync(`assets/icon-${size}.png`, png(size));
console.log("icons written to assets/");
