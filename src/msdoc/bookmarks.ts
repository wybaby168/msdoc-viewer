import { BinaryReader } from '../core/binary.js';
import { slugify } from '../core/utils.js';
import type { BookmarkInfo, FibRgFcLcb } from '../types.js';

function readBytes(tableBytes: Uint8Array, fc: number | undefined, lcb: number | undefined): Uint8Array {
  if (fc == null || lcb == null || lcb <= 0) return new Uint8Array(0);
  if (fc < 0 || fc >= tableBytes.length) return new Uint8Array(0);
  return tableBytes.subarray(fc, Math.min(tableBytes.length, fc + lcb));
}

function decodeUtf16LE(bytes: Uint8Array): string {
  if (!bytes.length) return '';
  return new TextDecoder('utf-16le').decode(bytes).replace(/\0+$/g, '');
}

function readSttbfBkmk(tableBytes: Uint8Array, fib: FibRgFcLcb): string[] {
  const bytes = readBytes(tableBytes, fib.fcSttbfBkmk as number | undefined, fib.lcbSttbfBkmk as number | undefined);
  if (bytes.length < 6) return [];
  const reader = new BinaryReader(bytes);
  // SttbfBkmk is an extended STTB: fExtend MUST be 0xffff, cData and cbExtra are 2-byte values.
  const fExtend = reader.u16(0);
  if (fExtend !== 0xffff) return [];
  const cData = reader.u16(2);
  const cbExtra = reader.u16(4);
  let offset = 6;
  const names: string[] = [];
  for (let i = 0; i < cData && offset + 2 <= bytes.length; i += 1) {
    const cch = reader.u16(offset);
    offset += 2;
    const byteLength = cch * 2;
    if (offset + byteLength > bytes.length) break;
    names.push(decodeUtf16LE(bytes.subarray(offset, offset + byteLength)));
    offset += byteLength + cbExtra;
  }
  return names;
}

interface FixedPlcEntry {
  index: number;
  cp: number;
  data: Uint8Array;
}

function readFixedPlcEntries(tableBytes: Uint8Array, fc: number | undefined, lcb: number | undefined, dataSize: number): FixedPlcEntry[] {
  const bytes = readBytes(tableBytes, fc, lcb);
  if (!bytes.length || dataSize < 0 || bytes.length < 4) return [];
  const count = Math.floor((bytes.length - 4) / (4 + dataSize));
  if (count <= 0 || (bytes.length - 4) % (4 + dataSize) !== 0) return [];
  const cpsByteLength = (count + 1) * 4;
  const reader = new BinaryReader(bytes);
  const entries: FixedPlcEntry[] = [];
  for (let i = 0; i < count; i += 1) {
    entries.push({
      index: i,
      cp: reader.u32(i * 4),
      data: bytes.subarray(cpsByteLength + i * dataSize, cpsByteLength + i * dataSize + dataSize),
    });
  }
  return entries;
}

function safeBookmarkId(name: string, index: number): string {
  const base = slugify(name || `bookmark-${index + 1}`) || `bookmark-${index + 1}`;
  return `msdoc-bkmk-${base}-${index + 1}`;
}

/**
 * Reads ordinary document bookmarks from SttbfBkmk + PlcfBkf/PlcfBkl.
 * FBKF.ibkl points into the paired PlcfBkl; FBKLD.ibkf points back into PlcfBkf.
 * The parser keeps both directions tolerant because old Word producers sometimes
 * persist inconsistent bookmark PLC ordering after repair operations.
 */
export function parseBookmarks(tableBytes: Uint8Array, fib: FibRgFcLcb): BookmarkInfo[] {
  const names = readSttbfBkmk(tableBytes, fib);
  if (!names.length) return [];

  const starts = readFixedPlcEntries(
    tableBytes,
    fib.fcPlcfBkf as number | undefined,
    fib.lcbPlcfBkf as number | undefined,
    4,
  );
  const ends = readFixedPlcEntries(
    tableBytes,
    fib.fcPlcfBkl as number | undefined,
    fib.lcbPlcfBkl as number | undefined,
    4,
  );

  const bookmarks: BookmarkInfo[] = [];
  for (let i = 0; i < starts.length && i < names.length; i += 1) {
    const start = starts[i]!;
    const reader = new BinaryReader(start.data);
    const ibkl = reader.u16(0);
    const kind = reader.u16(2);
    const endEntry = ends[ibkl] || ends.find((entry) => new BinaryReader(entry.data).u16(0) === i) || ends[i];
    const cpStart = start.cp;
    const cpEnd = Math.max(cpStart, endEntry?.cp ?? cpStart);
    const name = names[i] || `bookmark-${i + 1}`;
    bookmarks.push({
      id: safeBookmarkId(name, i),
      name,
      cpStart,
      cpEnd,
      hidden: name.startsWith('_'),
      kind,
      depth: endEntry ? new BinaryReader(endEntry.data).u16(2) : undefined,
    });
  }

  return bookmarks.sort((a, b) => a.cpStart - b.cpStart || a.cpEnd - b.cpEnd || a.name.localeCompare(b.name));
}
