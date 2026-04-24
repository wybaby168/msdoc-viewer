import type { BookmarkInfo, FibRgFcLcb } from '../types.js';
/**
 * Reads ordinary document bookmarks from SttbfBkmk + PlcfBkf/PlcfBkl.
 * FBKF.ibkl points into the paired PlcfBkl; FBKLD.ibkf points back into PlcfBkf.
 * The parser keeps both directions tolerant because old Word producers sometimes
 * persist inconsistent bookmark PLC ordering after repair operations.
 */
export declare function parseBookmarks(tableBytes: Uint8Array, fib: FibRgFcLcb): BookmarkInfo[];
