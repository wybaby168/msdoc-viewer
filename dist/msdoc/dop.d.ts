import type { FibRgFcLcb, MsDocDocumentProperties, ParsedFib } from '../types.js';
/**
 * DOP is a versioned document properties structure. Its exact tail differs by
 * Fib.nFibNew, so this routine decodes stable DopBase fields and exposes compact
 * compatibility diagnostics instead of pretending that every historical bit has
 * layout semantics in a browser.
 */
export declare function parseDop(tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, fib: ParsedFib): MsDocDocumentProperties | undefined;
