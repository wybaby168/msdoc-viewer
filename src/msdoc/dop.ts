import { BinaryReader } from '../core/binary.js';
import type { FibRgFcLcb, MsDocDocumentProperties, ParsedFib } from '../types.js';

function readBytes(tableBytes: Uint8Array, fc: number | undefined, lcb: number | undefined): Uint8Array {
  if (fc == null || lcb == null || lcb <= 0) return new Uint8Array(0);
  if (fc < 0 || fc >= tableBytes.length) return new Uint8Array(0);
  return tableBytes.subarray(fc, Math.min(tableBytes.length, fc + lcb));
}

function dopVariant(fib: ParsedFib, lcb: number): MsDocDocumentProperties['variant'] {
  if (!lcb) return 'unknown';
  if (!fib.cswNew) return 'Dop97';
  if (fib.nFibNew === 0x00d9) return 'Dop2000';
  if (fib.nFibNew === 0x0101) return 'Dop2002';
  if (fib.nFibNew === 0x010c) return 'Dop2003';
  if (fib.nFibNew === 0x0112) {
    if (lcb === 674) return 'Dop2007';
    if (lcb === 690) return 'Dop2010';
    if (lcb === 694) return 'Dop2013';
  }
  return 'unknown';
}

/**
 * DOP is a versioned document properties structure. Its exact tail differs by
 * Fib.nFibNew, so this routine decodes stable DopBase fields and exposes compact
 * compatibility diagnostics instead of pretending that every historical bit has
 * layout semantics in a browser.
 */
export function parseDop(tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, fib: ParsedFib): MsDocDocumentProperties | undefined {
  const fc = fibRgFcLcb.fcDop as number | undefined;
  const lcb = fibRgFcLcb.lcbDop as number | undefined;
  const bytes = readBytes(tableBytes, fc, lcb);
  if (!bytes.length) return undefined;
  const reader = new BinaryReader(bytes);

  // DopBase starts with dense bit fields. The first words are useful for
  // diagnostics and compatibility switches, while dxaTab is the first stable
  // layout value that this renderer can directly apply.
  const flags0 = reader.u16(0);
  const flags1 = reader.u16(2);
  const fpc = reader.u16(4);
  const nFtn = reader.u16(6);
  const nEdn = reader.u16(46);
  const dxaTab = reader.i16(10);
  const revisionCount = reader.ensure(88, 4) ? reader.i32(88) : undefined;

  // The count offsets are stable enough across Dop97+ for diagnostics, but they
  // are intentionally marked as estimates because the spec version changes the
  // tail and some producers leave them stale.
  const cWords = reader.ensure(100, 4) ? reader.i32(100) : undefined;
  const cCh = reader.ensure(104, 4) ? reader.i32(104) : undefined;
  const cPg = reader.ensure(108, 2) ? reader.u16(108) : undefined;
  const cParas = reader.ensure(110, 4) ? reader.i32(110) : undefined;

  const safeCount = (value: number | undefined, max: number): number | undefined => (value && value > 0 && value <= max ? value : undefined);

  return {
    rawLength: bytes.length,
    variant: dopVariant(fib, bytes.length),
    defaultTabStopTwips: dxaTab > 0 ? dxaTab : undefined,
    revisionCount: safeCount(revisionCount, 0x7fffffff),
    pageCountEstimate: safeCount(cPg, 100000),
    wordCountEstimate: safeCount(cWords, 1000000),
    charCountEstimate: safeCount(cCh, 1000000),
    paragraphCountEstimate: safeCount(cParas, 1000000),
    compatibility: {
      facingPages: Boolean(flags0 & 0x0001),
      widowControl: Boolean(flags0 & 0x0010),
      footnotePosition: fpc,
      footnoteInitialNumber: nFtn,
      endnoteInitialNumber: nEdn,
      rawFlags0: flags0,
      rawFlags1: flags1,
    },
    diagnostics: {
      fc,
      lcb,
      nFibNew: fib.nFibNew,
      cswNew: fib.cswNew,
    },
  };
}
