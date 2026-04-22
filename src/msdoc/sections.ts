import { BinaryReader } from '../core/binary.js';
import { readFixedPlc } from './stories.js';
import { getSprmOperandLength } from './sprm.js';
import type { DecodedProperty, FibRgFcLcb, SectionDescriptor, SectionPageSettings } from '../types.js';

const SectionSprmCodes = {
  sprmSBkc: 0x3009,
  sprmSFTitlePage: 0x300A,
  sprmSCcolumns: 0x500B,
  sprmSDxaColumns: 0x900C,
  sprmSFPgnRestart: 0x3011,
  sprmSDyaHdrTop: 0xB017,
  sprmSDyaHdrBottom: 0xB018,
  sprmSBOrientation: 0x301D,
  sprmSXaPage: 0xB01F,
  sprmSYaPage: 0xB020,
  sprmSDxaLeft: 0xB021,
  sprmSDxaRight: 0xB022,
  sprmSDyaTop: 0x9023,
  sprmSDyaBottom: 0x9024,
  sprmSDzaGutter: 0xB025,
  sprmSPgnStart97: 0x501C,
  sprmSPgnStart: 0x7044,
} as const;

const DEFAULT_PAGE_SETTINGS: SectionPageSettings = {
  pageWidthTwips: 12240,
  pageHeightTwips: 15840,
  marginLeftTwips: 1440,
  marginRightTwips: 1440,
  marginTopTwips: 1440,
  marginBottomTwips: 1440,
  headerTopTwips: 720,
  footerBottomTwips: 720,
  gutterTwips: 0,
  columns: 1,
  columnSpacingTwips: 720,
  evenlySpacedColumns: true,
  titlePage: false,
  orientation: 'portrait',
  breakCode: 2,
  restartPageNumber: false,
  pageNumberStart: undefined,
};

function u16(bytes: Uint8Array, offset = 0): number {
  if (offset + 2 > bytes.length) return 0;
  return (bytes[offset] ?? 0) | ((bytes[offset + 1] ?? 0) << 8);
}

function i16(bytes: Uint8Array, offset = 0): number {
  const value = u16(bytes, offset);
  return value > 0x7fff ? value - 0x10000 : value;
}

function u32(bytes: Uint8Array, offset = 0): number {
  if (offset + 4 > bytes.length) return 0;
  return (((bytes[offset] ?? 0))
    | ((bytes[offset + 1] ?? 0) << 8)
    | ((bytes[offset + 2] ?? 0) << 16)
    | (((bytes[offset + 3] ?? 0) << 24) >>> 0)) >>> 0;
}

function setMeta<TValue>(name: string, value: TValue, raw: number, bytes: Uint8Array): DecodedProperty<TValue> {
  return { kind: 'section', name, value, raw, operandBytes: bytes };
}

function decodeSectionSprm(sprm: number, operandBytes: Uint8Array): DecodedProperty {
  switch (sprm) {
    case SectionSprmCodes.sprmSBkc: return setMeta('breakCode', operandBytes[0] ?? 0, sprm, operandBytes);
    case SectionSprmCodes.sprmSFTitlePage: return setMeta('titlePage', Boolean(operandBytes[0] ?? 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSCcolumns: return setMeta('columns', Math.max(1, u16(operandBytes, 0) + 1), sprm, operandBytes);
    case SectionSprmCodes.sprmSDxaColumns: return setMeta('columnSpacingTwips', i16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSFPgnRestart: return setMeta('restartPageNumber', Boolean(operandBytes[0] ?? 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDyaHdrTop: return setMeta('headerTopTwips', u16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDyaHdrBottom: return setMeta('footerBottomTwips', u16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSBOrientation: return setMeta('orientation', (operandBytes[0] ?? 0) ? 'landscape' : 'portrait', sprm, operandBytes);
    case SectionSprmCodes.sprmSXaPage: return setMeta('pageWidthTwips', u16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSYaPage: return setMeta('pageHeightTwips', u16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDxaLeft: return setMeta('marginLeftTwips', i16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDxaRight: return setMeta('marginRightTwips', i16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDyaTop: return setMeta('marginTopTwips', i16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDyaBottom: return setMeta('marginBottomTwips', i16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSDzaGutter: return setMeta('gutterTwips', u16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSPgnStart97: return setMeta('pageNumberStart', u16(operandBytes, 0), sprm, operandBytes);
    case SectionSprmCodes.sprmSPgnStart: return setMeta('pageNumberStart', u32(operandBytes, 0), sprm, operandBytes);
    default:
      return { kind: 'section', name: `sprm_${sprm.toString(16)}`, value: operandBytes, raw: sprm, operandBytes };
  }
}

function decodeSectionGrpprl(buffer: Uint8Array, startOffset: number, endOffset: number): DecodedProperty[] {
  const properties: DecodedProperty[] = [];
  let offset = startOffset;
  while (offset + 2 <= endOffset) {
    const sprm = u16(buffer, offset);
    offset += 2;
    const operandLength = getSprmOperandLength(buffer, offset, sprm);
    if (!operandLength || offset + operandLength > endOffset) break;
    const bytes = buffer.subarray(offset, offset + operandLength);
    offset += operandLength;
    properties.push(decodeSectionSprm(sprm, bytes));
  }
  return properties;
}

function parseSepx(wordBytes: Uint8Array, fcSepx: number): DecodedProperty[] {
  if (!fcSepx || fcSepx === 0xffffffff || fcSepx < 0 || fcSepx + 2 > wordBytes.length) return [];
  const reader = new BinaryReader(wordBytes);
  const cb = reader.u16(fcSepx);
  if (!cb) return [];
  const start = fcSepx + 2;
  const end = Math.min(wordBytes.length, start + cb);
  return decodeSectionGrpprl(wordBytes, start, end);
}

export function sectionPropsToPageSettings(properties: DecodedProperty[]): SectionPageSettings {
  const page: SectionPageSettings = { ...DEFAULT_PAGE_SETTINGS };
  for (const property of properties) {
    switch (property.name) {
      case 'pageWidthTwips': page.pageWidthTwips = Math.max(144, Number(property.value) || page.pageWidthTwips); break;
      case 'pageHeightTwips': page.pageHeightTwips = Math.max(144, Number(property.value) || page.pageHeightTwips); break;
      case 'marginLeftTwips': page.marginLeftTwips = Math.abs(Number(property.value) || 0); break;
      case 'marginRightTwips': page.marginRightTwips = Math.abs(Number(property.value) || 0); break;
      case 'marginTopTwips': page.marginTopTwips = Math.abs(Number(property.value) || 0); break;
      case 'marginBottomTwips': page.marginBottomTwips = Math.abs(Number(property.value) || 0); break;
      case 'headerTopTwips': page.headerTopTwips = Math.max(0, Number(property.value) || 0); break;
      case 'footerBottomTwips': page.footerBottomTwips = Math.max(0, Number(property.value) || 0); break;
      case 'gutterTwips': page.gutterTwips = Math.max(0, Number(property.value) || 0); break;
      case 'columns': page.columns = Math.max(1, Number(property.value) || 1); break;
      case 'columnSpacingTwips': page.columnSpacingTwips = Math.max(0, Number(property.value) || 0); break;
      case 'titlePage': page.titlePage = Boolean(property.value); break;
      case 'orientation': page.orientation = property.value === 'landscape' ? 'landscape' : 'portrait'; break;
      case 'breakCode': page.breakCode = Number(property.value) || 0; break;
      case 'restartPageNumber': page.restartPageNumber = Boolean(property.value); break;
      case 'pageNumberStart': page.pageNumberStart = Number(property.value) || undefined; break;
      default:
        break;
    }
  }
  if (page.orientation === 'landscape' && page.pageHeightTwips > page.pageWidthTwips) {
    [page.pageWidthTwips, page.pageHeightTwips] = [page.pageHeightTwips, page.pageWidthTwips];
  }
  if (page.orientation === 'portrait' && page.pageWidthTwips > page.pageHeightTwips) {
    // Some producers rely on the dimensions alone; keep the physical sizes intact.
    page.orientation = 'landscape';
  }
  return page;
}

export function readSections(wordBytes: Uint8Array, tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, mainStoryLength: number): SectionDescriptor[] {
  const entries = readFixedPlc(tableBytes, fibRgFcLcb.fcPlcfSed as number | undefined, fibRgFcLcb.lcbPlcfSed as number | undefined, 12);
  if (!entries.length) {
    return [{
      index: 0,
      cpStart: 0,
      cpEnd: mainStoryLength,
      fcSepx: 0,
      properties: [],
      page: { ...DEFAULT_PAGE_SETTINGS },
    }];
  }

  return entries.map((entry) => {
    const reader = new BinaryReader(entry.data);
    const fcSepx = reader.u32(2);
    const properties = parseSepx(wordBytes, fcSepx);
    return {
      index: entry.index,
      cpStart: Math.max(0, entry.cpStart),
      cpEnd: Math.max(Math.max(0, entry.cpStart), entry.cpEnd),
      fcSepx,
      properties,
      page: sectionPropsToPageSettings(properties),
    } satisfies SectionDescriptor;
  });
}

export function findSectionIndex(sections: SectionDescriptor[], cp: number): number {
  for (const section of sections) {
    if (cp >= section.cpStart && cp < section.cpEnd) return section.index;
  }
  return sections.length ? sections[Math.max(0, sections.length - 1)]!.index : 0;
}
