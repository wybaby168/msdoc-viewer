import { parseCFB } from '../core/cfb.js';
import { cleanTextControlChars, pushWarning, shallowEqual, uniqueId } from '../core/utils.js';
import { parseClx, buildPieceTextCache, getTextByCp, splitParagraphRanges } from './clx.js';
import { DOC_CONTROL } from './constants.js';
import { parseFib } from './fib.js';
import { readChpxRuns, readPapxRuns, type ChpxRun, type PapxRun } from './fkp.js';
import { parseFonts } from './fonts.js';
import { extractObjectPool, extractPictureAsset } from './objects.js';
import {
  applyTableStateToCells,
  charPropsToState,
  getTableDepth,
  paraPropsToState,
  tablePropsToState,
} from './properties.js';
import { mergePropertyArrays, parseStyles, splitPropertiesByKind } from './styles.js';
import type {
  AttachmentAsset,
  CharSegment,
  CharState,
  DecodedProperty,
  FieldInstruction,
  FontInfo,
  FontsCollection,
  ImageAsset,
  InlineNode,
  MsDocAsset,
  MsDocParseOptions,
  MsDocParseResult,
  ObjectPoolInfo,
  ParagraphBlock,
  ParagraphModel,
  ParagraphRange,
  ResolvedStyle,
  StyleCollection,
  TableBlock,
  TableCellBlock,
  TableRowBlock,
} from '../types.js';

function getOverlappingRuns<T extends { cpStart: number; cpEnd: number }>(
  runs: T[],
  cpStart: number,
  cpEnd: number,
  cursorRef?: { index: number },
): T[] {
  let cursor = cursorRef?.index || 0;
  while (cursor < runs.length && runs[cursor]!.cpEnd <= cpStart) cursor += 1;
  if (cursorRef) cursorRef.index = cursor;
  const list: T[] = [];
  let i = cursor;
  while (i < runs.length && runs[i]!.cpStart < cpEnd) {
    if (runs[i]!.cpEnd > cpStart) list.push(runs[i] as T);
    i += 1;
  }
  return list;
}

function normalizeTextStyleName(name: unknown): string {
  return String(name || '').trim();
}

function decodeFieldInstruction(instruction: unknown): FieldInstruction | null {
  const normalized = String(instruction || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!normalized) return null;
  const upper = normalized.toUpperCase();
  if (upper.startsWith('HYPERLINK')) {
    const body = normalized.slice(9).trim();
    const quoted = body.match(/"([^"]+)"/);
    const href = quoted?.[1] || body.split(/\s+/)[0] || '';
    return href ? { type: 'hyperlink', href } : null;
  }
  if (upper.startsWith('INCLUDEPICTURE')) {
    const body = normalized.slice('INCLUDEPICTURE'.length).trim();
    const quoted = body.match(/"([^"]+)"/);
    return { type: 'includePicture', target: quoted?.[1] || body.split(/\s+/)[0] || '' };
  }
  if (upper.startsWith('EMBED')) {
    return { type: 'embed', raw: normalized };
  }
  if (upper.startsWith('LINK')) {
    return { type: 'link', raw: normalized };
  }
  return { type: 'unknown', raw: normalized };
}


function isRenderableExternalImageUrl(url: string): boolean {
  return /^(?:https?:|blob:)/i.test(url) || /^data:image\//i.test(url);
}

function createFieldImageAsset(target: string): ImageAsset | null {
  const normalized = String(target || '').trim();
  if (!normalized) return null;

  let mime = 'application/octet-stream';
  if (/\.png(?:$|[?#])/i.test(normalized)) mime = 'image/png';
  else if (/\.jpe?g(?:$|[?#])/i.test(normalized)) mime = 'image/jpeg';
  else if (/\.gif(?:$|[?#])/i.test(normalized)) mime = 'image/gif';
  else if (/\.bmp(?:$|[?#])/i.test(normalized)) mime = 'image/bmp';
  else if (/\.svg(?:$|[?#])/i.test(normalized)) mime = 'image/svg+xml';
  else if (/\.tiff?(?:$|[?#])/i.test(normalized)) mime = 'image/tiff';

  const dataUrl = /^data:image\//i.test(normalized) ? normalized : '';
  const sourceUrl = dataUrl ? undefined : normalized;
  const localExternal = /^(?:file:|\\\\|[a-zA-Z]:[\\/]|\/)/i.test(normalized);

  return {
    id: uniqueId('asset-img-field'),
    type: 'image',
    mime,
    bytes: new Uint8Array(0),
    dataUrl,
    sourceUrl,
    displayable: isRenderableExternalImageUrl(normalized) && mime.startsWith('image/'),
    meta: {
      sourceKind: 'linked',
      linkedPath: normalized,
      localExternal,
      browserRenderable: isRenderableExternalImageUrl(normalized),
      pictureOffset: -1,
      lcb: 0,
      cbHeader: 0,
    },
  };
}

function mergeTextNode(target: InlineNode[], node: Extract<InlineNode, { type: 'text' }>): void {
  if (!node.text) return;
  const last = target[target.length - 1];
  if (last && last.type === 'text' && last.href === node.href && shallowEqual(last.style, node.style)) {
    last.text += node.text;
    return;
  }
  target.push(node);
}

interface FieldFrame {
  instruction: string;
  parsed: FieldInstruction | null;
  readingInstruction: boolean;
  nodes: InlineNode[];
  resultStyle?: CharState;
}

function emitInline(targetStack: FieldFrame[], output: InlineNode[], node: InlineNode | null | undefined): void {
  if (!node) return;
  const target = targetStack.length ? targetStack[targetStack.length - 1]!.nodes : output;
  if (node.type === 'text') mergeTextNode(target, node);
  else target.push(node);
}

function getObjectPoolInfo(objectPool: Map<string, ObjectPoolInfo>, pictureOffset: number): ObjectPoolInfo | null {
  const candidates = [
    `_${pictureOffset}`,
    `_${String(pictureOffset)}`,
    `_${pictureOffset.toString(16)}`,
    `_${pictureOffset.toString(16).toUpperCase()}`,
  ];
  for (const key of candidates) {
    if (objectPool.has(key)) return objectPool.get(key) || null;
  }
  return null;
}

function createAssetResolver(
  dataBytes: Uint8Array,
  objectPool: Map<string, ObjectPoolInfo>,
  assets: MsDocAsset[],
  usedAttachmentNames: Set<string>,
  assetCache: Map<number, MsDocAsset | null>,
  options: MsDocParseOptions = {},
): (charState: CharState) => MsDocAsset | null {
  return function resolveAsset(charState: CharState): MsDocAsset | null {
    const pictureOffset = charState?.pictureOffset;
    if (pictureOffset == null) return null;
    if (assetCache.has(pictureOffset)) return assetCache.get(pictureOffset) || null;

    let asset: MsDocAsset | null = null;
    const objectInfo = getObjectPoolInfo(objectPool, pictureOffset);

    if ((charState.ole2 || charState.object || charState.data) && objectInfo?.attachment) {
      asset = objectInfo.attachment;
      usedAttachmentNames.add(objectInfo.entry.name);
    }

    if (!asset && dataBytes?.length) {
      const extracted = extractPictureAsset(dataBytes, pictureOffset, options);
      if (extracted && extracted.mime !== 'application/octet-stream') {
        asset = extracted;
      } else if (!asset && objectInfo?.attachment) {
        asset = objectInfo.attachment;
        usedAttachmentNames.add(objectInfo.entry.name);
      } else if (extracted) {
        asset = extracted;
      }
    }

    if (!asset && objectInfo?.attachment) {
      asset = objectInfo.attachment;
      usedAttachmentNames.add(objectInfo.entry.name);
    }

    if (asset) assets.push(asset);
    assetCache.set(pictureOffset, asset);
    return asset;
  };
}

function normalizePlainTextChar(ch: string): string {
  switch (ch) {
    case DOC_CONTROL.nonBreakingHyphen:
      return '-';
    case DOC_CONTROL.nonRequiredHyphen:
      return '';
    case DOC_CONTROL.annotationRef:
      return '';
    default:
      return ch;
  }
}

function buildInlineNodes(segments: CharSegment[], resolveAsset: (charState: CharState) => MsDocAsset | null): InlineNode[] {
  const output: InlineNode[] = [];
  const fieldStack: FieldFrame[] = [];

  for (const segment of segments) {
    if (!segment.text || segment.state.hidden) continue;
    for (const rawCh of segment.text) {
      const ch = normalizePlainTextChar(rawCh);

      if (ch === DOC_CONTROL.fieldStart) {
        fieldStack.push({ instruction: '', parsed: null, readingInstruction: true, nodes: [], resultStyle: segment.state });
        continue;
      }
      if (ch === DOC_CONTROL.fieldSeparator) {
        const current = fieldStack[fieldStack.length - 1];
        if (current) {
          current.parsed = decodeFieldInstruction(current.instruction);
          current.readingInstruction = false;
        }
        continue;
      }
      if (ch === DOC_CONTROL.fieldEnd) {
        const current = fieldStack.pop();
        if (!current) continue;
        let nodes = current.nodes;
        if (current.readingInstruction) {
          current.parsed = decodeFieldInstruction(current.instruction);
        }
        if (current.parsed?.type === 'includePicture' && !nodes.some((node) => node.type === 'image' || node.type === 'attachment')) {
          const asset = createFieldImageAsset(current.parsed.target);
          if (asset) nodes = [{ type: 'image', asset, style: current.resultStyle || segment.state }];
        }
        if (current.parsed?.type === 'hyperlink') {
          const href = current.parsed.href;
          nodes = nodes.map((node) => {
            if (node.type === 'lineBreak' || node.type === 'pageBreak') return node;
            return { ...node, href } as InlineNode;
          });
        }
        for (const node of nodes) emitInline(fieldStack, output, node);
        continue;
      }

      const currentField = fieldStack[fieldStack.length - 1];
      if (currentField?.readingInstruction) {
        currentField.instruction += ch;
        continue;
      }
      if (currentField && !currentField.resultStyle) currentField.resultStyle = segment.state;

      if (ch === DOC_CONTROL.hardLineBreak) {
        emitInline(fieldStack, output, { type: 'lineBreak' });
        continue;
      }
      if (ch === DOC_CONTROL.pageBreak) {
        emitInline(fieldStack, output, { type: 'pageBreak' });
        continue;
      }
      if (ch === DOC_CONTROL.picture) {
        const asset = resolveAsset(segment.state);
        if (asset?.type === 'image') {
          emitInline(fieldStack, output, { type: 'image', asset, style: segment.state });
        } else if (asset?.type === 'attachment') {
          emitInline(fieldStack, output, { type: 'attachment', asset, style: segment.state });
        }
        continue;
      }

      emitInline(fieldStack, output, {
        type: 'text',
        text: cleanTextControlChars(ch),
        style: segment.state,
      });
    }
  }

  while (fieldStack.length) {
    const current = fieldStack.pop()!;
    for (const node of current.nodes) emitInline(fieldStack, output, node);
  }

  return output;
}

function buildCharSegments(
  range: ParagraphRange,
  paragraphText: string,
  chpxRuns: ChpxRun[],
  styles: StyleCollection,
  baseCharProps: DecodedProperty[],
  resolveFont: (fontId: number | undefined) => FontInfo | null,
  cursorRef: { index: number },
): CharSegment[] {
  const overlaps = getOverlappingRuns(chpxRuns, range.cpStart, range.cpEnd, cursorRef);
  const boundaries = new Set<number>([range.cpStart, range.cpEnd]);
  for (const run of overlaps) {
    boundaries.add(Math.max(range.cpStart, run.cpStart));
    boundaries.add(Math.min(range.cpEnd, run.cpEnd));
  }
  const points = Array.from(boundaries).sort((a, b) => a - b);
  const segments: CharSegment[] = [];

  for (let i = 0; i < points.length - 1; i += 1) {
    const cpStart = points[i]!;
    const cpEnd = points[i + 1]!;
    if (cpEnd <= cpStart) continue;
    const coveringRun = overlaps.find((run) => run.cpStart <= cpStart && run.cpEnd >= cpEnd);
    const directProps = coveringRun?.properties || [];
    const directState = charPropsToState(directProps);
    const charStyleProps = directState.charStyleId != null ? styles.resolveStyle(directState.charStyleId).charProps : [];
    const finalProps = mergePropertyArrays(baseCharProps, charStyleProps, directProps);
    const state = charPropsToState(finalProps);
    const font = resolveFont(state.fontFamilyId);
    if (font) state.fontFamily = font.name || font.altName || undefined;
    const localStart = cpStart - range.cpStart;
    const localEnd = cpEnd - range.cpStart;
    const text = paragraphText.slice(localStart, localEnd);
    if (!text) continue;
    const last = segments[segments.length - 1];
    if (last && shallowEqual(last.state, state)) {
      last.text += text;
      last.cpEnd = cpEnd;
      continue;
    }
    segments.push({ cpStart, cpEnd, text, state });
  }

  if (!segments.length && paragraphText) {
    const state = charPropsToState(baseCharProps);
    const font = resolveFont(state.fontFamilyId);
    if (font) state.fontFamily = font.name || font.altName || undefined;
    segments.push({ cpStart: range.cpStart, cpEnd: range.cpEnd, text: paragraphText, state });
  }

  return segments;
}

function buildParagraphModel(
  range: ParagraphRange,
  paragraphText: string,
  styles: StyleCollection,
  fonts: FontsCollection,
  chpxRuns: ChpxRun[],
  resolveAsset: (charState: CharState) => MsDocAsset | null,
  chpxCursor: { index: number },
): ParagraphModel {
  const directSplit = splitPropertiesByKind(range.properties || []);
  const paraStyleId = range.styleId || (directSplit.para.find((item) => item.name === 'styleId')?.value as number | undefined) || 0;
  const paraStyle: ResolvedStyle = styles.resolveStyle(paraStyleId);
  const paraProps = mergePropertyArrays(paraStyle.paraProps, directSplit.para);
  const paraState = paraPropsToState(paraProps);

  const directTableState = tablePropsToState(directSplit.table);
  const tableStyleProps = directTableState.styleId != null ? styles.resolveStyle(directTableState.styleId).tableProps : [];
  const tableProps = mergePropertyArrays(tableStyleProps, directSplit.table);
  const tableState = tablePropsToState(tableProps);

  const baseCharProps = paraStyle.charProps;
  const resolveFont = (fontId: number | undefined) => fonts.byIndex(fontId);
  const segments = buildCharSegments(range, paragraphText, chpxRuns, styles, baseCharProps, resolveFont, chpxCursor);
  const inlines = buildInlineNodes(segments, resolveAsset);

  return {
    id: uniqueId('para'),
    cpStart: range.cpStart,
    cpEnd: range.cpEnd,
    terminator: range.terminator || '',
    text: paragraphText,
    rawProperties: range.properties || [],
    styleId: paraStyleId,
    styleName: normalizeTextStyleName(styles.styles.get(paraStyleId)?.name),
    paraProps,
    paraState,
    tableProps,
    tableState,
    segments,
    inlines,
  };
}

function finalizeTableGrid(rows: TableRowBlock[]): void {
  for (const row of rows) {
    for (let i = 0; i < row.cells.length; i += 1) {
      const cell = row.cells[i]!;
      const merge = cell.meta?.merge || 0;
      cell.colIndex = i;
      cell.colspan = 1;
      cell.rowspan = 1;
      cell.hidden = false;
      if (merge === 1) {
        cell.hidden = true;
        continue;
      }
      if (merge > 1) {
        let j = i + 1;
        while (j < row.cells.length && (row.cells[j]!.meta?.merge || 0) === 1) {
          row.cells[j]!.hidden = true;
          cell.colspan += 1;
          j += 1;
        }
      }
    }
  }

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex]!;
    for (const cell of row.cells) {
      if (cell.hidden) continue;
      const vertMerge = cell.meta?.vertMerge || 0;
      if (vertMerge === 1) {
        cell.hidden = true;
        continue;
      }
      if (vertMerge > 1) {
        let nextIndex = rowIndex + 1;
        while (nextIndex < rows.length) {
          let canMerge = true;
          for (let col = cell.colIndex || 0; col < (cell.colIndex || 0) + (cell.colspan || 1); col += 1) {
            const nextCell = rows[nextIndex]!.cells[col];
            if (!nextCell || (nextCell.meta?.vertMerge || 0) !== 1) {
              canMerge = false;
              break;
            }
          }
          if (!canMerge) break;
          for (let col = cell.colIndex || 0; col < (cell.colIndex || 0) + (cell.colspan || 1); col += 1) {
            rows[nextIndex]!.cells[col]!.hidden = true;
          }
          cell.rowspan = (cell.rowspan || 1) + 1;
          nextIndex += 1;
        }
      }
    }
  }
}

function paragraphToBlock(paragraph: ParagraphModel): ParagraphBlock {
  return {
    type: 'paragraph',
    id: paragraph.id,
    styleId: paragraph.styleId,
    styleName: paragraph.styleName,
    paraState: paragraph.paraState,
    inlines: paragraph.inlines,
    text: paragraph.text,
  };
}

function buildTableBlock(tableParagraphs: ParagraphModel[]): TableBlock {
  const rows: TableRowBlock[] = [];
  let pendingRow: { cells: TableCellBlock[] } = { cells: [] };
  let pendingCellParagraphs: ParagraphModel[] = [];

  for (const paragraph of tableParagraphs) {
    pendingCellParagraphs.push(paragraph);
    if (paragraph.terminator === DOC_CONTROL.cellMark) {
      pendingRow.cells.push({
        id: uniqueId('cell'),
        paragraphs: pendingCellParagraphs.map(paragraphToBlock),
        meta: null,
      });
      pendingCellParagraphs = [];

      if (paragraph.paraState.tableRowEnd || paragraph.paraState.innerTableRowEnd) {
        const cellDefs = applyTableStateToCells(paragraph.tableState);
        while (
          cellDefs.length &&
          pendingRow.cells.length > cellDefs.length &&
          pendingRow.cells[pendingRow.cells.length - 1]!.paragraphs.every((block) => !block.text && !(block.inlines || []).length)
        ) {
          pendingRow.cells.pop();
        }
        pendingRow.cells.forEach((cell, index) => {
          cell.meta = cellDefs[index] || { index };
        });
        const gridWidthTwips = cellDefs.length
          ? ((cellDefs[cellDefs.length - 1]!.rightBoundary || 0) - (cellDefs[0]!.leftBoundary || 0))
          : 0;
        rows.push({
          id: uniqueId('row'),
          cells: pendingRow.cells,
          state: paragraph.tableState,
          gridWidthTwips,
        });
        pendingRow = { cells: [] };
      }
    }
  }

  if (pendingCellParagraphs.length) {
    pendingRow.cells.push({ id: uniqueId('cell'), paragraphs: pendingCellParagraphs.map(paragraphToBlock), meta: null });
  }
  if (pendingRow.cells.length) {
    rows.push({ id: uniqueId('row'), cells: pendingRow.cells, state: tableParagraphs[0]?.tableState || tablePropsToState([]), gridWidthTwips: 0 });
  }

  finalizeTableGrid(rows);

  const gridWidthTwips = rows.find((row) => row.gridWidthTwips)?.gridWidthTwips || 0;
  const depth = Math.max(...tableParagraphs.map((paragraph) => getTableDepth(paragraph.paraState)), 1);

  return {
    type: 'table',
    id: uniqueId('table'),
    depth,
    rows,
    state: rows[0]?.state || tablePropsToState([]),
    gridWidthTwips,
  };
}

function buildBlocks(paragraphs: ParagraphModel[]): MsDocParseResult['blocks'] {
  const blocks: MsDocParseResult['blocks'] = [];
  let index = 0;
  while (index < paragraphs.length) {
    const paragraph = paragraphs[index]!;
    const depth = getTableDepth(paragraph.paraState);
    if (depth <= 0) {
      blocks.push(paragraphToBlock(paragraph));
      index += 1;
      continue;
    }
    const tableParagraphs: ParagraphModel[] = [];
    while (index < paragraphs.length && getTableDepth(paragraphs[index]!.paraState) > 0) {
      tableParagraphs.push(paragraphs[index]!);
      index += 1;
    }
    blocks.push(buildTableBlock(tableParagraphs));
  }
  return blocks;
}

function collectAssetWarnings(assets: MsDocAsset[], warnings: MsDocParseResult['warnings']): void {
  for (const asset of assets) {
    if (asset.type !== 'image') continue;
    if (asset.mime === 'application/octet-stream') {
      pushWarning(warnings, 'Encountered picture data that could not be decoded to a supported image payload', {
        code: 'unsupported-image-payload',
        severity: 'warning',
        offset: asset.meta?.pictureOffset,
        details: { mime: asset.mime, sourceKind: asset.meta?.sourceKind },
      });
      continue;
    }
    if (asset.displayable === false) {
      pushWarning(warnings, 'Encountered an image that was parsed but is not directly browser-displayable', {
        code: asset.meta?.localExternal ? 'linked-local-image' : 'non-displayable-image',
        severity: 'warning',
        offset: asset.meta?.pictureOffset,
        details: { mime: asset.mime, sourceUrl: asset.sourceUrl, linkedPath: asset.meta?.linkedPath },
      });
    }
  }
}

/**
 * Main MS-DOC entry point.
 * It parses the OLE container, restores text through the piece table, resolves
 * paragraph/character/table properties, and finally produces a normalized AST
 * that the HTML renderer can consume.
 */
export function parseMsDoc(input: ArrayBuffer | Uint8Array | ArrayBufferView, options: MsDocParseOptions = {}): MsDocParseResult {
  const warnings = [] as MsDocParseResult['warnings'];
  const cfb = parseCFB(input, options);
  warnings.push(...(cfb.warnings || []));

  const wordBytes = cfb.getStream('/WordDocument');
  if (!wordBytes) throw new Error('Missing WordDocument stream');

  const fib = parseFib(wordBytes);
  if (fib.base.wIdent !== 0xA5EC) {
    pushWarning(warnings, `Unexpected FIB identifier: 0x${fib.base.wIdent.toString(16)}`);
  }
  if (fib.base.fEncrypted) {
    throw new Error('Encrypted .doc files are not supported yet');
  }

  const tableBytes = cfb.getStream(fib.base.fWhichTblStm ? '/1Table' : '/0Table');
  if (!tableBytes) throw new Error('Missing table stream');

  const dataBytes = cfb.getStream('/Data') || new Uint8Array(0);
  const clx = parseClx(tableBytes, fib.fibRgFcLcb);
  const pieceTexts = buildPieceTextCache(wordBytes, clx);
  const documentText = pieceTexts.join('');
  const mainStoryEnd = fib.fibRgLw.ccpText > 0 ? fib.fibRgLw.ccpText : documentText.length;

  const styles = parseStyles(tableBytes, fib.fibRgFcLcb);
  const fonts = parseFonts(tableBytes, fib.fibRgFcLcb);
  const chpxRuns = readChpxRuns(wordBytes, tableBytes, fib, clx).filter((run) => run.cpStart < mainStoryEnd);
  const papxRuns = readPapxRuns(wordBytes, tableBytes, fib, clx)
    .filter((run) => run.cpStart < mainStoryEnd)
    .map((run) => ({ ...run, cpEnd: Math.min(run.cpEnd, mainStoryEnd) }));

  const ranges: ParagraphRange[] = papxRuns.length
    ? papxRuns.map((run) => ({
        cpStart: run.cpStart,
        cpEnd: run.cpEnd,
        terminator: documentText[run.cpEnd - 1] || '',
        styleId: run.styleId,
        properties: run.properties,
      }))
    : splitParagraphRanges(documentText.slice(0, mainStoryEnd)).map((range) => ({ ...range, styleId: 0, properties: [] }));

  const objectPool = extractObjectPool(cfb);
  const assets: MsDocAsset[] = [];
  const usedAttachmentNames = new Set<string>();
  const assetCache = new Map<number, MsDocAsset | null>();
  const resolveAsset = createAssetResolver(dataBytes, objectPool, assets, usedAttachmentNames, assetCache, options);
  const chpxCursor = { index: 0 };

  const paragraphs = ranges.map((range) => {
    const rawParagraphText = getTextByCp(wordBytes, clx, pieceTexts, range.cpStart, range.cpEnd);
    const terminator = range.terminator === DOC_CONTROL.paragraph || range.terminator === DOC_CONTROL.cellMark ? range.terminator : '';
    const paragraphText = terminator && rawParagraphText.endsWith(terminator)
      ? rawParagraphText.slice(0, -1)
      : rawParagraphText;
    return buildParagraphModel({ ...range, terminator }, paragraphText, styles, fonts, chpxRuns, resolveAsset, chpxCursor);
  });

  const blocks = buildBlocks(paragraphs);
  const trailingAttachments = Array.from(objectPool.values())
    .filter((item) => item?.attachment && !usedAttachmentNames.has(item.entry.name))
    .map((item) => item.attachment as AttachmentAsset);

  for (const attachment of trailingAttachments) assets.push(attachment);
  if (trailingAttachments.length) {
    blocks.push({ type: 'attachments', id: uniqueId('attachments'), items: trailingAttachments });
  }

  collectAssetWarnings(assets, warnings);

  return {
    kind: 'msdoc',
    version: 1,
    warnings,
    meta: {
      fib: {
        wIdent: fib.base.wIdent,
        nFib: fib.base.nFib,
        fWhichTblStm: fib.base.fWhichTblStm,
        fComplex: fib.base.fComplex,
        fEncrypted: fib.base.fEncrypted,
        ccpText: fib.fibRgLw.ccpText,
      },
      counts: {
        paragraphs: paragraphs.length,
        blocks: blocks.length,
        assets: assets.length,
        styles: styles.styles.size,
        fonts: fonts.fonts.length,
      },
    },
    fonts: fonts.fonts,
    styles: Array.from(styles.styles.values()).map((style) => ({
      istd: style.istd,
      name: style.name,
      type: style.stdfBase?.stk,
      basedOn: style.stdfBase?.istdBase,
      next: style.stdfBase?.istdNext,
    })),
    blocks,
    assets,
  };
}
