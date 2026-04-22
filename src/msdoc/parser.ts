import { parseCFB } from '../core/cfb.js';
import { cleanTextControlChars, pushWarning, shallowEqual, uniqueId } from '../core/utils.js';
import { parseClx, buildPieceTextCache, getTextByCp, splitParagraphRanges } from './clx.js';
import { DOC_CONTROL } from './constants.js';
import { parseFib } from './fib.js';
import { readChpxRuns, readPapxRuns, type ChpxRun, type PapxRun } from './fkp.js';
import { parseFonts } from './fonts.js';
import { extractObjectPool, extractPictureAsset } from './objects.js';
import { readShapeAnchors } from './shapes.js';
import {
  buildHeaderStoryDescriptors,
  buildStoryWindows,
  parseCommentRefMeta,
  parseSttbfRMark,
  parseTextboxMeta,
  parseXstArray,
  readFixedPlc,
} from './stories.js';
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
  CommentItem,
  CommentsBlock,
  DecodedProperty,
  FieldInstruction,
  FontInfo,
  FontsCollection,
  HeaderFooterStory,
  HeadersBlock,
  ImageAsset,
  InlineNode,
  MsDocAsset,
  MsDocParseOptions,
  MsDocParseResult,
  NoteItem,
  NotesBlock,
  ObjectPoolInfo,
  ParagraphBlock,
  ParagraphModel,
  ParagraphRange,
  ResolvedStyle,
  ShapeAnchorInfo,
  ShapesBlock,
  StyleCollection,
  TableBlock,
  TableCellBlock,
  TableRowBlock,
  TableState,
  TextboxItem,
  TextboxesBlock,
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


interface NoteReferenceInfo {
  kind: 'footnote' | 'endnote';
  id: string;
  label: string;
}

interface CommentReferenceInfo {
  id: string;
  label: string;
  author?: string;
}

interface InlineBuildContext {
  noteRefs?: Map<number, NoteReferenceInfo>;
  commentRefs?: Map<number, CommentReferenceInfo>;
}

interface ParsedStoryContent {
  paragraphs: ParagraphModel[];
  blocks: Array<ParagraphBlock | TableBlock>;
  text: string;
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
    case '':
      return '';
    default:
      return ch;
  }
}

function inlineNodesToPlainText(nodes: InlineNode[]): string {
  return nodes.map((node) => {
    if (node.type === 'text') return node.text;
    if (node.type === 'lineBreak' || node.type === 'pageBreak') return '\n';
    if (node.type === 'noteRef' || node.type === 'commentRef') return node.label;
    return '';
  }).join('');
}

function buildInlineNodes(
  segments: CharSegment[],
  resolveAsset: (charState: CharState) => MsDocAsset | null,
  context: InlineBuildContext = {},
): InlineNode[] {
  const output: InlineNode[] = [];
  const fieldStack: FieldFrame[] = [];

  for (const segment of segments) {
    if (!segment.text || segment.state.hidden) continue;
    let cp = segment.cpStart;
    for (const rawCh of segment.text) {
      const noteRef = context.noteRefs?.get(cp);
      if (noteRef) {
        emitInline(fieldStack, output, {
          type: 'noteRef',
          noteType: noteRef.kind,
          refId: noteRef.id,
          label: noteRef.label,
        });
        cp += 1;
        continue;
      }
      const commentRef = context.commentRefs?.get(cp);
      if (commentRef) {
        emitInline(fieldStack, output, {
          type: 'commentRef',
          commentId: commentRef.id,
          label: commentRef.label,
          author: commentRef.author,
        });
        cp += 1;
        continue;
      }

      const ch = normalizePlainTextChar(rawCh);
      cp += 1;
      if (!ch) continue;

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
        const parentField = fieldStack[fieldStack.length - 1];
        if (parentField?.readingInstruction) {
          parentField.instruction += inlineNodesToPlainText(nodes);
        } else {
          for (const node of nodes) emitInline(fieldStack, output, node);
        }
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
    const parentField = fieldStack[fieldStack.length - 1];
    if (parentField?.readingInstruction) {
      parentField.instruction += current.instruction + inlineNodesToPlainText(current.nodes);
      continue;
    }
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
  revisionAuthors: string[] = [],
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
    if (state.revisionAuthorIndex != null && revisionAuthors[state.revisionAuthorIndex]) {
      state.revisionAuthor = revisionAuthors[state.revisionAuthorIndex];
    }
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
    if (state.revisionAuthorIndex != null && revisionAuthors[state.revisionAuthorIndex]) {
      state.revisionAuthor = revisionAuthors[state.revisionAuthorIndex];
    }
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
  revisionAuthors: string[] = [],
  inlineContext: InlineBuildContext = {},
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
  const segments = buildCharSegments(range, paragraphText, chpxRuns, styles, baseCharProps, resolveFont, chpxCursor, revisionAuthors);
  const inlines = buildInlineNodes(segments, resolveAsset, inlineContext);

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

function buildRangesForCpInterval(
  cpStart: number,
  cpEnd: number,
  documentText: string,
  papxRuns: PapxRun[],
): ParagraphRange[] {
  if (cpEnd <= cpStart) return [];
  const rangesFromPapx = papxRuns
    .filter((run) => run.cpStart < cpEnd && run.cpEnd > cpStart)
    .map((run) => ({
      cpStart: Math.max(cpStart, run.cpStart),
      cpEnd: Math.min(cpEnd, run.cpEnd),
      terminator: documentText[Math.min(cpEnd, run.cpEnd) - 1] || '',
      styleId: run.styleId,
      properties: run.properties,
    }))
    .filter((range) => range.cpEnd > range.cpStart);

  if (rangesFromPapx.length) return rangesFromPapx;

  return splitParagraphRanges(documentText.slice(cpStart, cpEnd)).map((range) => ({
    cpStart: cpStart + range.cpStart,
    cpEnd: cpStart + range.cpEnd,
    terminator: range.terminator,
    styleId: 0,
    properties: [],
  }));
}

function buildParagraphModelsForInterval(
  cpStart: number,
  cpEnd: number,
  documentText: string,
  wordBytes: Uint8Array,
  clx: ReturnType<typeof parseClx>,
  pieceTexts: string[],
  styles: StyleCollection,
  fonts: FontsCollection,
  papxRuns: PapxRun[],
  chpxRuns: ChpxRun[],
  resolveAsset: (charState: CharState) => MsDocAsset | null,
  revisionAuthors: string[] = [],
  inlineContext: InlineBuildContext = {},
): ParagraphModel[] {
  const ranges = buildRangesForCpInterval(cpStart, cpEnd, documentText, papxRuns);
  const chpxCursor = { index: 0 };
  return ranges.map((range) => {
    const rawParagraphText = getTextByCp(wordBytes, clx, pieceTexts, range.cpStart, range.cpEnd);
    const terminatorCandidate = range.terminator === DOC_CONTROL.paragraph || range.terminator === DOC_CONTROL.cellMark ? range.terminator : '';
    const paragraphText = terminatorCandidate && rawParagraphText.endsWith(terminatorCandidate)
      ? rawParagraphText.slice(0, -1)
      : rawParagraphText;
    return buildParagraphModel(
      { ...range, terminator: terminatorCandidate },
      paragraphText,
      styles,
      fonts,
      chpxRuns,
      resolveAsset,
      chpxCursor,
      revisionAuthors,
      inlineContext,
    );
  });
}

function parseIntervalToContent(
  cpStart: number,
  cpEnd: number,
  documentText: string,
  wordBytes: Uint8Array,
  clx: ReturnType<typeof parseClx>,
  pieceTexts: string[],
  styles: StyleCollection,
  fonts: FontsCollection,
  papxRuns: PapxRun[],
  chpxRuns: ChpxRun[],
  resolveAsset: (charState: CharState) => MsDocAsset | null,
  revisionAuthors: string[] = [],
  inlineContext: InlineBuildContext = {},
): ParsedStoryContent {
  const paragraphs = buildParagraphModelsForInterval(
    cpStart,
    cpEnd,
    documentText,
    wordBytes,
    clx,
    pieceTexts,
    styles,
    fonts,
    papxRuns,
    chpxRuns,
    resolveAsset,
    revisionAuthors,
    inlineContext,
  );
  return {
    paragraphs,
    blocks: buildBlocks(paragraphs),
    text: paragraphs.map((paragraph) => paragraph.text).filter(Boolean).join('\n'),
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
    cpStart: paragraph.cpStart,
    cpEnd: paragraph.cpEnd,
    styleId: paragraph.styleId,
    styleName: paragraph.styleName,
    paraState: paragraph.paraState,
    inlines: paragraph.inlines,
    text: paragraph.text,
  };
}

interface RawTableRow {
  cells: TableCellBlock[];
  rowEndParagraph: ParagraphModel | null;
  paragraphs: ParagraphModel[];
}

function cloneTableStateValue<TValue>(value: TValue): TValue {
  if (Array.isArray(value)) return value.map((item) => cloneTableStateValue(item)) as TValue;
  if (!value || typeof value !== 'object') return value;
  const out: Record<string, unknown> = {};
  for (const [key, entry] of Object.entries(value as Record<string, unknown>)) out[key] = cloneTableStateValue(entry);
  return out as TValue;
}

function cloneTableState(state: TableState | null | undefined): TableState {
  const base = tablePropsToState([]);
  if (!state) return base;
  return {
    ...base,
    ...cloneTableStateValue(state),
    operations: [...(state.operations || [])],
  };
}

function isEmptyParagraphBlock(block: ParagraphBlock): boolean {
  if (!block.text && !(block.inlines || []).length) return true;
  return !String(block.text || '').trim() && !(block.inlines || []).some((inline) => inline.type !== 'text' || String((inline as { text?: string }).text || '').trim());
}

function cellHasRenderableContent(cell: TableCellBlock | null | undefined): boolean {
  if (!cell) return false;
  return cell.paragraphs.some((block) => !isEmptyParagraphBlock(block));
}

function createEmptyCellBlock(): TableCellBlock {
  return {
    id: uniqueId('cell'),
    paragraphs: [],
    meta: null,
  };
}

function tableStateScore(state: TableState | null | undefined): number {
  if (!state) return -1;
  const cellCount = applyTableStateToCells(state).length;
  let score = cellCount * 10;
  if (state.defTable?.cells?.length) score += 1000;
  if (state.tableWidth?.wWidth) score += 25;
  score += state.operations?.length || 0;
  return score;
}

function buildEffectiveTableState(templateState: TableState | null, rowState: TableState | null): TableState {
  const effective = cloneTableState(templateState);
  if (!rowState) return effective;
  if (rowState.styleId != null) effective.styleId = rowState.styleId;
  if (rowState.tableWidth) effective.tableWidth = cloneTableStateValue(rowState.tableWidth);
  if (rowState.widthBefore != null) effective.widthBefore = cloneTableStateValue(rowState.widthBefore);
  if (rowState.widthAfter != null) effective.widthAfter = cloneTableStateValue(rowState.widthAfter);
  if (rowState.cellSpacing) effective.cellSpacing = cloneTableStateValue(rowState.cellSpacing);
  if (rowState.defTable) effective.defTable = cloneTableStateValue(rowState.defTable);
  if (rowState.leftIndent) effective.leftIndent = rowState.leftIndent;
  if (rowState.gapHalf) effective.gapHalf = rowState.gapHalf;
  if (rowState.rowHeight) effective.rowHeight = rowState.rowHeight;
  if (rowState.absLeft != null) effective.absLeft = rowState.absLeft;
  if (rowState.absTop != null) effective.absTop = rowState.absTop;
  if (rowState.distanceLeft != null) effective.distanceLeft = rowState.distanceLeft;
  if (rowState.distanceTop != null) effective.distanceTop = rowState.distanceTop;
  if (rowState.positionCode != null) effective.positionCode = rowState.positionCode;
  if (rowState.autoFit != null) effective.autoFit = cloneTableStateValue(rowState.autoFit);
  effective.alignment = rowState.alignment;
  effective.cantSplit = rowState.cantSplit || effective.cantSplit;
  effective.header = rowState.header || effective.header;
  effective.rtl = rowState.rtl || effective.rtl;
  effective.operations = [...(templateState?.operations || []), ...(rowState.operations || [])];
  return effective;
}

function inferExpectedColumnCount(rawRows: RawTableRow[], explicitCellCount: number): number {
  if (explicitCellCount > 0) return explicitCellCount;
  let best = 0;
  for (const row of rawRows) {
    let count = row.cells.length;
    while (count > 0 && !cellHasRenderableContent(row.cells[count - 1])) count -= 1;
    best = Math.max(best, count || row.cells.length);
  }
  return Math.max(best, 1);
}

function inferColumnMeta(rawRows: RawTableRow[], expectedColumnCount: number, tableWidthTwips: number): TableCellBlock['meta'][] {
  const safeCount = Math.max(expectedColumnCount, 1);
  const weights = new Array(safeCount).fill(1);
  for (const row of rawRows) {
    const cells = row.cells.slice(0, safeCount);
    for (let index = 0; index < cells.length; index += 1) {
      const textLength = cells[index]!.paragraphs.map((block) => block.text.trim().length).join('').length;
      if (textLength > 0) weights[index] = Math.max(weights[index], Math.min(textLength, 32));
    }
  }
  const totalWeight = weights.reduce((sum, value) => sum + value, 0) || safeCount;
  const width = tableWidthTwips > 0 ? tableWidthTwips : safeCount * 1440;
  let cursor = 0;
  return weights.map((weight, index) => {
    const nextCursor = index === safeCount - 1 ? width : cursor + Math.round((width * weight) / totalWeight);
    const meta = {
      index,
      width: nextCursor - cursor,
      leftBoundary: cursor,
      rightBoundary: nextCursor,
      borders: {},
    };
    cursor = nextCursor;
    return meta;
  });
}

function applyEdgeVerticalMergeHints(rows: TableRowBlock[]): void {
  for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex]!;
    const contentIndices = row.cells.map((cell, index) => (cellHasRenderableContent(cell) ? index : -1)).filter((index) => index >= 0);
    if (!contentIndices.length || contentIndices.length === row.cells.length) continue;
    const firstContent = contentIndices[0]!;
    const lastContent = contentIndices[contentIndices.length - 1]!;
    const candidates = [
      ...Array.from({ length: firstContent }, (_, index) => index),
      ...Array.from({ length: Math.max(0, row.cells.length - lastContent - 1) }, (_, index) => lastContent + 1 + index),
    ];
    for (const columnIndex of candidates) {
      const cell = row.cells[columnIndex];
      if (!cell || cellHasRenderableContent(cell)) continue;
      const above = rows[rowIndex - 1]!.cells[columnIndex];
      if (!above) continue;
      if ((above.meta?.vertMerge || 0) === 1) {
        cell.meta = { ...(cell.meta || { index: columnIndex }), vertMerge: 1 };
        continue;
      }
      if (!cellHasRenderableContent(above)) continue;
      above.meta = { ...(above.meta || { index: columnIndex }), vertMerge: Math.max(above.meta?.vertMerge || 0, 2) };
      cell.meta = { ...(cell.meta || { index: columnIndex }), vertMerge: 1 };
    }
  }
}

function isTableCandidateParagraph(paragraph: ParagraphModel): boolean {
  return paragraph.terminator === DOC_CONTROL.cellMark
    || paragraph.paraState.inTable
    || paragraph.paraState.innerTableCell
    || paragraph.paraState.tableRowEnd
    || paragraph.paraState.innerTableRowEnd;
}

function buildTableBlock(tableParagraphs: ParagraphModel[]): TableBlock {
  const rawRows: RawTableRow[] = [];
  let pendingRowCells: TableCellBlock[] = [];
  let pendingRowParagraphs: ParagraphModel[] = [];
  let pendingCellParagraphs: ParagraphModel[] = [];

  const flushCell = (): void => {
    if (!pendingCellParagraphs.length) return;
    pendingRowCells.push({
      id: uniqueId('cell'),
      paragraphs: pendingCellParagraphs.map(paragraphToBlock),
      meta: null,
    });
    pendingCellParagraphs = [];
  };

  const flushRow = (rowEndParagraph: ParagraphModel | null): void => {
    if (pendingCellParagraphs.length) flushCell();
    if (!pendingRowCells.length && !pendingRowParagraphs.length) return;
    rawRows.push({
      cells: pendingRowCells,
      rowEndParagraph,
      paragraphs: pendingRowParagraphs,
    });
    pendingRowCells = [];
    pendingRowParagraphs = [];
  };

  for (const paragraph of tableParagraphs) {
    pendingRowParagraphs.push(paragraph);
    pendingCellParagraphs.push(paragraph);

    const endsCell = paragraph.terminator === DOC_CONTROL.cellMark || paragraph.paraState.tableRowEnd || paragraph.paraState.innerTableRowEnd;
    const endsRow = paragraph.paraState.tableRowEnd || paragraph.paraState.innerTableRowEnd;

    if (endsCell) flushCell();
    if (endsRow) flushRow(paragraph);
  }
  flushRow(null);

  const templateState = rawRows.reduce<TableState | null>((best, row) => {
    const candidate = row.rowEndParagraph?.tableState || null;
    return tableStateScore(candidate) >= tableStateScore(best) ? candidate : best;
  }, null);
  const templateCellMeta = templateState ? applyTableStateToCells(templateState) : [];
  const expectedColumnCount = inferExpectedColumnCount(rawRows, templateCellMeta.length);
  const inferredCellMeta = inferColumnMeta(rawRows, expectedColumnCount, templateState?.tableWidth?.wWidth || 0);

  const rows: TableRowBlock[] = rawRows.map((rawRow) => {
    const rowState = buildEffectiveTableState(templateState, rawRow.rowEndParagraph?.tableState || null);
    let cells = rawRow.cells.map((cell) => ({ ...cell, meta: cell.meta ? cloneTableStateValue(cell.meta) : null }));

    while (cells.length > expectedColumnCount && !cellHasRenderableContent(cells[cells.length - 1])) cells.pop();
    while (cells.length < expectedColumnCount) cells.push(createEmptyCellBlock());

    const rowCellMeta = applyTableStateToCells(rowState);
    const cellMetaSource = rowCellMeta.length ? rowCellMeta : templateCellMeta.length ? templateCellMeta : inferredCellMeta;
    cells.forEach((cell, index) => {
      cell.meta = cloneTableStateValue(cellMetaSource[index] || inferredCellMeta[index] || { index });
    });

    const gridWidthTwips = cellMetaSource.length
      ? Math.max(0, (cellMetaSource[cellMetaSource.length - 1]!.rightBoundary || 0) - (cellMetaSource[0]!.leftBoundary || 0))
      : (rowState.tableWidth?.wWidth || 0);

    return {
      id: uniqueId('row'),
      cells,
      state: rowState,
      gridWidthTwips,
    };
  });

  applyEdgeVerticalMergeHints(rows);
  finalizeTableGrid(rows);

  const firstParagraph = tableParagraphs[0];
  const lastParagraph = tableParagraphs[tableParagraphs.length - 1];
  const gridWidthTwips = rows.find((row) => row.gridWidthTwips)?.gridWidthTwips || templateState?.tableWidth?.wWidth || 0;
  const depthCandidates = rawRows.map((row) => getTableDepth(row.rowEndParagraph?.paraState || tableParagraphs[0]!.paraState)).filter((value) => value > 0);
  const depth = depthCandidates.length ? Math.max(...depthCandidates) : 1;

  return {
    type: 'table',
    id: uniqueId('table'),
    cpStart: firstParagraph?.cpStart || 0,
    cpEnd: lastParagraph?.cpEnd || firstParagraph?.cpEnd || 0,
    depth,
    rows,
    state: rows[0]?.state || templateState || tablePropsToState([]),
    gridWidthTwips,
  };
}

function buildBlocks(paragraphs: ParagraphModel[]): Array<ParagraphBlock | TableBlock> {
  const blocks: Array<ParagraphBlock | TableBlock> = [];
  let index = 0;
  while (index < paragraphs.length) {
    const paragraph = paragraphs[index]!;
    if (!isTableCandidateParagraph(paragraph)) {
      blocks.push(paragraphToBlock(paragraph));
      index += 1;
      continue;
    }

    const tableParagraphs: ParagraphModel[] = [];
    while (index < paragraphs.length && isTableCandidateParagraph(paragraphs[index]!)) {
      tableParagraphs.push(paragraphs[index]!);
      index += 1;
    }

    if (tableParagraphs.some((item) => item.terminator === DOC_CONTROL.cellMark || item.paraState.tableRowEnd || item.paraState.innerTableRowEnd)) {
      blocks.push(buildTableBlock(tableParagraphs));
    } else {
      for (const item of tableParagraphs) blocks.push(paragraphToBlock(item));
    }
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

function buildNoteItems(
  kind: 'footnote' | 'endnote',
  storyCpBase: number,
  textEntries: Array<{ index: number; cpStart: number; cpEnd: number }>,
  refEntries: Array<{ index: number; cpStart: number }>,
  parseContent: (cpStart: number, cpEnd: number, context?: InlineBuildContext) => ParsedStoryContent,
): { items: NoteItem[]; refMap: Map<number, NoteReferenceInfo> } {
  const items: NoteItem[] = [];
  const refMap = new Map<number, NoteReferenceInfo>();
  for (let i = 0; i < textEntries.length; i += 1) {
    const entry = textEntries[i]!;
    if (entry.cpEnd <= entry.cpStart) continue;
    const id = uniqueId(kind);
    const label = String(i + 1);
    const refCp = refEntries[i]?.cpStart;
    if (refCp != null) refMap.set(refCp, { kind, id, label });
    const content = parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd);
    items.push({
      id,
      kind,
      index: i,
      label,
      refCp,
      blocks: content.blocks,
      text: content.text,
    });
  }
  return { items, refMap };
}

function buildCommentItems(
  storyCpBase: number,
  textEntries: Array<{ index: number; cpStart: number; cpEnd: number }>,
  refEntries: Array<{ index: number; cpStart: number; data: Uint8Array }>,
  commentAuthors: string[],
  revisionAuthors: string[],
  parseContent: (cpStart: number, cpEnd: number, context?: InlineBuildContext) => ParsedStoryContent,
): { items: CommentItem[]; refMap: Map<number, CommentReferenceInfo> } {
  const items: CommentItem[] = [];
  const refMap = new Map<number, CommentReferenceInfo>();
  for (let i = 0; i < textEntries.length; i += 1) {
    const entry = textEntries[i]!;
    if (entry.cpEnd <= entry.cpStart) continue;
    const meta = parseCommentRefMeta(refEntries[i]?.data || new Uint8Array(0));
    const author = commentAuthors[meta.authorIndex] || revisionAuthors[meta.authorIndex] || meta.initials || undefined;
    const id = uniqueId('comment');
    const label = String(i + 1);
    const refCp = refEntries[i]?.cpStart;
    if (refCp != null) refMap.set(refCp, { id, label, author });
    const content = parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd);
    items.push({
      id,
      index: i,
      label,
      refCp,
      author,
      initials: meta.initials || undefined,
      bookmarkId: meta.bookmarkId,
      blocks: content.blocks,
      text: content.text,
    });
  }
  return { items, refMap };
}

function buildTextboxItems(
  header: boolean,
  storyCpBase: number,
  entries: Array<{ index: number; cpStart: number; cpEnd: number; data: Uint8Array }>,
  parseContent: (cpStart: number, cpEnd: number, context?: InlineBuildContext) => ParsedStoryContent,
  shapeById: Map<number, ShapeAnchorInfo> = new Map(),
): TextboxItem[] {
  const items: TextboxItem[] = [];
  for (const entry of entries) {
    const meta = parseTextboxMeta(entry.data);
    const content = entry.cpEnd > entry.cpStart
      ? parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd)
      : { paragraphs: [], blocks: [], text: '' };
    if (!content.blocks.length && !content.text && meta.reusable) continue;
    const item: TextboxItem = {
      id: uniqueId(header ? 'hdr-textbox' : 'textbox'),
      index: entry.index,
      label: `${header ? 'Header textbox' : 'Textbox'} ${entry.index + 1}`,
      header,
      reusable: meta.reusable,
      shapeId: meta.shapeId || undefined,
      shape: meta.shapeId ? shapeById.get(meta.shapeId) : undefined,
      blocks: content.blocks,
      text: content.text,
    };
    if (item.shape) item.shape.matchedTextboxId = item.id;
    items.push(item);
  }
  return items;
}

function buildHeaderStories(
  descriptors: ReturnType<typeof buildHeaderStoryDescriptors>,
  parseContent: (cpStart: number, cpEnd: number, context?: InlineBuildContext) => ParsedStoryContent,
): HeaderFooterStory[] {
  const stories: HeaderFooterStory[] = [];
  const latestByRole = new Map<string, HeaderFooterStory>();
  for (const descriptor of descriptors) {
    let blocks: HeaderFooterStory['blocks'] = [];
    let text = '';
    if (descriptor.cpEnd > descriptor.cpStart) {
      const content = parseContent(descriptor.cpStart, descriptor.cpEnd);
      blocks = content.blocks;
      text = content.text;
    } else if (descriptor.inheritedFromSection != null) {
      const inherited = latestByRole.get(descriptor.role);
      if (inherited) {
        blocks = inherited.blocks;
        text = inherited.text;
      }
    }
    if (!blocks.length && !text) continue;
    const story: HeaderFooterStory = {
      id: uniqueId('header-story'),
      role: descriptor.role,
      roleLabel: descriptor.roleLabel,
      sectionIndex: descriptor.sectionIndex,
      inheritedFromSection: descriptor.inheritedFromSection,
      blocks,
      text,
    };
    stories.push(story);
    if (text || blocks.length) latestByRole.set(descriptor.role, story);
  }
  return stories;
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
  const storyWindows = buildStoryWindows(fib.fibRgLw, documentText.length);

  const styles = parseStyles(tableBytes, fib.fibRgFcLcb);
  const fonts = parseFonts(tableBytes, fib.fibRgFcLcb);
  const revisionAuthors = parseSttbfRMark(tableBytes, fib.fibRgFcLcb);
  const commentAuthors = parseXstArray(tableBytes, fib.fibRgFcLcb);
  const chpxRuns = readChpxRuns(wordBytes, tableBytes, fib, clx)
    .filter((run) => run.cpStart < documentText.length)
    .map((run) => ({ ...run, cpEnd: Math.min(run.cpEnd, documentText.length) }));
  const papxRuns = readPapxRuns(wordBytes, tableBytes, fib, clx)
    .filter((run) => run.cpStart < documentText.length)
    .map((run) => ({ ...run, cpEnd: Math.min(run.cpEnd, documentText.length) }));

  const objectPool = extractObjectPool(cfb);
  const assets: MsDocAsset[] = [];
  const usedAttachmentNames = new Set<string>();
  const assetCache = new Map<number, MsDocAsset | null>();
  const resolveAsset = createAssetResolver(dataBytes, objectPool, assets, usedAttachmentNames, assetCache, options);

  const parseContent = (cpStart: number, cpEnd: number, inlineContext: InlineBuildContext = {}): ParsedStoryContent => parseIntervalToContent(
    cpStart,
    cpEnd,
    documentText,
    wordBytes,
    clx,
    pieceTexts,
    styles,
    fonts,
    papxRuns,
    chpxRuns,
    resolveAsset,
    revisionAuthors,
    inlineContext,
  );

  const footnoteTextEntries = readFixedPlc(
    tableBytes,
    fib.fibRgFcLcb.fcPlcffndTxt as number | undefined,
    fib.fibRgFcLcb.lcbPlcffndTxt as number | undefined,
    0,
  );
  const footnoteRefEntries = readFixedPlc(
    tableBytes,
    fib.fibRgFcLcb.fcPlcffndRef as number | undefined,
    fib.fibRgFcLcb.lcbPlcffndRef as number | undefined,
    2,
  ).map((entry) => ({ index: entry.index, cpStart: entry.cpStart }));
  const { items: footnoteItems, refMap: footnoteRefMap } = buildNoteItems(
    'footnote',
    storyWindows.footnote.cpStart,
    footnoteTextEntries,
    footnoteRefEntries,
    parseContent,
  );

  const endnoteTextEntries = readFixedPlc(
    tableBytes,
    fib.fibRgFcLcb.fcPlcfendTxt as number | undefined,
    fib.fibRgFcLcb.lcbPlcfendTxt as number | undefined,
    0,
  );
  const endnoteRefEntries = readFixedPlc(
    tableBytes,
    fib.fibRgFcLcb.fcPlcfendRef as number | undefined,
    fib.fibRgFcLcb.lcbPlcfendRef as number | undefined,
    2,
  ).map((entry) => ({ index: entry.index, cpStart: entry.cpStart }));
  const { items: endnoteItems, refMap: endnoteRefMap } = buildNoteItems(
    'endnote',
    storyWindows.endnote.cpStart,
    endnoteTextEntries,
    endnoteRefEntries,
    parseContent,
  );

  const commentTextEntries = readFixedPlc(
    tableBytes,
    fib.fibRgFcLcb.fcPlcfandTxt as number | undefined,
    fib.fibRgFcLcb.lcbPlcfandTxt as number | undefined,
    0,
  );
  const commentRefEntries = readFixedPlc(
    tableBytes,
    fib.fibRgFcLcb.fcPlcfandRef as number | undefined,
    fib.fibRgFcLcb.lcbPlcfandRef as number | undefined,
    30,
  );
  const { items: commentItems, refMap: commentRefMap } = buildCommentItems(
    storyWindows.comment.cpStart,
    commentTextEntries,
    commentRefEntries,
    commentAuthors,
    revisionAuthors,
    parseContent,
  );

  const noteRefs = new Map<number, NoteReferenceInfo>();
  for (const [cp, info] of footnoteRefMap.entries()) noteRefs.set(cp, info);
  for (const [cp, info] of endnoteRefMap.entries()) noteRefs.set(cp, info);

  const mainContent = parseContent(storyWindows.main.cpStart, storyWindows.main.cpEnd, {
    noteRefs,
    commentRefs: commentRefMap,
  });

  const mainShapeAnchors = readShapeAnchors(tableBytes, fib.fibRgFcLcb, storyWindows.main.cpStart, 'main');
  const headerShapeAnchors = readShapeAnchors(tableBytes, fib.fibRgFcLcb, storyWindows.header.cpStart, 'header');
  const mainShapeById = new Map(mainShapeAnchors.map((anchor) => [anchor.shapeId, anchor]));
  const headerShapeById = new Map(headerShapeAnchors.map((anchor) => [anchor.shapeId, anchor]));

  const headerStories = buildHeaderStories(
    buildHeaderStoryDescriptors(tableBytes, fib.fibRgFcLcb, storyWindows.header),
    parseContent,
  );

  const textboxItems = buildTextboxItems(
    false,
    storyWindows.textbox.cpStart,
    readFixedPlc(
      tableBytes,
      fib.fibRgFcLcb.fcPlcftxbxTxt as number | undefined,
      fib.fibRgFcLcb.lcbPlcftxbxTxt as number | undefined,
      22,
    ),
    parseContent,
    mainShapeById,
  );
  const headerTextboxItems = buildTextboxItems(
    true,
    storyWindows.headerTextbox.cpStart,
    readFixedPlc(
      tableBytes,
      fib.fibRgFcLcb.fcPlcfHdrtxbxTxt as number | undefined,
      fib.fibRgFcLcb.lcbPlcfHdrtxbxTxt as number | undefined,
      22,
    ),
    parseContent,
    headerShapeById,
  );
  const floatingShapes = mainShapeAnchors.filter((anchor) => !anchor.matchedTextboxId);
  const headerFloatingShapes = headerShapeAnchors.filter((anchor) => !anchor.matchedTextboxId);
  if (floatingShapes.length || headerFloatingShapes.length) {
    pushWarning(warnings, 'Floating shape anchors were parsed and exposed as structured metadata cards when no textbox story was available', {
      code: 'floating-shapes-partial-render',
      severity: 'info',
      details: { mainShapes: floatingShapes.length, headerShapes: headerFloatingShapes.length },
    });
  }

  const blocks: MsDocParseResult['blocks'] = [...mainContent.blocks];
  if (footnoteItems.length) {
    const footnotesBlock: NotesBlock = { type: 'notes', id: uniqueId('notes-footnote'), kind: 'footnote', items: footnoteItems };
    blocks.push(footnotesBlock);
  }
  if (endnoteItems.length) {
    const endnotesBlock: NotesBlock = { type: 'notes', id: uniqueId('notes-endnote'), kind: 'endnote', items: endnoteItems };
    blocks.push(endnotesBlock);
  }
  if (commentItems.length) {
    const commentsBlock: CommentsBlock = { type: 'comments', id: uniqueId('comments'), items: commentItems };
    blocks.push(commentsBlock);
  }
  if (headerStories.length) {
    const headersBlock: HeadersBlock = { type: 'headers', id: uniqueId('headers'), stories: headerStories };
    blocks.push(headersBlock);
  }
  if (textboxItems.length) {
    const textboxesBlock: TextboxesBlock = { type: 'textboxes', id: uniqueId('textboxes'), header: false, items: textboxItems };
    blocks.push(textboxesBlock);
  }
  if (headerTextboxItems.length) {
    const headerTextboxesBlock: TextboxesBlock = { type: 'textboxes', id: uniqueId('header-textboxes'), header: true, items: headerTextboxItems };
    blocks.push(headerTextboxesBlock);
  }
  if (floatingShapes.length) {
    const shapesBlock: ShapesBlock = { type: 'shapes', id: uniqueId('shapes'), header: false, items: floatingShapes };
    blocks.push(shapesBlock);
  }
  if (headerFloatingShapes.length) {
    const headerShapesBlock: ShapesBlock = { type: 'shapes', id: uniqueId('header-shapes'), header: true, items: headerFloatingShapes };
    blocks.push(headerShapesBlock);
  }

  const trailingAttachments = Array.from(objectPool.values())
    .filter((item) => item?.attachment && !usedAttachmentNames.has(item.entry.name))
    .map((item) => item.attachment as AttachmentAsset);

  for (const attachment of trailingAttachments) assets.push(attachment);
  if (trailingAttachments.length) {
    blocks.push({ type: 'attachments', id: uniqueId('attachments'), items: trailingAttachments });
  }

  collectAssetWarnings(assets, warnings);

  const countInnerBlocks = (innerBlocks: Array<ParagraphBlock | TableBlock>): number => innerBlocks.reduce((sum, block) => {
    if (block.type === 'paragraph') return sum + 1;
    return sum + block.rows.reduce((rowSum, row) => rowSum + row.cells.reduce((cellSum, cell) => cellSum + cell.paragraphs.length, 0), 0);
  }, 0);
  const paragraphCount = blocks.reduce((sum, block) => {
    if (block.type === 'paragraph' || block.type === 'table') {
      return sum + countInnerBlocks([block]);
    }
    if (block.type === 'notes') {
      return sum + block.items.reduce((itemSum, item) => itemSum + countInnerBlocks(item.blocks), 0);
    }
    if (block.type === 'comments') {
      return sum + block.items.reduce((itemSum, item) => itemSum + countInnerBlocks(item.blocks), 0);
    }
    if (block.type === 'headers') {
      return sum + block.stories.reduce((storySum, story) => storySum + countInnerBlocks(story.blocks), 0);
    }
    if (block.type === 'textboxes') {
      return sum + block.items.reduce((itemSum, item) => itemSum + countInnerBlocks(item.blocks), 0);
    }
    return sum;
  }, 0);

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
        paragraphs: paragraphCount,
        blocks: blocks.length,
        assets: assets.length,
        styles: styles.styles.size,
        fonts: fonts.fonts.length,
        footnotes: footnoteItems.length,
        endnotes: endnoteItems.length,
        comments: commentItems.length,
        headers: headerStories.length,
        textboxes: textboxItems.length,
        headerTextboxes: headerTextboxItems.length,
        shapes: mainShapeAnchors.length,
        headerShapes: headerShapeAnchors.length,
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
