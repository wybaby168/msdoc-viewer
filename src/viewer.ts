import { twipsToPx } from './core/utils.js';
import { parseMsDoc } from './msdoc/parser.js';
import { renderBlockList, renderMsDoc } from './render/html.js';
import type {
  HeaderFooterStory,
  ImageAsset,
  MsDocBlock,
  MsDocParseResult,
  MsDocParseToHtmlOptions,
  MsDocRenderResult,
  MsDocViewer,
  MsDocViewerConfig,
  MsDocViewerLoadOptions,
  ParagraphBlock,
  SectionDescriptor,
  ShapeAnchorInfo,
  TableBlock,
  TextboxItem,
  ViewerInput,
} from './types.js';

async function normalizeInput(input: ViewerInput): Promise<ArrayBuffer> {
  if (input instanceof ArrayBuffer) return input;
  if (ArrayBuffer.isView(input)) {
    const bytes = new Uint8Array(input.byteLength);
    bytes.set(new Uint8Array(input.buffer, input.byteOffset, input.byteLength));
    return bytes.buffer;
  }
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return await input.arrayBuffer();
  }
  if (typeof input === 'string') {
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Failed to fetch document: ${response.status}`);
    return await response.arrayBuffer();
  }
  throw new TypeError('Unsupported input type');
}

interface RuntimePageModel {
  sectionIndex: number;
  physicalPageNumber: number;
  sectionPageIndex: number;
  displayPageNumber: number;
  blank?: boolean;
  blocks: HTMLElement[];
}

interface ChromeTemplate {
  source: 'story' | 'textbox';
  kind: 'header' | 'footer';
  role?: HeaderFooterStory['role'];
  roleLabel?: string;
  sectionIndex?: number;
  item?: TextboxItem;
  story?: HeaderFooterStory;
  pageNumberLike: boolean;
}

function runtimeViewerCss(): string {
  return `
.msdoc-viewer-shell{display:flex;flex-direction:column;gap:12px}
.msdoc-viewer-toolbar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:8px 0}
.msdoc-viewer-title{font-weight:600;color:#111827;margin-right:8px}
.msdoc-viewer-button{appearance:none;border:1px solid #d0d7de;background:#fff;color:#111827;border-radius:999px;padding:6px 12px;font:600 13px/1.2 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;cursor:pointer}
.msdoc-viewer-button[aria-pressed="true"]{background:#111827;color:#fff;border-color:#111827}
.msdoc-viewer-hint{font-size:12px;color:#6b7280}
.msdoc-root{position:relative}
.msdoc-root[data-msdoc-view="paged"] .msdoc-flow-view{display:none}
.msdoc-root[data-msdoc-view="flow"] .msdoc-paged-view{display:none}
.msdoc-flow-view{min-width:0}
.msdoc-paged-view{display:flex;flex-direction:column;gap:24px;align-items:center;padding:24px;background:#f3f4f6;border-radius:16px}
.msdoc-page{position:relative;background:#fff;border:1px solid #d1d5db;box-shadow:0 16px 40px rgba(15,23,42,.12);overflow:hidden}
.msdoc-page-label{position:absolute;top:10px;right:12px;z-index:7;font:600 11px/1.2 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#6b7280;background:rgba(255,255,255,.92);padding:3px 8px;border-radius:999px;border:1px solid #e5e7eb;display:none}
.msdoc-page-header-band,.msdoc-page-footer-band{position:absolute;left:0;right:0;z-index:5;padding:0 8px;overflow:visible}
.msdoc-page-header-band{top:0}
.msdoc-page-footer-band{bottom:0}
.msdoc-page-header-band-inner,.msdoc-page-footer-band-inner{position:relative}
.msdoc-page-header-band .msdoc-paragraph,.msdoc-page-footer-band .msdoc-paragraph{margin-bottom:0}
.msdoc-page-body{position:absolute;z-index:3;overflow:hidden}
.msdoc-page-body-content{height:100%;overflow:hidden}
.msdoc-page-overlay{position:absolute;inset:0;z-index:6;pointer-events:none}
.msdoc-page-overlay-item{position:absolute;pointer-events:auto;max-width:100%;overflow:hidden}
.msdoc-page-overlay-item .msdoc-floating,.msdoc-page-overlay-item .msdoc-story-card{margin:0;box-shadow:none;border-color:#d1d5db;background:transparent}
.msdoc-page-overlay-shape-asset{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;overflow:hidden;pointer-events:none}
.msdoc-page-overlay-shape-asset img,.msdoc-page-overlay-image{display:block;max-width:100%;max-height:100%;width:100%;height:100%;object-fit:contain}
.msdoc-page-overlay-item-content{position:relative;z-index:1}
.msdoc-page-overlay-item-page-number{display:flex;align-items:flex-end;justify-content:flex-end;font-size:12pt}
.msdoc-page-overlay-item-page-number .msdoc-page-number{display:inline-block;min-width:1.5em;text-align:right}
.msdoc-page-guides{position:absolute;inset:0;z-index:1;pointer-events:none}
.msdoc-page-guide-corner{position:absolute;width:18px;height:18px}
.msdoc-page-guide-corner-top-left{border-right:1px solid rgba(59,130,246,.42);border-bottom:1px solid rgba(59,130,246,.42)}
.msdoc-page-guide-corner-top-right{border-left:1px solid rgba(59,130,246,.42);border-bottom:1px solid rgba(59,130,246,.42)}
.msdoc-page-guide-corner-bottom-left{border-right:1px solid rgba(59,130,246,.42);border-top:1px solid rgba(59,130,246,.42)}
.msdoc-page-guide-corner-bottom-right{border-left:1px solid rgba(59,130,246,.42);border-top:1px solid rgba(59,130,246,.42)}
.msdoc-page-measure{position:absolute;left:-99999px;top:0;visibility:hidden;pointer-events:none;overflow:hidden}
.msdoc-page-measure .msdoc-page-label,.msdoc-page-measure .msdoc-page-overlay,.msdoc-page-measure .msdoc-page-guides,.msdoc-page-measure .msdoc-page-header-band,.msdoc-page-measure .msdoc-page-footer-band{display:none}
.msdoc-page-blank{background:linear-gradient(180deg,#fff,#fafafa)}
.msdoc-page-blank-note{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#9ca3af;font:600 14px/1.4 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
@media (max-width: 768px){
  .msdoc-paged-view{padding:12px}
}
`;
}

function normalizeRenderableText(text: unknown): string {
  return String(text ?? '')
    .replace(/[\u0000-\u001f\u007f]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function blockListText(blocks: Array<ParagraphBlock | TableBlock>): string {
  return normalizeRenderableText(blocks.map((block) => {
    if (block.type === 'table') {
      return block.rows
        .map((row) => row.cells.filter((cell) => !cell.hidden).map((cell) => cell.paragraphs.map((paragraph) => paragraph.text).join(' ')).join(' | '))
        .join(' || ');
    }
    return block.text;
  }).join(' '));
}

function blockListHasRenderableContent(blocks: Array<ParagraphBlock | TableBlock>): boolean {
  return blocks.some((block) => {
    if (block.type === 'table') return true;
    if (normalizeRenderableText(block.text)) return true;
    return (block.inlines || []).some((node) => node.type !== 'lineBreak' && node.type !== 'pageBreak');
  });
}

function isPageNumberLike(text: string): boolean {
  const normalized = normalizeRenderableText(text);
  if (!normalized) return false;
  return /^[-–—\s\d]+$/.test(normalized) || /\bPAGE\b/i.test(text);
}

function getSections(parsed: MsDocParseResult): SectionDescriptor[] {
  return parsed.sections?.length ? parsed.sections : [{
    index: 0,
    cpStart: 0,
    cpEnd: parsed.meta.fib.ccpText,
    fcSepx: 0,
    properties: [],
    page: {
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
      pageNumberStart: 1,
      documentGridLinePitchTwips: undefined,
      documentGridMode: undefined,
    },
  }];
}

function getMainBlocks(parsed: MsDocParseResult): Array<ParagraphBlock | TableBlock> {
  return parsed.blocks.filter((block): block is ParagraphBlock | TableBlock => block.type === 'paragraph' || block.type === 'table');
}

function getHeaderStories(parsed: MsDocParseResult): HeaderFooterStory[] {
  return parsed.blocks.find((block): block is Extract<MsDocBlock, { type: 'headers' }> => block.type === 'headers')?.stories || [];
}

function getHeaderTextboxes(parsed: MsDocParseResult): TextboxItem[] {
  return parsed.blocks
    .filter((block): block is Extract<MsDocBlock, { type: 'textboxes' }> => block.type === 'textboxes' && block.header)
    .flatMap((block) => block.items);
}

function getMainTextboxes(parsed: MsDocParseResult): TextboxItem[] {
  return parsed.blocks
    .filter((block): block is Extract<MsDocBlock, { type: 'textboxes' }> => block.type === 'textboxes' && !block.header)
    .flatMap((block) => block.items);
}

function getHeaderShapes(parsed: MsDocParseResult): ShapeAnchorInfo[] {
  return parsed.blocks
    .filter((block): block is Extract<MsDocBlock, { type: 'shapes' }> => block.type === 'shapes' && block.header)
    .flatMap((block) => block.items);
}

function getMainShapes(parsed: MsDocParseResult): ShapeAnchorInfo[] {
  return parsed.blocks
    .filter((block): block is Extract<MsDocBlock, { type: 'shapes' }> => block.type === 'shapes' && !block.header)
    .flatMap((block) => block.items);
}

function buildChromeTemplates(parsed: MsDocParseResult): ChromeTemplate[] {
  const templates: ChromeTemplate[] = [];
  for (const story of getHeaderStories(parsed)) {
    if (!blockListHasRenderableContent(story.blocks)) continue;
    templates.push({
      source: 'story',
      kind: story.role.endsWith('Footer') ? 'footer' : 'header',
      role: story.role,
      roleLabel: story.roleLabel,
      sectionIndex: story.sectionIndex != null ? Math.max(0, story.sectionIndex - 1) : undefined,
      story,
      pageNumberLike: isPageNumberLike(story.text || blockListText(story.blocks)),
    });
  }
  for (const item of getHeaderTextboxes(parsed)) {
    if (!blockListHasRenderableContent(item.blocks)) continue;
    const section = item.sectionIndex;
    const shapeTop = item.shape?.boundsTwips.top ?? 0;
    const pageHeight = section != null ? getSections(parsed)[section]?.page.pageHeightTwips ?? 0 : 0;
    const kind = item.shape?.headerKind
      || (item.shape?.headerRole ? (item.shape.headerRole.endsWith('Footer') ? 'footer' : 'header') : undefined)
      || (item.shape
        ? (shapeTop > (pageHeight / 2) ? 'footer' : 'header')
        : (isPageNumberLike(item.text || blockListText(item.blocks)) ? 'footer' : 'header'));
    templates.push({
      source: 'textbox',
      kind,
      sectionIndex: section,
      item,
      pageNumberLike: isPageNumberLike(item.text || blockListText(item.blocks)),
    });
  }
  return templates;
}

function pickStoryTemplate(templates: ChromeTemplate[], kind: 'header' | 'footer', page: RuntimePageModel, sections: SectionDescriptor[]): ChromeTemplate | null {
  const section = sections[page.sectionIndex] || sections[0];
  const candidates = templates.filter((template) => template.source === 'story' && template.kind === kind && (template.sectionIndex == null || template.sectionIndex === page.sectionIndex));
  if (!candidates.length) return null;
  const isEven = page.physicalPageNumber % 2 === 0;
  const isFirst = page.sectionPageIndex === 0;
  const preferredRoles: HeaderFooterStory['role'][] = kind === 'header'
    ? (isFirst && section?.page.titlePage
      ? ['firstHeader', isEven ? 'evenHeader' : 'oddHeader', isEven ? 'oddHeader' : 'evenHeader']
      : [isEven ? 'evenHeader' : 'oddHeader', isEven ? 'oddHeader' : 'evenHeader', 'firstHeader'])
    : (isFirst && section?.page.titlePage
      ? ['firstFooter', isEven ? 'evenFooter' : 'oddFooter', isEven ? 'oddFooter' : 'evenFooter']
      : [isEven ? 'evenFooter' : 'oddFooter', isEven ? 'oddFooter' : 'evenFooter', 'firstFooter']);
  for (const role of preferredRoles) {
    const hit = candidates.find((candidate) => candidate.role === role);
    if (hit) return hit;
  }
  return candidates[0] || null;
}

function pickTextboxTemplates(templates: ChromeTemplate[], kind: 'header' | 'footer', page: RuntimePageModel): ChromeTemplate[] {
  return templates.filter((template) => template.source === 'textbox' && template.kind === kind && (template.sectionIndex == null || template.sectionIndex === page.sectionIndex));
}

function createMeasureFrame(section: SectionDescriptor): { page: HTMLDivElement; body: HTMLDivElement } {
  const page = document.createElement('div');
  page.className = 'msdoc-page msdoc-page-measure';
  const widthPx = twipsToPx(section.page.pageWidthTwips) || 816;
  const heightPx = twipsToPx(section.page.pageHeightTwips) || 1056;
  page.style.width = `${widthPx}px`;
  page.style.height = `${heightPx}px`;

  const body = document.createElement('div');
  body.className = 'msdoc-page-body';
  applyPageBodyStyle(body, section);
  page.appendChild(body);
  document.body.appendChild(page);
  return { page, body };
}

function sectionPageMetrics(section: SectionDescriptor): {
  widthPx: number;
  heightPx: number;
  marginLeftPx: number;
  marginRightPx: number;
  marginTopPx: number;
  marginBottomPx: number;
  headerTopPx: number;
  footerBottomPx: number;
  bodyWidthPx: number;
  bodyHeightPx: number;
} {
  const widthPx = twipsToPx(section.page.pageWidthTwips) || 816;
  const heightPx = twipsToPx(section.page.pageHeightTwips) || 1056;
  const gutterPx = twipsToPx(section.page.gutterTwips) || 0;
  const marginLeftPx = (twipsToPx(section.page.marginLeftTwips) || 96) + gutterPx;
  const marginRightPx = twipsToPx(section.page.marginRightTwips) || 96;
  const marginTopPx = twipsToPx(section.page.marginTopTwips) || 96;
  const marginBottomPx = twipsToPx(section.page.marginBottomTwips) || 96;
  const headerTopPx = twipsToPx(section.page.headerTopTwips) || 48;
  const footerBottomPx = twipsToPx(section.page.footerBottomTwips) || 48;
  return {
    widthPx,
    heightPx,
    marginLeftPx,
    marginRightPx,
    marginTopPx,
    marginBottomPx,
    headerTopPx,
    footerBottomPx,
    bodyWidthPx: Math.max(120, widthPx - marginLeftPx - marginRightPx),
    bodyHeightPx: Math.max(120, heightPx - marginTopPx - marginBottomPx),
  };
}

function applyPageBodyStyle(body: HTMLElement, section: SectionDescriptor): void {
  const metrics = sectionPageMetrics(section);
  body.style.left = `${metrics.marginLeftPx}px`;
  body.style.top = `${metrics.marginTopPx}px`;
  body.style.width = `${metrics.bodyWidthPx}px`;
  body.style.height = `${metrics.bodyHeightPx}px`;
  if (section.page.columns > 1) {
    body.style.columnCount = String(section.page.columns);
    body.style.columnGap = `${twipsToPx(section.page.columnSpacingTwips) || 48}px`;
  } else {
    body.style.columnCount = 'auto';
    body.style.columnGap = 'normal';
  }
}

function shouldForceNewPageBeforeSection(prevSection: SectionDescriptor | undefined, nextSection: SectionDescriptor | undefined): boolean {
  if (!prevSection || !nextSection) return false;
  if (prevSection.index === nextSection.index) return false;
  if (nextSection.page.breakCode === 0) {
    const prevMetrics = sectionPageMetrics(prevSection);
    const nextMetrics = sectionPageMetrics(nextSection);
    return prevMetrics.widthPx !== nextMetrics.widthPx
      || prevMetrics.heightPx !== nextMetrics.heightPx
      || prevMetrics.marginLeftPx !== nextMetrics.marginLeftPx
      || prevMetrics.marginRightPx !== nextMetrics.marginRightPx
      || prevMetrics.marginTopPx !== nextMetrics.marginTopPx
      || prevMetrics.marginBottomPx !== nextMetrics.marginBottomPx;
  }
  return true;
}

function requiresInsertedBlankPage(nextSection: SectionDescriptor | undefined, nextPhysicalPageNumber: number): boolean {
  if (!nextSection) return false;
  if (nextSection.page.breakCode === 3) return nextPhysicalPageNumber % 2 !== 0;
  if (nextSection.page.breakCode === 4) return nextPhysicalPageNumber % 2 === 0;
  return false;
}

function shouldKeepBlockWithNext(previous: HTMLElement | undefined, next: HTMLElement): boolean {
  if (!previous || !next.classList.contains('msdoc-table')) return false;
  if (previous.dataset.keepNext === '1') return true;
  // Word commonly keeps short heading/title paragraphs with the table that immediately follows.
  // This mirrors the same pagination intent when older producers omit sprmPFKeepFollow
  // but save the heading as a title/heading style before a table.
  return previous.dataset.headingLike === '1';
}

function buildPages(root: HTMLElement, rendered: MsDocRenderResult): RuntimePageModel[] {
  const sections = getSections(rendered.parsed);
  const body = root.querySelector('.msdoc-flow-view .msdoc-body');
  if (!body) return [];
  const blocks = Array.from(body.children).filter((node): node is HTMLElement => node instanceof HTMLElement && node.classList.contains('msdoc-flow-block') && node.dataset.msdocFloating !== '1');
  if (!blocks.length) return [];

  const pages: Array<{ sectionIndex: number; blocks: HTMLElement[]; blank?: boolean; syntheticBlank?: boolean }> = [];
  let currentSectionIndex = Number(blocks[0]?.dataset.sectionIndex || 0);
  let currentSection = sections[currentSectionIndex] || sections[0]!;
  let measure = createMeasureFrame(currentSection);
  let currentBlocks: HTMLElement[] = [];

  const finalize = (): void => {
    pages.push({ sectionIndex: currentSectionIndex, blocks: [...currentBlocks] });
    currentBlocks = [];
    measure.page.remove();
  };

  const resetMeasure = (): void => {
    measure.page.remove();
    currentSection = sections[currentSectionIndex] || sections[0]!;
    measure = createMeasureFrame(currentSection);
  };

  for (const block of blocks) {
    const blockSectionIndex = Number(block.dataset.sectionIndex || currentSectionIndex || 0);
    const nextSection = sections[blockSectionIndex] || currentSection;
    const forcedBefore = block.dataset.pageBreakBefore === '1';

    if (forcedBefore && currentBlocks.length) {
      finalize();
      currentSectionIndex = blockSectionIndex;
      resetMeasure();
    }

    if (blockSectionIndex !== currentSectionIndex && shouldForceNewPageBeforeSection(currentSection, nextSection)) {
      if (currentBlocks.length) finalize();
      currentSectionIndex = blockSectionIndex;
      resetMeasure();
    } else if (blockSectionIndex !== currentSectionIndex) {
      currentSectionIndex = blockSectionIndex;
      resetMeasure();
      for (const existing of currentBlocks) measure.body.appendChild(existing.cloneNode(true));
    }

    const clone = block.cloneNode(true) as HTMLElement;
    measure.body.appendChild(clone);
    const metrics = sectionPageMetrics(currentSection);
    if (measure.body.scrollHeight > metrics.bodyHeightPx + 1 && currentBlocks.length) {
      measure.body.removeChild(clone);
      const previousBlock = currentBlocks[currentBlocks.length - 1];
      if (shouldKeepBlockWithNext(previousBlock, block) && currentBlocks.length > 1) {
        measure.body.lastElementChild?.remove();
        const carriedBlock = currentBlocks.pop()!;
        finalize();
        resetMeasure();
        measure.body.appendChild(carriedBlock.cloneNode(true));
        measure.body.appendChild(clone);
        currentBlocks.push(carriedBlock);
      } else {
        finalize();
        resetMeasure();
        measure.body.appendChild(clone);
      }
    }
    currentBlocks.push(block);

    if (clone.querySelector('.msdoc-page-break')) {
      finalize();
      resetMeasure();
    }
  }

  if (currentBlocks.length) finalize();
  else measure.page.remove();

  const runtimePages: RuntimePageModel[] = [];
  let displayPageNumber = 0;
  const sectionPageCounts = new Map<number, number>();

  const pushRuntimePage = (sectionIndex: number, blocksForPage: HTMLElement[], blank = false): void => {
    const section = sections[sectionIndex] || sections[0]!;
    const sectionPageIndex = sectionPageCounts.get(sectionIndex) || 0;
    if (sectionPageIndex === 0 && section.page.restartPageNumber) {
      displayPageNumber = (section.page.pageNumberStart || 1) - 1;
    }
    displayPageNumber += 1;
    sectionPageCounts.set(sectionIndex, sectionPageIndex + 1);
    runtimePages.push({
      sectionIndex,
      physicalPageNumber: runtimePages.length + 1,
      sectionPageIndex,
      displayPageNumber,
      blank,
      blocks: blocksForPage,
    });
  };

  pages.forEach((pageData, index) => {
    const prev = pages[index - 1];
    const nextSection = sections[pageData.sectionIndex] || sections[0]!;
    if (prev && prev.sectionIndex !== pageData.sectionIndex && requiresInsertedBlankPage(nextSection, runtimePages.length + 1)) {
      pushRuntimePage(prev.sectionIndex, [], true);
    }
    pushRuntimePage(pageData.sectionIndex, pageData.blocks, Boolean(pageData.blank));
  });
  return runtimePages;
}

function shapeToPagePosition(shape: ShapeAnchorInfo, section: SectionDescriptor): { left: number; top: number; width: number; height: number } {
  const metrics = sectionPageMetrics(section);
  const width = Math.max(twipsToPx(shape.boundsTwips.width) || 0, 36);
  const height = Math.max(twipsToPx(shape.boundsTwips.height) || 0, 24);
  const left = (shape.anchorX === 'page' ? 0 : metrics.marginLeftPx) + (twipsToPx(shape.boundsTwips.left) || 0);

  let topBase = shape.anchorY === 'page'
    ? 0
    : shape.anchorY === 'margin'
      ? 0
      : metrics.marginTopPx;

  if (shape.story === 'header' && shape.headerKind === 'header' && shape.anchorY !== 'page') {
    topBase = metrics.headerTopPx;
  } else if (shape.story === 'header' && shape.headerKind === 'footer' && shape.anchorY !== 'page') {
    topBase = Math.max(metrics.heightPx - metrics.footerBottomPx - height, 0);
  }

  const top = topBase + (twipsToPx(shape.boundsTwips.top) || 0);
  return { left, top, width, height };
}

function createCornerGuide(className: string, left: number, top: number): HTMLDivElement {
  const guide = document.createElement('div');
  guide.className = `msdoc-page-guide-corner ${className}`;
  guide.style.left = `${left}px`;
  guide.style.top = `${top}px`;
  return guide;
}

function renderDynamicPageNumber(page: RuntimePageModel): string {
  return `<p class="msdoc-paragraph"><span class="msdoc-page-number">${page.displayPageNumber}</span></p>`;
}

function appendHtml(target: HTMLElement, html: string): void {
  if (!html) return;
  const range = document.createRange();
  const fragment = range.createContextualFragment(html);
  target.appendChild(fragment);
}

function isIgnorableChromeNode(node: Element | null): boolean {
  if (!(node instanceof HTMLElement)) return false;
  if (node.matches('table,.msdoc-table,.msdoc-attachment,.msdoc-image-fallback')) return false;
  if (node.querySelector('table,.msdoc-table,img,.msdoc-image,.msdoc-attachment,.msdoc-image-fallback')) return false;
  const text = normalizeRenderableText(node.textContent || '');
  return !text;
}

function trimChromeWhitespace(container: HTMLElement): void {
  while (isIgnorableChromeNode(container.firstElementChild)) container.firstElementChild?.remove();
  while (isIgnorableChromeNode(container.lastElementChild)) container.lastElementChild?.remove();
}

function finalizeChromeBand(band: HTMLElement, kind: 'header' | 'footer', metrics: ReturnType<typeof sectionPageMetrics>): void {
  const content = band.firstElementChild;
  if (!(content instanceof HTMLElement)) return;
  trimChromeWhitespace(content);
  const minHeight = kind === 'header' ? metrics.marginTopPx : metrics.marginBottomPx;
  const contentHeight = Math.ceil(content.offsetHeight || content.scrollHeight || 0);
  const inset = kind === 'header' ? metrics.headerTopPx : metrics.footerBottomPx;
  band.style.minHeight = `${Math.max(minHeight, inset + contentHeight)}px`;
}

function applyDynamicPageNumber(target: HTMLElement, page: RuntimePageModel, totalPageCount: number): void {
  let replaced = false;
  for (const element of Array.from(target.querySelectorAll<HTMLElement>('.msdoc-field-page'))) {
    element.textContent = String(page.displayPageNumber);
    replaced = true;
  }
  for (const element of Array.from(target.querySelectorAll<HTMLElement>('.msdoc-field-numpages,.msdoc-field-sectionpages'))) {
    element.textContent = String(totalPageCount);
    replaced = true;
  }

  const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
  while (walker.nextNode()) {
    const textNode = walker.currentNode as Text;
    const value = textNode.nodeValue || '';
    if (!value.trim()) continue;
    let nextValue = value;
    nextValue = nextValue.replace(/\b(?:NUMPAGES|SECTIONPAGES)\b(?:\s+\\\*\s+MERGEFORMAT)?/ig, String(totalPageCount));
    nextValue = nextValue.replace(/\bPAGE\b(?:\s+\\\*\s+MERGEFORMAT)?/ig, String(page.displayPageNumber));
    if (nextValue !== value) {
      textNode.nodeValue = nextValue;
      replaced = true;
      continue;
    }
    if (/\d+/.test(value) && !replaced) {
      textNode.nodeValue = value.replace(/\d+/, String(page.displayPageNumber));
      replaced = true;
      break;
    }
  }
  if (!replaced) appendHtml(target, renderDynamicPageNumber(page));
}

function renderChromeContent(template: ChromeTemplate, _page: RuntimePageModel): string {
  if (template.story) return renderBlockList(template.story.blocks);
  if (template.item) return renderBlockList(template.item.blocks);
  return '';
}

function sanitizeImageSource(src: string | undefined | null): string | null {
  const value = String(src || '').trim();
  if (!value) return null;
  if (/^data:image\//i.test(value)) return value;
  if (/^(?:https?:|blob:)/i.test(value)) return value;
  return null;
}

function buildImageAssetMap(rendered: MsDocRenderResult): Map<string, ImageAsset> {
  return new Map(rendered.assets.filter((asset): asset is ImageAsset => asset.type === 'image').map((asset) => [asset.id, asset] as const));
}

function pickHeaderShapesForPage(shapes: ShapeAnchorInfo[], kind: 'header' | 'footer', page: RuntimePageModel, sections: SectionDescriptor[]): ShapeAnchorInfo[] {
  const section = sections[page.sectionIndex] || sections[0];
  const isEven = page.physicalPageNumber % 2 === 0;
  const isFirst = page.sectionPageIndex === 0;
  const preferredRoles: Array<HeaderFooterStory['role']> = kind === 'header'
    ? (isFirst && section?.page.titlePage
      ? ['firstHeader', isEven ? 'evenHeader' : 'oddHeader', isEven ? 'oddHeader' : 'evenHeader']
      : [isEven ? 'evenHeader' : 'oddHeader', isEven ? 'oddHeader' : 'evenHeader', 'firstHeader'])
    : (isFirst && section?.page.titlePage
      ? ['firstFooter', isEven ? 'evenFooter' : 'oddFooter', isEven ? 'oddFooter' : 'evenFooter']
      : [isEven ? 'evenFooter' : 'oddFooter', isEven ? 'oddFooter' : 'evenFooter', 'firstFooter']);
  const candidates = shapes.filter((shape) => (shape.headerKind || 'header') === kind && (shape.sectionIndex == null || shape.sectionIndex === page.sectionIndex));
  const exact = candidates.filter((shape) => shape.headerRole && preferredRoles.includes(shape.headerRole));
  return exact.length ? exact : candidates;
}

function createShapeAssetLayer(shape: ShapeAnchorInfo, asset: ImageAsset): HTMLDivElement | null {
  const src = sanitizeImageSource(asset.sourceUrl) || sanitizeImageSource(asset.dataUrl);
  if (!src || asset.displayable === false) return null;
  const assetEl = document.createElement('div');
  assetEl.className = 'msdoc-page-overlay-shape-asset';
  const img = document.createElement('img');
  img.className = 'msdoc-page-overlay-image';
  img.src = src;
  img.alt = shape.drawingDescription || shape.drawingName || '';
  assetEl.appendChild(img);
  return assetEl;
}

function renderShapeAssetOverlay(shape: ShapeAnchorInfo, asset: ImageAsset, section: SectionDescriptor, positionOverride?: { left: number; top: number; width: number; height: number }): HTMLDivElement | null {
  const position = positionOverride || shapeToPagePosition(shape, section);
  const itemEl = document.createElement('div');
  itemEl.className = 'msdoc-page-overlay-item';
  itemEl.style.left = `${position.left}px`;
  itemEl.style.top = `${position.top}px`;
  itemEl.style.width = `${position.width}px`;
  itemEl.style.height = `${position.height}px`;
  if (shape.behindText) itemEl.style.opacity = '0.92';
  const assetEl = createShapeAssetLayer(shape, asset);
  if (!assetEl) return null;
  itemEl.appendChild(assetEl);
  return itemEl;
}

function findPageIndexForShape(pages: RuntimePageModel[], shape: ShapeAnchorInfo): number {
  const targetSection = shape.sectionIndex ?? 0;
  let fallback = -1;
  for (let index = 0; index < pages.length; index += 1) {
    const page = pages[index]!;
    if (page.sectionIndex !== targetSection || page.blank || !page.blocks.length) continue;
    fallback = index;
    const firstCpStart = Number(page.blocks[0]?.dataset.cpStart || 0);
    const lastCpEnd = Number(page.blocks[page.blocks.length - 1]?.dataset.cpEnd || firstCpStart);
    if (shape.anchorCp <= firstCpStart || shape.anchorCp <= lastCpEnd) return index;
  }
  if (fallback >= 0) return fallback;
  return pages.findIndex((page) => page.sectionIndex === targetSection && !page.blank) || 0;
}

function findAnchorBlockElement(bodyContent: HTMLElement, anchorCp: number): HTMLElement | null {
  const blocks = Array.from(bodyContent.children).filter((node): node is HTMLElement => node instanceof HTMLElement);
  for (const block of blocks) {
    const cpEnd = Number(block.dataset.cpEnd || 0);
    if (!cpEnd || cpEnd < anchorCp) continue;
    return block;
  }
  return blocks.length ? blocks[Math.max(0, blocks.length - 1)]! : null;
}

function shapeToBodyRelativePosition(shape: ShapeAnchorInfo, section: SectionDescriptor, body: HTMLElement, bodyContent: HTMLElement): { left: number; top: number; width: number; height: number } {
  const metrics = sectionPageMetrics(section);
  const anchorBlock = findAnchorBlockElement(bodyContent, shape.anchorCp);
  const paragraphTop = anchorBlock ? body.offsetTop + anchorBlock.offsetTop : body.offsetTop;
  const leftBase = shape.anchorX === 'page'
    ? 0
    : shape.anchorX === 'column'
      ? body.offsetLeft
      : metrics.marginLeftPx;
  const topBase = shape.anchorY === 'page'
    ? 0
    : shape.anchorY === 'margin'
      ? metrics.marginTopPx
      : paragraphTop;
  const left = leftBase + (twipsToPx(shape.boundsTwips.left) || 0);
  const top = topBase + (twipsToPx(shape.boundsTwips.top) || 0);
  const width = Math.max(twipsToPx(shape.boundsTwips.width) || 0, 36);
  const height = Math.max(twipsToPx(shape.boundsTwips.height) || 0, 24);
  return { left, top, width, height };
}

function renderMainShapeOverlay(shape: ShapeAnchorInfo, textbox: TextboxItem | undefined, asset: ImageAsset | undefined, section: SectionDescriptor, body: HTMLElement, bodyContent: HTMLElement): HTMLDivElement | null {
  const position = shapeToBodyRelativePosition(shape, section, body, bodyContent);
  if (textbox && blockListHasRenderableContent(textbox.blocks)) {
    const itemEl = document.createElement('div');
    itemEl.className = 'msdoc-page-overlay-item';
    itemEl.style.left = `${position.left}px`;
    itemEl.style.top = `${position.top}px`;
    itemEl.style.width = `${position.width}px`;
    itemEl.style.minHeight = `${position.height}px`;
    if (shape.behindText) itemEl.style.opacity = '0.92';
    if (asset) {
      const assetEl = createShapeAssetLayer(shape, asset);
      if (assetEl) itemEl.appendChild(assetEl);
    }
    const contentEl = document.createElement('div');
    contentEl.className = 'msdoc-page-overlay-item-content';
    appendHtml(contentEl, renderBlockList(textbox.blocks));
    itemEl.appendChild(contentEl);
    return itemEl;
  }
  if (asset) return renderShapeAssetOverlay(shape, asset, section, position);
  return null;
}

function renderPagedView(root: HTMLElement, rendered: MsDocRenderResult): void {
  const pagedView = root.querySelector('.msdoc-paged-view');
  if (!(pagedView instanceof HTMLElement)) return;
  pagedView.innerHTML = '';

  const sections = getSections(rendered.parsed);
  const chromeTemplates = buildChromeTemplates(rendered.parsed);
  const mainTextboxes = getMainTextboxes(rendered.parsed);
  const mainShapes = getMainShapes(rendered.parsed);
  const headerShapes = getHeaderShapes(rendered.parsed);
  const imageAssets = buildImageAssetMap(rendered);
  const pages = buildPages(root, rendered);
  const mainTextboxById = new Map(mainTextboxes.map((item) => [item.id, item] as const));
  const mainTextboxByShapeId = new Map(mainTextboxes.filter((item) => item.shapeId != null).map((item) => [item.shapeId as number, item] as const));

  for (const [pageIndex, page] of pages.entries()) {
    const section = sections[page.sectionIndex] || sections[0]!;
    const metrics = sectionPageMetrics(section);
    const pageEl = document.createElement('article');
    pageEl.className = `msdoc-page${page.blank ? ' msdoc-page-blank' : ''}`;
    pageEl.style.width = `${metrics.widthPx}px`;
    pageEl.style.height = `${metrics.heightPx}px`;
    pageEl.dataset.sectionIndex = String(page.sectionIndex);

    const label = document.createElement('div');
    label.className = 'msdoc-page-label';
    label.textContent = `Page ${page.displayPageNumber}`;
    pageEl.appendChild(label);

    const guides = document.createElement('div');
    guides.className = 'msdoc-page-guides';
    guides.appendChild(createCornerGuide('msdoc-page-guide-corner-top-left', metrics.marginLeftPx - 18, metrics.marginTopPx - 18));
    guides.appendChild(createCornerGuide('msdoc-page-guide-corner-top-right', metrics.widthPx - metrics.marginRightPx, metrics.marginTopPx - 18));
    guides.appendChild(createCornerGuide('msdoc-page-guide-corner-bottom-left', metrics.marginLeftPx - 18, metrics.heightPx - metrics.marginBottomPx));
    guides.appendChild(createCornerGuide('msdoc-page-guide-corner-bottom-right', metrics.widthPx - metrics.marginRightPx, metrics.heightPx - metrics.marginBottomPx));
    pageEl.appendChild(guides);

    const headerBand = document.createElement('div');
    headerBand.className = 'msdoc-page-header-band';
    headerBand.style.left = `${metrics.marginLeftPx}px`;
    headerBand.style.width = `${metrics.bodyWidthPx}px`;
    headerBand.style.paddingTop = `${Math.max(metrics.headerTopPx, 0)}px`;
    const headerBandInner = document.createElement('div');
    headerBandInner.className = 'msdoc-page-header-band-inner';
    const headerStory = pickStoryTemplate(chromeTemplates, 'header', page, sections);
    if (headerStory) {
      appendHtml(headerBandInner, renderChromeContent(headerStory, page));
      if (headerStory.pageNumberLike) applyDynamicPageNumber(headerBandInner, page, pages.length);
    }
    headerBand.appendChild(headerBandInner);
    pageEl.appendChild(headerBand);

    const footerBand = document.createElement('div');
    footerBand.className = 'msdoc-page-footer-band';
    footerBand.style.left = `${metrics.marginLeftPx}px`;
    footerBand.style.width = `${metrics.bodyWidthPx}px`;
    footerBand.style.paddingBottom = `${Math.max(metrics.footerBottomPx, 0)}px`;
    footerBand.style.display = 'flex';
    footerBand.style.flexDirection = 'column';
    footerBand.style.justifyContent = 'flex-end';
    const footerBandInner = document.createElement('div');
    footerBandInner.className = 'msdoc-page-footer-band-inner';
    const footerStory = pickStoryTemplate(chromeTemplates, 'footer', page, sections);
    if (footerStory) {
      appendHtml(footerBandInner, renderChromeContent(footerStory, page));
      if (footerStory.pageNumberLike) applyDynamicPageNumber(footerBandInner, page, pages.length);
    }
    footerBand.appendChild(footerBandInner);
    pageEl.appendChild(footerBand);

    const body = document.createElement('div');
    body.className = 'msdoc-page-body';
    applyPageBodyStyle(body, section);
    const bodyContent = document.createElement('div');
    bodyContent.className = 'msdoc-page-body-content';
    if (!page.blank) {
      for (const block of page.blocks) bodyContent.appendChild(block.cloneNode(true));
    }
    body.appendChild(bodyContent);
    pageEl.appendChild(body);

    const overlay = document.createElement('div');
    overlay.className = 'msdoc-page-overlay';

    if (!page.blank) {
      const pageMainShapes = mainShapes.filter((shape) => findPageIndexForShape(pages, shape) === pageIndex);
      for (const shape of pageMainShapes) {
        const textbox = (shape.matchedTextboxId ? mainTextboxById.get(shape.matchedTextboxId) : undefined) || mainTextboxByShapeId.get(shape.shapeId);
        const asset = shape.imageAssetId ? imageAssets.get(shape.imageAssetId) : undefined;
        const itemEl = renderMainShapeOverlay(shape, textbox, asset, section, body, bodyContent);
        if (itemEl) overlay.appendChild(itemEl);
      }
    }

    for (const kind of ['header', 'footer'] as const) {
      for (const template of pickTextboxTemplates(chromeTemplates, kind, page)) {
        if (!template.item) continue;
        if (!template.item.shape) {
          const target = kind === 'header' ? headerBandInner : footerBandInner;
          appendHtml(target, renderChromeContent(template, page));
          if (template.pageNumberLike) applyDynamicPageNumber(target, page, pages.length);
          finalizeChromeBand(kind === 'header' ? headerBand : footerBand, kind, metrics);
          continue;
        }
        const position = shapeToPagePosition(template.item.shape, section);
        const itemEl = document.createElement('div');
        itemEl.className = `msdoc-page-overlay-item msdoc-page-overlay-item-${kind}${template.pageNumberLike ? ' msdoc-page-overlay-item-page-number' : ''}`;
        itemEl.style.left = `${position.left}px`;
        itemEl.style.top = `${position.top}px`;
        itemEl.style.width = `${position.width}px`;
        itemEl.style.minHeight = `${position.height}px`;
        if (template.item.shape.behindText) itemEl.style.opacity = '0.92';
        const asset = template.item.shape.imageAssetId ? imageAssets.get(template.item.shape.imageAssetId) : undefined;
        if (asset) {
          const assetEl = createShapeAssetLayer(template.item.shape, asset);
          if (assetEl) itemEl.appendChild(assetEl);
        }
        const contentEl = document.createElement('div');
        contentEl.className = 'msdoc-page-overlay-item-content';
        appendHtml(contentEl, renderChromeContent(template, page));
        if (template.pageNumberLike) applyDynamicPageNumber(contentEl, page, pages.length);
        itemEl.appendChild(contentEl);
        overlay.appendChild(itemEl);
      }

      for (const shape of pickHeaderShapesForPage(headerShapes, kind, page, sections)) {
        const asset = shape.imageAssetId ? imageAssets.get(shape.imageAssetId) : undefined;
        if (!asset) continue;
        const itemEl = renderShapeAssetOverlay(shape, asset, section);
        if (itemEl) overlay.appendChild(itemEl);
      }
    }
    pageEl.appendChild(overlay);
    pagedView.appendChild(pageEl);
    finalizeChromeBand(headerBand, 'header', metrics);
    finalizeChromeBand(footerBand, 'footer', metrics);
  }
}

function installViewerRuntime(container: HTMLElement, rendered: MsDocRenderResult): () => void {
  const root = container.querySelector('.msdoc-root');
  if (!(root instanceof HTMLElement)) return () => undefined;
  const toolbar = container.querySelector('.msdoc-viewer-toolbar');
  const flowButton = toolbar?.querySelector('[data-view="flow"]');
  const pagedButton = toolbar?.querySelector('[data-view="paged"]');

  const setView = (view: 'flow' | 'paged'): void => {
    root.dataset.msdocView = view;
    if (flowButton instanceof HTMLButtonElement) flowButton.setAttribute('aria-pressed', String(view === 'flow'));
    if (pagedButton instanceof HTMLButtonElement) pagedButton.setAttribute('aria-pressed', String(view === 'paged'));
  };

  const renderPages = (): void => renderPagedView(root, rendered);
  renderPages();
  setView('paged');

  const onClick = (event: Event): void => {
    const target = event.target as HTMLElement | null;
    const view = target?.getAttribute?.('data-view');
    if (view === 'flow' || view === 'paged') setView(view);
  };
  toolbar?.addEventListener('click', onClick);

  let frame = 0;
  const onResize = (): void => {
    if (frame) cancelAnimationFrame(frame);
    frame = requestAnimationFrame(() => {
      frame = 0;
      renderPages();
    });
  };
  window.addEventListener('resize', onResize);

  return () => {
    if (frame) cancelAnimationFrame(frame);
    toolbar?.removeEventListener('click', onClick);
    window.removeEventListener('resize', onResize);
  };
}

export function mountMsDoc(container: HTMLElement, rendered: MsDocRenderResult): HTMLElement {
  if (!container) throw new Error('A container element is required');
  container.innerHTML = `
    <style data-msdoc>${rendered.css}</style>
    <style data-msdoc-viewer>${runtimeViewerCss()}</style>
    <div class="msdoc-viewer-shell">
      <div class="msdoc-viewer-toolbar" role="toolbar" aria-label="Document view">
        <span class="msdoc-viewer-title">视图</span>
        <button type="button" class="msdoc-viewer-button" data-view="paged" aria-pressed="true">分页</button>
        <button type="button" class="msdoc-viewer-button" data-view="flow" aria-pressed="false">文档流</button>
        <span class="msdoc-viewer-hint">分页视图会按节属性应用纸张尺寸、页边距、页眉页脚位置与分页效果。</span>
      </div>
      <div class="msdoc-root" data-msdoc-view="paged">
        <div class="msdoc-flow-view">${rendered.html}</div>
        <div class="msdoc-paged-view"></div>
      </div>
    </div>`;
  const previousCleanup = (container as HTMLElement & { __msdocCleanup?: () => void }).__msdocCleanup;
  previousCleanup?.();
  (container as HTMLElement & { __msdocCleanup?: () => void }).__msdocCleanup = installViewerRuntime(container, rendered);
  return container;
}

export async function parseMsDocToHtml(input: ViewerInput, options: MsDocParseToHtmlOptions = {}): Promise<MsDocRenderResult> {
  const buffer = await normalizeInput(input);
  if (options.workerClient) {
    return options.workerClient.parseToHtml(buffer, {
      parseOptions: options.parseOptions || {},
      renderOptions: options.renderOptions || {},
    });
  }
  const parsed = parseMsDoc(buffer, options.parseOptions || {});
  return renderMsDoc(parsed, options.renderOptions || {});
}

/**
 * Small DOM-oriented helper that keeps browser integration trivial.
 * Apps can either use it directly or consume the lower-level parse/render APIs.
 */
export function createMsDocViewer(container: HTMLElement, config: MsDocViewerConfig = {}): MsDocViewer {
  let current: MsDocRenderResult | null = null;
  return {
    async load(input: ViewerInput, options: MsDocViewerLoadOptions = {}): Promise<MsDocRenderResult> {
      const rendered = await parseMsDocToHtml(input, {
        workerClient: options.workerClient || config.workerClient,
        parseOptions: { ...(config.parseOptions || {}), ...(options.parseOptions || {}) },
        renderOptions: { ...(config.renderOptions || {}), ...(options.renderOptions || {}) },
      });
      mountMsDoc(container, rendered);
      current = rendered;
      return rendered;
    },
    mount(rendered: MsDocRenderResult): HTMLElement {
      current = rendered;
      return mountMsDoc(container, rendered);
    },
    clear(): void {
      (container as HTMLElement & { __msdocCleanup?: () => void }).__msdocCleanup?.();
      delete (container as HTMLElement & { __msdocCleanup?: () => void }).__msdocCleanup;
      container.innerHTML = '';
      current = null;
    },
    destroy(): void {
      this.clear();
    },
    get value(): MsDocRenderResult | null {
      return current;
    },
  };
}
