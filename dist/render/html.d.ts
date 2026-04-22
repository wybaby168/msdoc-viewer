import type { MsDocParseResult, MsDocRenderOptions, MsDocRenderResult, ParagraphBlock, TableBlock } from '../types.js';
export declare function renderBlockList(blocks: Array<ParagraphBlock | TableBlock>): string;
export declare function defaultMsDocCss(): string;
/**
 * Converts the parsed AST into HTML and a companion CSS string.
 * Keeping rendering separate from parsing makes it easier for downstream apps
 * to customize styles or consume the AST directly.
 */
export declare function renderMsDoc(parsed: MsDocParseResult, options?: MsDocRenderOptions): MsDocRenderResult;
