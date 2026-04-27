import type { ListCollectionSummary, ListLevelDefinition, ParagraphModel, FibRgFcLcb, CharState, DecodedProperty } from '../types.js';
export declare function parseLists(tableBytes: Uint8Array, fib: FibRgFcLcb): ListCollectionSummary;
export interface ApplyListFormattingOptions {
    /**
     * Resolves LVL character properties into a renderable style for the generated
     * list label. The caller owns font/style inheritance because that context
     * lives in the parser, not in the raw list table reader.
     */
    resolveLabelStyle?: (charProps: DecodedProperty[], paragraph: ParagraphModel, level: ListLevelDefinition) => CharState | undefined;
}
/**
 * Applies list labels to paragraph models using sprmPIlfo/sprmPIlvl and the
 * PlfLst/PlfLfo model. This intentionally computes display labels in the AST so
 * renderers do not need to emulate Word's numbering counters themselves.
 */
export declare function applyListFormatting(paragraphs: ParagraphModel[], lists: ListCollectionSummary, options?: ApplyListFormattingOptions): void;
