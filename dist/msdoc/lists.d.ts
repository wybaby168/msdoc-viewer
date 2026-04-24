import type { ListCollectionSummary, ParagraphModel, FibRgFcLcb } from '../types.js';
export declare function parseLists(tableBytes: Uint8Array, fib: FibRgFcLcb): ListCollectionSummary;
/**
 * Applies list labels to paragraph models using sprmPIlfo/sprmPIlvl and the
 * PlfLst/PlfLfo model. This intentionally computes display labels in the AST so
 * renderers do not need to emulate Word's numbering counters themselves.
 */
export declare function applyListFormatting(paragraphs: ParagraphModel[], lists: ListCollectionSummary): void;
