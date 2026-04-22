import type { FibRgFcLcb, ShapeAnchorInfo, ShapeAnchorStoryKind } from '../types.js';
/**
 * Parses PlcfSpaMom / PlcfSpaHdr records and exposes the floating-shape anchor
 * metadata described by the Spa structure. This gives the renderer enough
 * information to surface textbox dimensions, wrap modes, and unmatched drawing
 * anchors even when the full OfficeArt shape geometry is not yet rendered.
 */
export declare function readShapeAnchors(tableBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, storyCpBase: number, story: ShapeAnchorStoryKind): ShapeAnchorInfo[];
