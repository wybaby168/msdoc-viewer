import type { FibRgFcLcb, HeaderFooterRole, ImageAsset, MsDocWarning, ShapeBlipReference } from '../types.js';
export interface DrawingShapeInfo {
    shapeId: number;
    flags: number;
    name?: string;
    description?: string;
    shapeTypeCode?: number;
    blipRef?: ShapeBlipReference;
    imageAssetId?: string;
    imageAsset?: ImageAsset;
}
export interface DrawingGroupInfo {
    assets: ImageAsset[];
    shapes: Map<number, DrawingShapeInfo>;
    warnings: MsDocWarning[];
}
/**
 * Extracts reusable DggInfo metadata: the BLIP store and picture-shape bindings.
 * The BLIP payload can be embedded directly in the FBSE record or stored in a
 * delay stream (commonly WordDocument for legacy Word hosts), so both paths are
 * resolved here before rendering sees the shape anchors.
 */
export declare function parseDrawingGroup(tableBytes: Uint8Array, wordBytes: Uint8Array, fibRgFcLcb: FibRgFcLcb, dataBytes?: Uint8Array): DrawingGroupInfo;
export interface HeaderAnchorBinding {
    sectionIndex?: number;
    role?: HeaderFooterRole;
    kind?: 'header' | 'footer';
}
/**
 * Header/footer floating shapes are anchored in the shared header story range.
 * We map them back to the concrete odd/even/first story interval that actually
 * owns the anchor CP so pagination can place them in the correct page band.
 */
export declare function resolveHeaderAnchorBinding(roleWindows: Array<{
    cpStart: number;
    cpEnd: number;
    sectionIndex?: number;
    role: HeaderFooterRole;
}>, anchorCp: number): HeaderAnchorBinding;
