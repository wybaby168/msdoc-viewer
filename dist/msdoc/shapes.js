import { BinaryReader } from '../core/binary.js';
import { uniqueId } from '../core/utils.js';
import { readFixedPlc } from './stories.js';
const SPA_DATA_SIZE = 26;
function decodeAnchorX(code) {
    switch (code) {
        case 0: return 'margin';
        case 1: return 'page';
        case 2: return 'column';
        default: return `unknown-${code}`;
    }
}
function decodeAnchorY(code) {
    switch (code) {
        case 0: return 'margin';
        case 1: return 'page';
        case 2: return 'paragraph';
        default: return `unknown-${code}`;
    }
}
function decodeWrapStyle(code) {
    switch (code) {
        case 0: return 'around';
        case 1: return 'topBottom';
        case 2: return 'square';
        case 3: return 'none';
        case 4: return 'tight';
        case 5: return 'through';
        default: return `unknown-${code}`;
    }
}
function decodeWrapSide(code) {
    switch (code) {
        case 0: return 'both';
        case 1: return 'left';
        case 2: return 'right';
        case 3: return 'largest';
        default: return `unknown-${code}`;
    }
}
function parseSpaEntry(data, storyCpBase, anchorCp, story) {
    if (data.length < SPA_DATA_SIZE)
        return null;
    const reader = new BinaryReader(data);
    const left = reader.i32(4);
    const top = reader.i32(8);
    const right = reader.i32(12);
    const bottom = reader.i32(16);
    const flags = reader.u16(20);
    const shapeId = reader.u32(0);
    if (!shapeId)
        return null;
    return {
        id: uniqueId(story === 'header' ? 'shape-hdr' : 'shape'),
        story,
        anchorCp: storyCpBase + anchorCp,
        shapeId,
        boundsTwips: {
            left,
            top,
            right,
            bottom,
            width: Math.max(0, right - left),
            height: Math.max(0, bottom - top),
        },
        anchorX: decodeAnchorX((flags >> 1) & 0x3),
        anchorY: decodeAnchorY((flags >> 3) & 0x3),
        wrapStyle: decodeWrapStyle((flags >> 5) & 0xf),
        wrapSide: decodeWrapSide((flags >> 9) & 0xf),
        behindText: Boolean((flags >> 14) & 0x1),
        anchorLocked: Boolean((flags >> 15) & 0x1),
    };
}
/**
 * Parses PlcfSpaMom / PlcfSpaHdr records and exposes the floating-shape anchor
 * metadata described by the Spa structure. This gives the renderer enough
 * information to surface textbox dimensions, wrap modes, and unmatched drawing
 * anchors even when the full OfficeArt shape geometry is not yet rendered.
 */
export function readShapeAnchors(tableBytes, fibRgFcLcb, storyCpBase, story) {
    const fc = story === 'header'
        ? fibRgFcLcb.fcPlcSpaHdr
        : fibRgFcLcb.fcPlcSpaMom;
    const lcb = story === 'header'
        ? fibRgFcLcb.lcbPlcSpaHdr
        : fibRgFcLcb.lcbPlcSpaMom;
    return readFixedPlc(tableBytes, fc, lcb, SPA_DATA_SIZE)
        .map((entry) => parseSpaEntry(entry.data, storyCpBase, entry.cpStart, story))
        .filter((entry) => Boolean(entry));
}
//# sourceMappingURL=shapes.js.map