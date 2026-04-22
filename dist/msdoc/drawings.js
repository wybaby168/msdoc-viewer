import { BinaryReader } from '../core/binary.js';
import { dataUrlFromBytes, uniqueId } from '../core/utils.js';
import { convertMetafileToSvg } from './vector.js';
import { parseOfficeArtRecordHeader } from './objects.js';
const OFFICEART_CONTAINER = 0x0f;
const OFFICEART_BSTORE_CONTAINER = 0xf001;
const OFFICEART_FBSE = 0xf007;
const OFFICEART_SP_CONTAINER = 0xf004;
const OFFICEART_FSP = 0xf00a;
const OFFICEART_FOPT = 0xf00b;
const OFFICEART_SECONDARY_FOPT = 0xf121;
const OFFICEART_TERTIARY_FOPT = 0xf122;
const OFFICEART_BLIP_EMF = 0xf01a;
const OFFICEART_BLIP_WMF = 0xf01b;
const OFFICEART_BLIP_PICT = 0xf01c;
const OFFICEART_BLIP_JPEG = 0xf01d;
const OFFICEART_BLIP_PNG = 0xf01e;
const OFFICEART_BLIP_DIB = 0xf01f;
const OFFICEART_BLIP_TIFF = 0xf029;
const OFFICEART_BLIP_JPEG_ALT = 0xf02a;
const PROP_PIB = 0x0104;
const PROP_FILL_BLIP = 0x0186;
const PROP_SHAPE_NAME = 0x0380;
const PROP_ALT_TEXT = 0x0381;
function u16(bytes, offset) {
    if (offset + 2 > bytes.length)
        return 0;
    return (bytes[offset] ?? 0) | ((bytes[offset + 1] ?? 0) << 8);
}
function u32(bytes, offset) {
    if (offset + 4 > bytes.length)
        return 0;
    return (((bytes[offset] ?? 0))
        | ((bytes[offset + 1] ?? 0) << 8)
        | ((bytes[offset + 2] ?? 0) << 16)
        | (((bytes[offset + 3] ?? 0) << 24) >>> 0)) >>> 0;
}
function decodeUtf16(bytes) {
    if (!bytes?.length)
        return undefined;
    const trimmed = bytes.length >= 2 && bytes[bytes.length - 1] === 0 && bytes[bytes.length - 2] === 0
        ? bytes.subarray(0, bytes.length - 2)
        : bytes;
    const value = new TextDecoder('utf-16le').decode(trimmed).replace(/\0+$/g, '').trim();
    return value || undefined;
}
function decodeAnsi(bytes) {
    if (!bytes?.length)
        return undefined;
    const value = new TextDecoder('windows-1252').decode(bytes).replace(/\0+$/g, '').trim();
    return value || undefined;
}
function isBrowserDisplayableMime(mime) {
    return /^(?:image\/png|image\/jpeg|image\/gif|image\/bmp|image\/webp|image\/svg\+xml|image\/tiff)$/i.test(mime);
}
function createBitmapFileHeader(pixelOffset, totalSize) {
    const header = new Uint8Array(14);
    const view = new DataView(header.buffer);
    view.setUint8(0, 0x42);
    view.setUint8(1, 0x4d);
    view.setUint32(2, totalSize, true);
    view.setUint16(6, 0, true);
    view.setUint16(8, 0, true);
    view.setUint32(10, pixelOffset, true);
    return header;
}
function dibToBmp(dibBytes) {
    if (dibBytes.length < 12)
        return null;
    const reader = new BinaryReader(dibBytes);
    const headerSize = reader.u32(0);
    if (headerSize < 12 || headerSize > dibBytes.length)
        return null;
    let bitsPerPixel = 0;
    let compression = 0;
    let colorsUsed = 0;
    let paletteEntrySize = 4;
    if (headerSize === 12) {
        bitsPerPixel = reader.u16(10);
        paletteEntrySize = 3;
    }
    else {
        bitsPerPixel = reader.u16(14);
        compression = reader.u32(16);
        colorsUsed = reader.u32(32);
    }
    const colorCount = colorsUsed || (bitsPerPixel > 0 && bitsPerPixel <= 8 ? 1 << bitsPerPixel : 0);
    let paletteSize = colorCount * paletteEntrySize;
    if (compression === 3 && headerSize >= 40)
        paletteSize += 12;
    const pixelOffset = 14 + headerSize + paletteSize;
    if (pixelOffset > 14 + dibBytes.length)
        return null;
    const bmpHeader = createBitmapFileHeader(pixelOffset, 14 + dibBytes.length);
    const out = new Uint8Array(14 + dibBytes.length);
    out.set(bmpHeader, 0);
    out.set(dibBytes, 14);
    return out;
}
function detectMimeFromBlipType(blipType) {
    switch (blipType) {
        case 0x02: return 'image/emf';
        case 0x03: return 'image/wmf';
        case 0x04: return 'image/pict';
        case 0x05: return 'image/jpeg';
        case 0x06: return 'image/png';
        case 0x07: return 'image/dib';
        case 0x11: return 'image/tiff';
        case 0x12: return 'image/jpeg';
        default: return null;
    }
}
function extractBlipPayload(bytes, offset, header, extraMeta = {}) {
    const payloadOffset = offset + 8;
    const uidCount = (header.recInstance & 0x1) === 1 ? 2 : 1;
    const uidBytes = uidCount * 16;
    const atom = bytes.subarray(payloadOffset, payloadOffset + header.recLen);
    if (header.recType === OFFICEART_BLIP_PNG
        || header.recType === OFFICEART_BLIP_JPEG
        || header.recType === OFFICEART_BLIP_JPEG_ALT
        || header.recType === OFFICEART_BLIP_DIB
        || header.recType === OFFICEART_BLIP_TIFF) {
        const infoSize = uidBytes + 1;
        if (atom.length < infoSize)
            return null;
        let mime = 'application/octet-stream';
        let payload = atom.subarray(infoSize);
        if (header.recType === OFFICEART_BLIP_PNG)
            mime = 'image/png';
        else if (header.recType === OFFICEART_BLIP_JPEG || header.recType === OFFICEART_BLIP_JPEG_ALT)
            mime = 'image/jpeg';
        else if (header.recType === OFFICEART_BLIP_TIFF)
            mime = 'image/tiff';
        else if (header.recType === OFFICEART_BLIP_DIB) {
            const bmp = dibToBmp(payload);
            if (bmp) {
                payload = bmp;
                mime = 'image/bmp';
            }
            else {
                mime = 'image/dib';
            }
        }
        return {
            mime,
            bytes: payload,
            displayable: isBrowserDisplayableMime(mime),
            meta: {
                recType: header.recType,
                recInstance: header.recInstance,
                ...extraMeta,
            },
        };
    }
    if (header.recType === OFFICEART_BLIP_EMF
        || header.recType === OFFICEART_BLIP_WMF
        || header.recType === OFFICEART_BLIP_PICT) {
        const infoSize = uidBytes + 34;
        if (atom.length < infoSize)
            return null;
        const metafileCompression = atom[uidBytes + 32] ?? 0xff;
        const metafileFilter = atom[uidBytes + 33] ?? 0xff;
        const payload = atom.subarray(infoSize);
        const mime = header.recType === OFFICEART_BLIP_EMF
            ? 'image/emf'
            : header.recType === OFFICEART_BLIP_WMF
                ? 'image/wmf'
                : 'image/pict';
        return {
            mime,
            bytes: payload,
            displayable: isBrowserDisplayableMime(mime),
            meta: {
                recType: header.recType,
                recInstance: header.recInstance,
                metafileCompression,
                metafileFilter,
                metafileCompressed: metafileCompression !== 0xfe,
                ...extraMeta,
            },
        };
    }
    return null;
}
function maybeConvertVectorCandidate(candidate) {
    if (!/^(?:image\/emf|image\/wmf)$/i.test(candidate.mime))
        return candidate;
    if (!candidate.bytes.length)
        return candidate;
    if (candidate.meta?.metafileCompressed)
        return candidate;
    const converted = convertMetafileToSvg(candidate.mime, candidate.bytes);
    if (!converted)
        return candidate;
    return {
        mime: converted.mime,
        bytes: converted.bytes,
        displayable: true,
        meta: {
            ...(candidate.meta || {}),
            vectorConverted: true,
            vectorSourceMime: converted.sourceMime,
            vectorWidth: converted.width,
            vectorHeight: converted.height,
            vectorRecordCount: converted.recordCount,
        },
    };
}
function createImageAsset(candidate, fbseIndex, fbseOffset, fbse, sourceStream) {
    const meta = {
        pictureOffset: -1,
        lcb: candidate.bytes.length,
        cbHeader: 0,
        sourceKind: 'embedded',
        browserRenderable: candidate.displayable,
        fbseBlipType: fbse.btWin32 || fbse.btMacOS,
        fbseMime: detectMimeFromBlipType(fbse.btWin32 || fbse.btMacOS),
        fbseTag: fbse.tag,
        fbseSize: fbse.size,
        fbseRefCount: fbse.refCount,
        fbseDelayOffset: fbse.delayOffset,
        fbseDelayStream: sourceStream,
        fbseName: fbse.name,
        fbseIndex,
        dggOffset: fbseOffset,
        ...(candidate.meta || {}),
    };
    return {
        id: uniqueId(`asset-dgg-img-${fbseIndex}`),
        type: 'image',
        mime: candidate.mime,
        bytes: candidate.bytes,
        dataUrl: dataUrlFromBytes(candidate.bytes, candidate.mime),
        displayable: candidate.displayable,
        meta,
    };
}
function parseFbseInfo(bytes, offset, header) {
    if (header.recType !== OFFICEART_FBSE || header.recLen < 36 || offset + header.size > bytes.length)
        return null;
    const payloadOffset = offset + 8;
    const reader = new BinaryReader(bytes);
    const btWin32 = reader.u8(payloadOffset);
    const btMacOS = reader.u8(payloadOffset + 1);
    const tag = reader.u16(payloadOffset + 18);
    const size = reader.u32(payloadOffset + 20);
    const refCount = reader.u32(payloadOffset + 24);
    const delayOffset = reader.u32(payloadOffset + 28);
    const nameLength = reader.u8(payloadOffset + 33);
    const nameOffset = payloadOffset + 36;
    const nameBytes = nameLength > 0 && nameOffset + nameLength <= offset + header.size
        ? bytes.subarray(nameOffset, nameOffset + nameLength)
        : undefined;
    return {
        btWin32,
        btMacOS,
        tag,
        size,
        refCount,
        delayOffset,
        nameLength,
        name: decodeAnsi(nameBytes),
    };
}
function extractEmbeddedBlipFromFbse(bytes, offset, header, fbseIndex) {
    const fbse = parseFbseInfo(bytes, offset, header);
    if (!fbse)
        return null;
    const payloadOffset = offset + 8;
    const embeddedBlipOffset = payloadOffset + 36 + fbse.nameLength;
    if (embeddedBlipOffset + 8 > offset + header.size)
        return null;
    const embeddedHeader = parseOfficeArtRecordHeader(bytes, embeddedBlipOffset);
    if (!embeddedHeader)
        return null;
    const candidate = extractBlipPayload(bytes, embeddedBlipOffset, embeddedHeader, {
        fbseBlipType: fbse.btWin32 || fbse.btMacOS,
        fbseMime: detectMimeFromBlipType(fbse.btWin32 || fbse.btMacOS),
        fbseTag: fbse.tag,
        fbseSize: fbse.size,
        fbseRefCount: fbse.refCount,
        fbseDelayOffset: fbse.delayOffset,
        fbseName: fbse.name,
    });
    if (!candidate)
        return null;
    return createImageAsset(maybeConvertVectorCandidate(candidate), fbseIndex, offset, fbse);
}
function extractDelayedBlipFromFbse(bytes, offset, header, fbseIndex, streams) {
    const fbse = parseFbseInfo(bytes, offset, header);
    if (!fbse || !fbse.delayOffset)
        return null;
    for (const stream of streams) {
        if (!stream.bytes.length || fbse.delayOffset + 8 > stream.bytes.length)
            continue;
        const blipHeader = parseOfficeArtRecordHeader(stream.bytes, fbse.delayOffset);
        if (!blipHeader)
            continue;
        const candidate = extractBlipPayload(stream.bytes, fbse.delayOffset, blipHeader, {
            fbseBlipType: fbse.btWin32 || fbse.btMacOS,
            fbseMime: detectMimeFromBlipType(fbse.btWin32 || fbse.btMacOS),
            fbseTag: fbse.tag,
            fbseSize: fbse.size,
            fbseRefCount: fbse.refCount,
            fbseDelayOffset: fbse.delayOffset,
            fbseDelayStream: stream.name,
            fbseName: fbse.name,
        });
        if (!candidate)
            continue;
        return createImageAsset(maybeConvertVectorCandidate(candidate), fbseIndex, offset, fbse, stream.name);
    }
    return null;
}
function parseFoptProperties(bytes, offset, header) {
    const properties = [];
    const count = header.recInstance;
    const tableStart = offset + 8;
    const tableEnd = offset + header.size;
    let cursor = tableStart;
    for (let index = 0; index < count && cursor + 6 <= tableEnd; index += 1) {
        const opid = u16(bytes, cursor);
        const value = u32(bytes, cursor + 2);
        cursor += 6;
        properties.push({
            pid: opid & 0x3fff,
            fBid: Boolean(opid & 0x4000),
            fComplex: Boolean(opid & 0x8000),
            value,
        });
    }
    for (const property of properties) {
        if (!property.fComplex || property.value <= 0)
            continue;
        const next = Math.min(tableEnd, cursor + property.value);
        property.complexData = bytes.subarray(cursor, next);
        cursor = next;
    }
    return properties;
}
/**
 * DggInfo can contain host-specific non-OfficeArt bytes between drawing records.
 * Instead of assuming a single perfectly aligned tree, this scanner advances one
 * byte at a time until it can resynchronize on a valid OfficeArtRecordHeader.
 */
function scanOfficeArtRecords(bytes, start, end, visitor) {
    let cursor = Math.max(0, start);
    const limit = Math.min(end, bytes.length);
    while (cursor + 8 <= limit) {
        const header = parseOfficeArtRecordHeader(bytes, cursor);
        if (!header || cursor + header.size > limit) {
            cursor += 1;
            continue;
        }
        visitor(cursor, header);
        if (header.recVer === OFFICEART_CONTAINER) {
            scanOfficeArtRecords(bytes, cursor + 8, cursor + header.size, visitor);
        }
        cursor += header.size;
    }
}
function blipKindForPropertyId(id) {
    return id === PROP_PIB ? 'pib' : 'fillBlip';
}
function parseShapeContainer(bytes, offset, header) {
    let shapeId = 0;
    let flags = 0;
    let shapeTypeCode;
    let name;
    let description;
    const bidRefs = [];
    scanOfficeArtRecords(bytes, offset + 8, offset + header.size, (recordOffset, recordHeader) => {
        if (recordHeader.recType === OFFICEART_FSP) {
            shapeId = u32(bytes, recordOffset + 8);
            flags = u32(bytes, recordOffset + 12);
            shapeTypeCode = recordHeader.recInstance;
            return;
        }
        if (recordHeader.recType !== OFFICEART_FOPT && recordHeader.recType !== OFFICEART_SECONDARY_FOPT && recordHeader.recType !== OFFICEART_TERTIARY_FOPT) {
            return;
        }
        for (const property of parseFoptProperties(bytes, recordOffset, recordHeader)) {
            if (property.fBid && (property.pid === PROP_PIB || property.pid === PROP_FILL_BLIP) && property.value > 0) {
                bidRefs.push({ kind: blipKindForPropertyId(property.pid), index: property.value });
            }
            else if (property.pid === PROP_SHAPE_NAME) {
                name = decodeUtf16(property.complexData) || name;
            }
            else if (property.pid === PROP_ALT_TEXT) {
                description = decodeUtf16(property.complexData) || description;
            }
        }
    });
    if (!shapeId)
        return null;
    return {
        shapeId,
        flags,
        shapeTypeCode,
        bidRefs,
        name,
        description,
    };
}
function parseDggBytes(bytes, delayStreams) {
    const context = {
        assets: [],
        assetByBlipIndex: new Map(),
        shapes: new Map(),
        warnings: [],
        fbseIndex: 0,
        delayStreams,
    };
    scanOfficeArtRecords(bytes, 0, bytes.length, (offset, header) => {
        if (header.recType === OFFICEART_FBSE) {
            context.fbseIndex += 1;
            let asset = extractEmbeddedBlipFromFbse(bytes, offset, header, context.fbseIndex);
            if (!asset)
                asset = extractDelayedBlipFromFbse(bytes, offset, header, context.fbseIndex, context.delayStreams);
            if (asset) {
                context.assetByBlipIndex.set(context.fbseIndex, asset);
                context.assets.push(asset);
            }
            return;
        }
        if (header.recType === OFFICEART_BSTORE_CONTAINER || header.recType !== OFFICEART_SP_CONTAINER)
            return;
        const shape = parseShapeContainer(bytes, offset, header);
        if (shape)
            context.shapes.set(shape.shapeId, shape);
    });
    const resolvedShapes = new Map();
    for (const shape of context.shapes.values()) {
        let blipRef;
        let imageAsset;
        for (const ref of shape.bidRefs) {
            const asset = context.assetByBlipIndex.get(ref.index);
            if (asset) {
                blipRef = ref;
                imageAsset = asset;
                break;
            }
        }
        if (!imageAsset && shape.bidRefs.length) {
            context.warnings.push({
                code: 'officeart-shape-missing-blip',
                severity: 'info',
                message: 'OfficeArt shape references a BLIP that could not be resolved from DggInfo',
                details: { shapeId: shape.shapeId, bidRefs: shape.bidRefs },
            });
        }
        resolvedShapes.set(shape.shapeId, {
            shapeId: shape.shapeId,
            flags: shape.flags,
            name: shape.name,
            description: shape.description,
            shapeTypeCode: shape.shapeTypeCode,
            blipRef,
            imageAsset,
            imageAssetId: imageAsset?.id,
        });
    }
    return {
        assets: context.assets,
        shapes: resolvedShapes,
        warnings: context.warnings,
    };
}
/**
 * Extracts reusable DggInfo metadata: the BLIP store and picture-shape bindings.
 * The BLIP payload can be embedded directly in the FBSE record or stored in a
 * delay stream (commonly WordDocument for legacy Word hosts), so both paths are
 * resolved here before rendering sees the shape anchors.
 */
export function parseDrawingGroup(tableBytes, wordBytes, fibRgFcLcb, dataBytes = new Uint8Array(0)) {
    const fc = fibRgFcLcb.fcDggInfo;
    const lcb = fibRgFcLcb.lcbDggInfo;
    if (fc == null || lcb == null || lcb <= 0 || fc < 0 || fc + lcb > tableBytes.length) {
        return { assets: [], shapes: new Map(), warnings: [] };
    }
    const dggInfoBytes = tableBytes.subarray(fc, fc + lcb);
    return parseDggBytes(dggInfoBytes, [
        { name: 'WordDocument', bytes: wordBytes },
        { name: 'Data', bytes: dataBytes },
        { name: 'Table', bytes: tableBytes },
    ]);
}
function headerKindFromRole(role) {
    if (!role)
        return undefined;
    if (role.endsWith('Header'))
        return 'header';
    if (role.endsWith('Footer'))
        return 'footer';
    return undefined;
}
/**
 * Header/footer floating shapes are anchored in the shared header story range.
 * We map them back to the concrete odd/even/first story interval that actually
 * owns the anchor CP so pagination can place them in the correct page band.
 */
export function resolveHeaderAnchorBinding(roleWindows, anchorCp) {
    const nonEmpty = roleWindows.find((entry) => anchorCp >= entry.cpStart && anchorCp < entry.cpEnd);
    if (nonEmpty) {
        return {
            sectionIndex: entryToZeroBasedSection(nonEmpty.sectionIndex),
            role: nonEmpty.role,
            kind: headerKindFromRole(nonEmpty.role),
        };
    }
    const exact = roleWindows.find((entry) => entry.cpStart === anchorCp || entry.cpEnd === anchorCp);
    if (exact) {
        return {
            sectionIndex: entryToZeroBasedSection(exact.sectionIndex),
            role: exact.role,
            kind: headerKindFromRole(exact.role),
        };
    }
    const fallback = [...roleWindows]
        .filter((entry) => entry.cpStart <= anchorCp)
        .sort((left, right) => right.cpStart - left.cpStart)[0];
    if (!fallback)
        return {};
    return {
        sectionIndex: entryToZeroBasedSection(fallback.sectionIndex),
        role: fallback.role,
        kind: headerKindFromRole(fallback.role),
    };
}
function entryToZeroBasedSection(sectionIndex) {
    if (sectionIndex == null || sectionIndex <= 0)
        return undefined;
    return sectionIndex - 1;
}
//# sourceMappingURL=drawings.js.map