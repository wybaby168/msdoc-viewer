import { BinaryReader } from '../core/binary.js';
export function buildStoryWindows(fibRgLw, totalTextLength) {
    let cursor = 0;
    const make = (kind, length) => {
        const normalizedLength = Math.max(0, length || 0);
        const cpStart = cursor;
        const cpEnd = Math.min(totalTextLength, cpStart + normalizedLength);
        cursor += normalizedLength;
        return { kind, cpStart, cpEnd, length: normalizedLength };
    };
    return {
        main: make('main', fibRgLw.ccpText),
        footnote: make('footnote', fibRgLw.ccpFtn),
        header: make('header', fibRgLw.ccpHdd),
        macro: make('macro', fibRgLw.ccpMcr),
        comment: make('comment', fibRgLw.ccpAtn),
        endnote: make('endnote', fibRgLw.ccpEdn),
        textbox: make('textbox', fibRgLw.ccpTxbx),
        headerTextbox: make('headerTextbox', fibRgLw.ccpHdrTxbx),
    };
}
function readBytes(tableBytes, fc, lcb) {
    if (fc == null || lcb == null || lcb <= 0)
        return new Uint8Array(0);
    if (fc < 0 || fc >= tableBytes.length)
        return new Uint8Array(0);
    return tableBytes.subarray(fc, Math.min(tableBytes.length, fc + lcb));
}
export function readFixedPlc(tableBytes, fc, lcb, dataSize) {
    const bytes = readBytes(tableBytes, fc, lcb);
    if (!bytes.length)
        return [];
    if (dataSize < 0)
        return [];
    if (dataSize === 0) {
        const count = Math.floor(bytes.length / 4) - 1;
        if (count <= 0)
            return [];
        const reader = new BinaryReader(bytes);
        const entries = [];
        for (let i = 0; i < count; i += 1) {
            const cpStart = reader.u32(i * 4);
            const cpEnd = reader.u32((i + 1) * 4);
            entries.push({ index: i, cpStart, cpEnd, data: new Uint8Array(0) });
        }
        return entries;
    }
    if ((bytes.length - 4) < 0 || ((bytes.length - 4) % (4 + dataSize)) !== 0)
        return [];
    const count = Math.floor((bytes.length - 4) / (4 + dataSize));
    const cpsByteLength = (count + 1) * 4;
    const reader = new BinaryReader(bytes);
    const entries = [];
    for (let i = 0; i < count; i += 1) {
        const cpStart = reader.u32(i * 4);
        const cpEnd = reader.u32((i + 1) * 4);
        const dataOffset = cpsByteLength + i * dataSize;
        entries.push({ index: i, cpStart, cpEnd, data: bytes.subarray(dataOffset, dataOffset + dataSize) });
    }
    return entries;
}
export function readCpArray(tableBytes, fc, lcb) {
    const bytes = readBytes(tableBytes, fc, lcb);
    if (!bytes.length || bytes.length % 4 !== 0)
        return [];
    const reader = new BinaryReader(bytes);
    const values = [];
    for (let offset = 0; offset + 4 <= bytes.length; offset += 4)
        values.push(reader.u32(offset));
    return values;
}
function decodeUtf16LE(bytes) {
    if (!bytes.length)
        return '';
    return new TextDecoder('utf-16le').decode(bytes).replace(/\0+$/g, '');
}
export function parseSttbfRMark(tableBytes, fibRgFcLcb) {
    const fc = fibRgFcLcb.fcSttbfRMark;
    const lcb = fibRgFcLcb.lcbSttbfRMark;
    const bytes = readBytes(tableBytes, fc, lcb);
    if (bytes.length < 6)
        return [];
    const reader = new BinaryReader(bytes);
    const cData = reader.u16(2);
    const cbExtra = reader.u16(4);
    let offset = 6;
    const values = [];
    for (let i = 0; i < cData && offset + 2 <= bytes.length; i += 1) {
        const cch = reader.u16(offset);
        offset += 2;
        const byteLength = cch * 2;
        if (offset + byteLength > bytes.length)
            break;
        values.push(decodeUtf16LE(bytes.subarray(offset, offset + byteLength)));
        offset += byteLength + cbExtra;
    }
    return values;
}
export function parseXstArray(tableBytes, fibRgFcLcb) {
    const fc = fibRgFcLcb.fcGrpXstAtnOwners;
    const lcb = fibRgFcLcb.lcbGrpXstAtnOwners;
    const bytes = readBytes(tableBytes, fc, lcb);
    if (!bytes.length)
        return [];
    const reader = new BinaryReader(bytes);
    let offset = 0;
    const values = [];
    while (offset + 2 <= bytes.length) {
        const cch = reader.u16(offset);
        offset += 2;
        const byteLength = cch * 2;
        if (offset + byteLength > bytes.length)
            break;
        values.push(decodeUtf16LE(bytes.subarray(offset, offset + byteLength)));
        offset += byteLength;
    }
    return values;
}
function parseLpxCharBuffer9(bytes) {
    if (bytes.length < 20)
        return '';
    const reader = new BinaryReader(bytes);
    const cch = Math.min(reader.u16(0), 9);
    return decodeUtf16LE(bytes.subarray(2, 2 + cch * 2));
}
export function parseCommentRefMeta(data) {
    const reader = new BinaryReader(data);
    const initials = parseLpxCharBuffer9(data.subarray(0, 20));
    const authorIndex = data.length >= 22 ? reader.u16(20) : -1;
    const bookmarkId = data.length >= 30 ? reader.i32(26) : -1;
    return { initials, authorIndex, bookmarkId };
}
export function parseTextboxMeta(data) {
    const reader = new BinaryReader(data);
    return {
        reusable: Boolean(reader.u32(0) & 0x1),
        reserved: reader.u32(0),
        shapeId: reader.u32(4),
    };
}
const HEADER_SYSTEM_ROLES = [
    'footnoteSeparator',
    'footnoteContinuationSeparator',
    'footnoteContinuationNotice',
    'endnoteSeparator',
    'endnoteContinuationSeparator',
    'endnoteContinuationNotice',
];
const HEADER_SECTION_ROLES = [
    'evenHeader',
    'oddHeader',
    'evenFooter',
    'oddFooter',
    'firstHeader',
    'firstFooter',
];
const HEADER_ROLE_LABELS = {
    footnoteSeparator: 'Footnote separator',
    footnoteContinuationSeparator: 'Footnote continuation separator',
    footnoteContinuationNotice: 'Footnote continuation notice',
    endnoteSeparator: 'Endnote separator',
    endnoteContinuationSeparator: 'Endnote continuation separator',
    endnoteContinuationNotice: 'Endnote continuation notice',
    evenHeader: 'Even page header',
    oddHeader: 'Odd page header',
    evenFooter: 'Even page footer',
    oddFooter: 'Odd page footer',
    firstHeader: 'First page header',
    firstFooter: 'First page footer',
};
export function buildHeaderStoryDescriptors(tableBytes, fibRgFcLcb, headerWindow) {
    const plc = readCpArray(tableBytes, fibRgFcLcb.fcPlcfHdd, fibRgFcLcb.lcbPlcfHdd);
    if (plc.length < 2 || headerWindow.length <= 0)
        return [];
    const values = plc.slice(0, -1); // last CP is ignored by the spec
    const descriptors = [];
    const sectionEffective = new Map();
    for (let i = 0; i < values.length - 1; i += 1) {
        const localStart = values[i] ?? 0;
        const localEnd = values[i + 1] ?? 0;
        if (localEnd < localStart)
            continue;
        const isSystemRole = i < HEADER_SYSTEM_ROLES.length;
        const role = isSystemRole
            ? HEADER_SYSTEM_ROLES[i] || 'oddHeader'
            : HEADER_SECTION_ROLES[(i - HEADER_SYSTEM_ROLES.length) % HEADER_SECTION_ROLES.length] || 'oddHeader';
        const sectionIndex = isSystemRole ? undefined : Math.floor((i - HEADER_SYSTEM_ROLES.length) / HEADER_SECTION_ROLES.length) + 1;
        const descriptor = {
            index: i,
            role,
            roleLabel: HEADER_ROLE_LABELS[role] || role,
            sectionIndex,
            cpStart: headerWindow.cpStart + localStart,
            cpEnd: headerWindow.cpStart + localEnd,
        };
        if (sectionIndex && localStart === localEnd) {
            const inherited = sectionEffective.get(role);
            if (inherited)
                descriptor.inheritedFromSection = inherited.sectionIndex;
        }
        if (sectionIndex && localStart !== localEnd) {
            sectionEffective.set(role, { sectionIndex, cpStart: descriptor.cpStart, cpEnd: descriptor.cpEnd });
        }
        descriptors.push(descriptor);
    }
    return descriptors;
}
//# sourceMappingURL=stories.js.map