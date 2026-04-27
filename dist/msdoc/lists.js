import { BinaryReader } from '../core/binary.js';
import { decodeGrpprl } from './sprm.js';
function readBytes(tableBytes, fc, lcb) {
    if (fc == null || lcb == null || lcb <= 0)
        return new Uint8Array(0);
    if (fc < 0 || fc >= tableBytes.length)
        return new Uint8Array(0);
    return tableBytes.subarray(fc, Math.min(tableBytes.length, fc + lcb));
}
function decodeUtf16LE(bytes) {
    if (!bytes.length)
        return '';
    return new TextDecoder('utf-16le').decode(bytes).replace(/\0+$/g, '');
}
function normalizeLevelNumberFormat(nfc) {
    switch (nfc) {
        case 0: return 'decimal';
        case 1: return 'upperRoman';
        case 2: return 'lowerRoman';
        case 3: return 'upperLetter';
        case 4: return 'lowerLetter';
        case 23: return 'bullet';
        default: return 'decimal';
    }
}
function toRoman(input, upper) {
    let value = Math.max(1, Math.min(3999, input | 0));
    const pairs = [
        [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'],
        [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
    ];
    let out = '';
    for (const [amount, symbol] of pairs) {
        while (value >= amount) {
            out += symbol;
            value -= amount;
        }
    }
    return upper ? out.toUpperCase() : out;
}
function toLetters(input, upper) {
    let value = Math.max(1, input | 0);
    let out = '';
    while (value > 0) {
        value -= 1;
        out = String.fromCharCode(97 + (value % 26)) + out;
        value = Math.floor(value / 26);
    }
    return upper ? out.toUpperCase() : out;
}
function formatNumber(value, nfc) {
    switch (normalizeLevelNumberFormat(nfc)) {
        case 'upperRoman': return toRoman(value, true);
        case 'lowerRoman': return toRoman(value, false);
        case 'upperLetter': return toLetters(value, true);
        case 'lowerLetter': return toLetters(value, false);
        case 'bullet': return '•';
        default: return String(value);
    }
}
function parseLevel(bytes, offset, level) {
    if (offset + 28 > bytes.length)
        return null;
    const reader = new BinaryReader(bytes);
    const startAt = Math.max(0, reader.i32(offset));
    const nfc = reader.u8(offset + 4);
    const info = reader.u8(offset + 5);
    const alignment = info & 0x03;
    const legal = Boolean(info & 0x04);
    const noRestart = Boolean(info & 0x08);
    const placeholderOffsets = [];
    for (let i = 0; i < 9; i += 1) {
        const value = reader.u8(offset + 6 + i);
        if (!value)
            break;
        // rgbxchNums stores one-based character offsets into xst.rgtchar, not
        // level numbers. The level number itself is the control character stored at
        // that offset in the LVL template string.
        placeholderOffsets.push(value - 1);
    }
    const followByte = reader.u8(offset + 15);
    const follow = followByte === 0 ? 'tab' : followByte === 1 ? 'space' : 'none';
    const cbGrpprlChpx = reader.u8(offset + 24);
    const cbGrpprlPapx = reader.u8(offset + 25);
    let cursor = offset + 28;
    // LVL stores paragraph and character property modifiers before the Xst.
    // Producers follow the spec order (PAPX then CHPX); if a malformed document
    // points past the buffer, return a minimal level so numbering can still work.
    const papxStart = cursor;
    const papxEnd = Math.min(bytes.length, papxStart + cbGrpprlPapx);
    cursor = papxEnd;
    const chpxStart = cursor;
    const chpxEnd = Math.min(bytes.length, chpxStart + cbGrpprlChpx);
    cursor = chpxEnd;
    let template = '';
    if (cursor + 2 <= bytes.length) {
        const cch = reader.u16(cursor);
        cursor += 2;
        const byteLength = cch * 2;
        if (cursor + byteLength <= bytes.length) {
            template = decodeUtf16LE(bytes.subarray(cursor, cursor + byteLength));
            cursor += byteLength;
        }
    }
    const placeholderLevels = placeholderOffsets
        .map((charIndex) => template.charCodeAt(charIndex))
        .filter((charCode) => charCode >= 0 && charCode <= 8);
    return {
        level: {
            level,
            startAt,
            numberFormat: nfc,
            alignment,
            follow,
            template,
            placeholderLevels,
            legal,
            noRestart,
            paraProps: decodeGrpprl(bytes, papxStart, papxEnd),
            charProps: decodeGrpprl(bytes, chpxStart, chpxEnd),
        },
        nextOffset: cursor,
    };
}
function parsePlfLst(tableBytes, fib) {
    const bytes = readBytes(tableBytes, fib.fcPlfLst, fib.lcbPlfLst);
    if (bytes.length < 2)
        return [];
    const reader = new BinaryReader(bytes);
    const count = reader.u16(0);
    let offset = 2;
    const definitions = [];
    for (let i = 0; i < count && offset + 28 <= bytes.length; i += 1) {
        const listId = reader.i32(offset);
        const tplc = reader.i32(offset + 4);
        const styleIds = [];
        for (let level = 0; level < 9; level += 1)
            styleIds.push(reader.i16(offset + 8 + level * 2));
        const flags = reader.u8(offset + 26);
        const simple = Boolean(flags & 0x01);
        const autoNumber = Boolean(flags & 0x04);
        const hybrid = Boolean(flags & 0x10);
        offset += 28;
        const levelCount = simple ? 1 : 9;
        const levels = [];
        for (let level = 0; level < levelCount; level += 1) {
            const parsed = parseLevel(bytes, offset, level);
            if (!parsed)
                break;
            levels.push(parsed.level);
            offset = parsed.nextOffset;
        }
        definitions.push({ listId, tplc, styleIds, simple, autoNumber, hybrid, levels });
    }
    return definitions;
}
function parsePlfLfo(tableBytes, fib) {
    const bytes = readBytes(tableBytes, fib.fcPlfLfo, fib.lcbPlfLfo);
    if (bytes.length < 4)
        return [];
    const reader = new BinaryReader(bytes);
    const count = reader.u32(0);
    let offset = 4;
    const overrides = [];
    for (let i = 0; i < count && offset + 16 <= bytes.length; i += 1) {
        const listId = reader.i32(offset);
        const overrideCount = reader.u8(offset + 12);
        const fieldAutoNumberKind = reader.u8(offset + 13);
        overrides.push({ index: i + 1, listId, overrideCount, fieldAutoNumberKind, levels: [] });
        offset += 16;
    }
    for (const override of overrides) {
        const levels = [];
        for (let i = 0; i < override.overrideCount && offset + 8 <= bytes.length; i += 1) {
            const startAt = reader.i32(offset);
            const packed = reader.u32(offset + 4);
            const level = packed & 0x0f;
            const fStartAt = Boolean(packed & 0x10);
            const fFormatting = Boolean(packed & 0x20);
            offset += 8;
            let levelOverride;
            if (fFormatting) {
                const parsed = parseLevel(bytes, offset, level);
                if (parsed) {
                    levelOverride = parsed.level;
                    offset = parsed.nextOffset;
                }
            }
            levels.push({ level, startAt: fStartAt ? startAt : undefined, levelOverride });
        }
        override.levels = levels;
    }
    return overrides;
}
export function parseLists(tableBytes, fib) {
    return {
        definitions: parsePlfLst(tableBytes, fib),
        overrides: parsePlfLfo(tableBytes, fib),
    };
}
function findOverride(lists, overrideIndex) {
    if (!overrideIndex)
        return undefined;
    return lists.overrides.find((item) => item.index === overrideIndex) || lists.overrides[overrideIndex - 1];
}
function findLevel(lists, overrideIndex, levelIndex) {
    const override = findOverride(lists, overrideIndex);
    const definition = override
        ? lists.definitions.find((item) => item.listId === override.listId)
        : undefined;
    const overrideLevel = override?.levels.find((item) => item.level === levelIndex);
    const level = overrideLevel?.levelOverride
        || definition?.levels[levelIndex]
        || definition?.levels[0];
    return { definition, level, override };
}
function renderLevelTemplate(template, levelIndex, counters, currentLevel) {
    const placeholders = currentLevel.placeholderLevels.length
        ? currentLevel.placeholderLevels
        : Array.from({ length: levelIndex + 1 }, (_, index) => index);
    let out = template || '';
    if (out) {
        for (const placeholderLevel of placeholders) {
            const control = String.fromCharCode(placeholderLevel);
            const value = counters[placeholderLevel] || 0;
            out = out.split(control).join(formatNumber(value, currentLevel.numberFormat));
        }
        out = out.replace(/[\u0000-\u0008\u000e-\u001f]/g, '');
        if (out.trim())
            return out;
    }
    const joined = placeholders
        .map((placeholderLevel) => formatNumber(counters[placeholderLevel] || 1, placeholderLevel === levelIndex ? currentLevel.numberFormat : 0))
        .join('.');
    return currentLevel.numberFormat === 23 ? '•' : `${joined}.`;
}
/**
 * Applies list labels to paragraph models using sprmPIlfo/sprmPIlvl and the
 * PlfLst/PlfLfo model. This intentionally computes display labels in the AST so
 * renderers do not need to emulate Word's numbering counters themselves.
 */
export function applyListFormatting(paragraphs, lists, options = {}) {
    const countersByOverride = new Map();
    for (const paragraph of paragraphs) {
        const overrideIndex = paragraph.paraState.listId || 0;
        const levelIndex = paragraph.paraState.listLevel ?? -1;
        if (!overrideIndex || levelIndex < 0) {
            continue;
        }
        const { level, override } = findLevel(lists, overrideIndex, Math.max(0, Math.min(8, levelIndex)));
        const effectiveLevel = level || {
            level: levelIndex,
            startAt: 1,
            numberFormat: 0,
            alignment: 0,
            follow: 'tab',
            template: '',
            placeholderLevels: [],
            legal: false,
            noRestart: false,
            paraProps: [],
            charProps: [],
        };
        const counters = countersByOverride.get(overrideIndex) || new Array(9).fill(0);
        const overrideLevel = override?.levels.find((item) => item.level === levelIndex);
        if (!counters[levelIndex])
            counters[levelIndex] = overrideLevel?.startAt ?? (effectiveLevel.startAt || 1);
        else
            counters[levelIndex] += 1;
        for (let i = levelIndex + 1; i < counters.length; i += 1)
            counters[i] = 0;
        countersByOverride.set(overrideIndex, counters);
        paragraph.list = {
            listId: override?.listId ?? overrideIndex,
            overrideIndex,
            level: levelIndex,
            label: renderLevelTemplate(effectiveLevel.template, levelIndex, counters, effectiveLevel),
            format: effectiveLevel.numberFormat,
            template: effectiveLevel.template,
            follow: effectiveLevel.follow,
            style: options.resolveLabelStyle?.(effectiveLevel.charProps, paragraph, effectiveLevel),
        };
    }
}
//# sourceMappingURL=lists.js.map