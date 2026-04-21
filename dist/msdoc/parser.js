import { parseCFB } from '../core/cfb.js';
import { cleanTextControlChars, pushWarning, shallowEqual, uniqueId } from '../core/utils.js';
import { parseClx, buildPieceTextCache, getTextByCp, splitParagraphRanges } from './clx.js';
import { DOC_CONTROL } from './constants.js';
import { parseFib } from './fib.js';
import { readChpxRuns, readPapxRuns } from './fkp.js';
import { parseFonts } from './fonts.js';
import { extractObjectPool, extractPictureAsset } from './objects.js';
import { buildHeaderStoryDescriptors, buildStoryWindows, parseCommentRefMeta, parseSttbfRMark, parseTextboxMeta, parseXstArray, readFixedPlc, } from './stories.js';
import { applyTableStateToCells, charPropsToState, getTableDepth, paraPropsToState, tablePropsToState, } from './properties.js';
import { mergePropertyArrays, parseStyles, splitPropertiesByKind } from './styles.js';
function getOverlappingRuns(runs, cpStart, cpEnd, cursorRef) {
    let cursor = cursorRef?.index || 0;
    while (cursor < runs.length && runs[cursor].cpEnd <= cpStart)
        cursor += 1;
    if (cursorRef)
        cursorRef.index = cursor;
    const list = [];
    let i = cursor;
    while (i < runs.length && runs[i].cpStart < cpEnd) {
        if (runs[i].cpEnd > cpStart)
            list.push(runs[i]);
        i += 1;
    }
    return list;
}
function normalizeTextStyleName(name) {
    return String(name || '').trim();
}
function decodeFieldInstruction(instruction) {
    const normalized = String(instruction || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (!normalized)
        return null;
    const upper = normalized.toUpperCase();
    if (upper.startsWith('HYPERLINK')) {
        const body = normalized.slice(9).trim();
        const quoted = body.match(/"([^"]+)"/);
        const href = quoted?.[1] || body.split(/\s+/)[0] || '';
        return href ? { type: 'hyperlink', href } : null;
    }
    if (upper.startsWith('INCLUDEPICTURE')) {
        const body = normalized.slice('INCLUDEPICTURE'.length).trim();
        const quoted = body.match(/"([^"]+)"/);
        return { type: 'includePicture', target: quoted?.[1] || body.split(/\s+/)[0] || '' };
    }
    if (upper.startsWith('EMBED')) {
        return { type: 'embed', raw: normalized };
    }
    if (upper.startsWith('LINK')) {
        return { type: 'link', raw: normalized };
    }
    return { type: 'unknown', raw: normalized };
}
function isRenderableExternalImageUrl(url) {
    return /^(?:https?:|blob:)/i.test(url) || /^data:image\//i.test(url);
}
function createFieldImageAsset(target) {
    const normalized = String(target || '').trim();
    if (!normalized)
        return null;
    let mime = 'application/octet-stream';
    if (/\.png(?:$|[?#])/i.test(normalized))
        mime = 'image/png';
    else if (/\.jpe?g(?:$|[?#])/i.test(normalized))
        mime = 'image/jpeg';
    else if (/\.gif(?:$|[?#])/i.test(normalized))
        mime = 'image/gif';
    else if (/\.bmp(?:$|[?#])/i.test(normalized))
        mime = 'image/bmp';
    else if (/\.svg(?:$|[?#])/i.test(normalized))
        mime = 'image/svg+xml';
    else if (/\.tiff?(?:$|[?#])/i.test(normalized))
        mime = 'image/tiff';
    const dataUrl = /^data:image\//i.test(normalized) ? normalized : '';
    const sourceUrl = dataUrl ? undefined : normalized;
    const localExternal = /^(?:file:|\\\\|[a-zA-Z]:[\\/]|\/)/i.test(normalized);
    return {
        id: uniqueId('asset-img-field'),
        type: 'image',
        mime,
        bytes: new Uint8Array(0),
        dataUrl,
        sourceUrl,
        displayable: isRenderableExternalImageUrl(normalized) && mime.startsWith('image/'),
        meta: {
            sourceKind: 'linked',
            linkedPath: normalized,
            localExternal,
            browserRenderable: isRenderableExternalImageUrl(normalized),
            pictureOffset: -1,
            lcb: 0,
            cbHeader: 0,
        },
    };
}
function mergeTextNode(target, node) {
    if (!node.text)
        return;
    const last = target[target.length - 1];
    if (last && last.type === 'text' && last.href === node.href && shallowEqual(last.style, node.style)) {
        last.text += node.text;
        return;
    }
    target.push(node);
}
function emitInline(targetStack, output, node) {
    if (!node)
        return;
    const target = targetStack.length ? targetStack[targetStack.length - 1].nodes : output;
    if (node.type === 'text')
        mergeTextNode(target, node);
    else
        target.push(node);
}
function getObjectPoolInfo(objectPool, pictureOffset) {
    const candidates = [
        `_${pictureOffset}`,
        `_${String(pictureOffset)}`,
        `_${pictureOffset.toString(16)}`,
        `_${pictureOffset.toString(16).toUpperCase()}`,
    ];
    for (const key of candidates) {
        if (objectPool.has(key))
            return objectPool.get(key) || null;
    }
    return null;
}
function createAssetResolver(dataBytes, objectPool, assets, usedAttachmentNames, assetCache, options = {}) {
    return function resolveAsset(charState) {
        const pictureOffset = charState?.pictureOffset;
        if (pictureOffset == null)
            return null;
        if (assetCache.has(pictureOffset))
            return assetCache.get(pictureOffset) || null;
        let asset = null;
        const objectInfo = getObjectPoolInfo(objectPool, pictureOffset);
        if ((charState.ole2 || charState.object || charState.data) && objectInfo?.attachment) {
            asset = objectInfo.attachment;
            usedAttachmentNames.add(objectInfo.entry.name);
        }
        if (!asset && dataBytes?.length) {
            const extracted = extractPictureAsset(dataBytes, pictureOffset, options);
            if (extracted && extracted.mime !== 'application/octet-stream') {
                asset = extracted;
            }
            else if (!asset && objectInfo?.attachment) {
                asset = objectInfo.attachment;
                usedAttachmentNames.add(objectInfo.entry.name);
            }
            else if (extracted) {
                asset = extracted;
            }
        }
        if (!asset && objectInfo?.attachment) {
            asset = objectInfo.attachment;
            usedAttachmentNames.add(objectInfo.entry.name);
        }
        if (asset)
            assets.push(asset);
        assetCache.set(pictureOffset, asset);
        return asset;
    };
}
function normalizePlainTextChar(ch) {
    switch (ch) {
        case DOC_CONTROL.nonBreakingHyphen:
            return '-';
        case DOC_CONTROL.nonRequiredHyphen:
            return '';
        case DOC_CONTROL.annotationRef:
        case '':
            return '';
        default:
            return ch;
    }
}
function inlineNodesToPlainText(nodes) {
    return nodes.map((node) => {
        if (node.type === 'text')
            return node.text;
        if (node.type === 'lineBreak' || node.type === 'pageBreak')
            return '\n';
        if (node.type === 'noteRef' || node.type === 'commentRef')
            return node.label;
        return '';
    }).join('');
}
function buildInlineNodes(segments, resolveAsset, context = {}) {
    const output = [];
    const fieldStack = [];
    for (const segment of segments) {
        if (!segment.text || segment.state.hidden)
            continue;
        let cp = segment.cpStart;
        for (const rawCh of segment.text) {
            const noteRef = context.noteRefs?.get(cp);
            if (noteRef) {
                emitInline(fieldStack, output, {
                    type: 'noteRef',
                    noteType: noteRef.kind,
                    refId: noteRef.id,
                    label: noteRef.label,
                });
                cp += 1;
                continue;
            }
            const commentRef = context.commentRefs?.get(cp);
            if (commentRef) {
                emitInline(fieldStack, output, {
                    type: 'commentRef',
                    commentId: commentRef.id,
                    label: commentRef.label,
                    author: commentRef.author,
                });
                cp += 1;
                continue;
            }
            const ch = normalizePlainTextChar(rawCh);
            cp += 1;
            if (!ch)
                continue;
            if (ch === DOC_CONTROL.fieldStart) {
                fieldStack.push({ instruction: '', parsed: null, readingInstruction: true, nodes: [], resultStyle: segment.state });
                continue;
            }
            if (ch === DOC_CONTROL.fieldSeparator) {
                const current = fieldStack[fieldStack.length - 1];
                if (current) {
                    current.parsed = decodeFieldInstruction(current.instruction);
                    current.readingInstruction = false;
                }
                continue;
            }
            if (ch === DOC_CONTROL.fieldEnd) {
                const current = fieldStack.pop();
                if (!current)
                    continue;
                let nodes = current.nodes;
                if (current.readingInstruction) {
                    current.parsed = decodeFieldInstruction(current.instruction);
                }
                if (current.parsed?.type === 'includePicture' && !nodes.some((node) => node.type === 'image' || node.type === 'attachment')) {
                    const asset = createFieldImageAsset(current.parsed.target);
                    if (asset)
                        nodes = [{ type: 'image', asset, style: current.resultStyle || segment.state }];
                }
                if (current.parsed?.type === 'hyperlink') {
                    const href = current.parsed.href;
                    nodes = nodes.map((node) => {
                        if (node.type === 'lineBreak' || node.type === 'pageBreak')
                            return node;
                        return { ...node, href };
                    });
                }
                const parentField = fieldStack[fieldStack.length - 1];
                if (parentField?.readingInstruction) {
                    parentField.instruction += inlineNodesToPlainText(nodes);
                }
                else {
                    for (const node of nodes)
                        emitInline(fieldStack, output, node);
                }
                continue;
            }
            const currentField = fieldStack[fieldStack.length - 1];
            if (currentField?.readingInstruction) {
                currentField.instruction += ch;
                continue;
            }
            if (currentField && !currentField.resultStyle)
                currentField.resultStyle = segment.state;
            if (ch === DOC_CONTROL.hardLineBreak) {
                emitInline(fieldStack, output, { type: 'lineBreak' });
                continue;
            }
            if (ch === DOC_CONTROL.pageBreak) {
                emitInline(fieldStack, output, { type: 'pageBreak' });
                continue;
            }
            if (ch === DOC_CONTROL.picture) {
                const asset = resolveAsset(segment.state);
                if (asset?.type === 'image') {
                    emitInline(fieldStack, output, { type: 'image', asset, style: segment.state });
                }
                else if (asset?.type === 'attachment') {
                    emitInline(fieldStack, output, { type: 'attachment', asset, style: segment.state });
                }
                continue;
            }
            emitInline(fieldStack, output, {
                type: 'text',
                text: cleanTextControlChars(ch),
                style: segment.state,
            });
        }
    }
    while (fieldStack.length) {
        const current = fieldStack.pop();
        const parentField = fieldStack[fieldStack.length - 1];
        if (parentField?.readingInstruction) {
            parentField.instruction += current.instruction + inlineNodesToPlainText(current.nodes);
            continue;
        }
        for (const node of current.nodes)
            emitInline(fieldStack, output, node);
    }
    return output;
}
function buildCharSegments(range, paragraphText, chpxRuns, styles, baseCharProps, resolveFont, cursorRef, revisionAuthors = []) {
    const overlaps = getOverlappingRuns(chpxRuns, range.cpStart, range.cpEnd, cursorRef);
    const boundaries = new Set([range.cpStart, range.cpEnd]);
    for (const run of overlaps) {
        boundaries.add(Math.max(range.cpStart, run.cpStart));
        boundaries.add(Math.min(range.cpEnd, run.cpEnd));
    }
    const points = Array.from(boundaries).sort((a, b) => a - b);
    const segments = [];
    for (let i = 0; i < points.length - 1; i += 1) {
        const cpStart = points[i];
        const cpEnd = points[i + 1];
        if (cpEnd <= cpStart)
            continue;
        const coveringRun = overlaps.find((run) => run.cpStart <= cpStart && run.cpEnd >= cpEnd);
        const directProps = coveringRun?.properties || [];
        const directState = charPropsToState(directProps);
        const charStyleProps = directState.charStyleId != null ? styles.resolveStyle(directState.charStyleId).charProps : [];
        const finalProps = mergePropertyArrays(baseCharProps, charStyleProps, directProps);
        const state = charPropsToState(finalProps);
        if (state.revisionAuthorIndex != null && revisionAuthors[state.revisionAuthorIndex]) {
            state.revisionAuthor = revisionAuthors[state.revisionAuthorIndex];
        }
        const font = resolveFont(state.fontFamilyId);
        if (font)
            state.fontFamily = font.name || font.altName || undefined;
        const localStart = cpStart - range.cpStart;
        const localEnd = cpEnd - range.cpStart;
        const text = paragraphText.slice(localStart, localEnd);
        if (!text)
            continue;
        const last = segments[segments.length - 1];
        if (last && shallowEqual(last.state, state)) {
            last.text += text;
            last.cpEnd = cpEnd;
            continue;
        }
        segments.push({ cpStart, cpEnd, text, state });
    }
    if (!segments.length && paragraphText) {
        const state = charPropsToState(baseCharProps);
        if (state.revisionAuthorIndex != null && revisionAuthors[state.revisionAuthorIndex]) {
            state.revisionAuthor = revisionAuthors[state.revisionAuthorIndex];
        }
        const font = resolveFont(state.fontFamilyId);
        if (font)
            state.fontFamily = font.name || font.altName || undefined;
        segments.push({ cpStart: range.cpStart, cpEnd: range.cpEnd, text: paragraphText, state });
    }
    return segments;
}
function buildParagraphModel(range, paragraphText, styles, fonts, chpxRuns, resolveAsset, chpxCursor, revisionAuthors = [], inlineContext = {}) {
    const directSplit = splitPropertiesByKind(range.properties || []);
    const paraStyleId = range.styleId || directSplit.para.find((item) => item.name === 'styleId')?.value || 0;
    const paraStyle = styles.resolveStyle(paraStyleId);
    const paraProps = mergePropertyArrays(paraStyle.paraProps, directSplit.para);
    const paraState = paraPropsToState(paraProps);
    const directTableState = tablePropsToState(directSplit.table);
    const tableStyleProps = directTableState.styleId != null ? styles.resolveStyle(directTableState.styleId).tableProps : [];
    const tableProps = mergePropertyArrays(tableStyleProps, directSplit.table);
    const tableState = tablePropsToState(tableProps);
    const baseCharProps = paraStyle.charProps;
    const resolveFont = (fontId) => fonts.byIndex(fontId);
    const segments = buildCharSegments(range, paragraphText, chpxRuns, styles, baseCharProps, resolveFont, chpxCursor, revisionAuthors);
    const inlines = buildInlineNodes(segments, resolveAsset, inlineContext);
    return {
        id: uniqueId('para'),
        cpStart: range.cpStart,
        cpEnd: range.cpEnd,
        terminator: range.terminator || '',
        text: paragraphText,
        rawProperties: range.properties || [],
        styleId: paraStyleId,
        styleName: normalizeTextStyleName(styles.styles.get(paraStyleId)?.name),
        paraProps,
        paraState,
        tableProps,
        tableState,
        segments,
        inlines,
    };
}
function buildRangesForCpInterval(cpStart, cpEnd, documentText, papxRuns) {
    if (cpEnd <= cpStart)
        return [];
    const rangesFromPapx = papxRuns
        .filter((run) => run.cpStart < cpEnd && run.cpEnd > cpStart)
        .map((run) => ({
        cpStart: Math.max(cpStart, run.cpStart),
        cpEnd: Math.min(cpEnd, run.cpEnd),
        terminator: documentText[Math.min(cpEnd, run.cpEnd) - 1] || '',
        styleId: run.styleId,
        properties: run.properties,
    }))
        .filter((range) => range.cpEnd > range.cpStart);
    if (rangesFromPapx.length)
        return rangesFromPapx;
    return splitParagraphRanges(documentText.slice(cpStart, cpEnd)).map((range) => ({
        cpStart: cpStart + range.cpStart,
        cpEnd: cpStart + range.cpEnd,
        terminator: range.terminator,
        styleId: 0,
        properties: [],
    }));
}
function buildParagraphModelsForInterval(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors = [], inlineContext = {}) {
    const ranges = buildRangesForCpInterval(cpStart, cpEnd, documentText, papxRuns);
    const chpxCursor = { index: 0 };
    return ranges.map((range) => {
        const rawParagraphText = getTextByCp(wordBytes, clx, pieceTexts, range.cpStart, range.cpEnd);
        const terminatorCandidate = range.terminator === DOC_CONTROL.paragraph || range.terminator === DOC_CONTROL.cellMark ? range.terminator : '';
        const paragraphText = terminatorCandidate && rawParagraphText.endsWith(terminatorCandidate)
            ? rawParagraphText.slice(0, -1)
            : rawParagraphText;
        return buildParagraphModel({ ...range, terminator: terminatorCandidate }, paragraphText, styles, fonts, chpxRuns, resolveAsset, chpxCursor, revisionAuthors, inlineContext);
    });
}
function parseIntervalToContent(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors = [], inlineContext = {}) {
    const paragraphs = buildParagraphModelsForInterval(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors, inlineContext);
    return {
        paragraphs,
        blocks: buildBlocks(paragraphs),
        text: paragraphs.map((paragraph) => paragraph.text).filter(Boolean).join('\n'),
    };
}
function finalizeTableGrid(rows) {
    for (const row of rows) {
        for (let i = 0; i < row.cells.length; i += 1) {
            const cell = row.cells[i];
            const merge = cell.meta?.merge || 0;
            cell.colIndex = i;
            cell.colspan = 1;
            cell.rowspan = 1;
            cell.hidden = false;
            if (merge === 1) {
                cell.hidden = true;
                continue;
            }
            if (merge > 1) {
                let j = i + 1;
                while (j < row.cells.length && (row.cells[j].meta?.merge || 0) === 1) {
                    row.cells[j].hidden = true;
                    cell.colspan += 1;
                    j += 1;
                }
            }
        }
    }
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex];
        for (const cell of row.cells) {
            if (cell.hidden)
                continue;
            const vertMerge = cell.meta?.vertMerge || 0;
            if (vertMerge === 1) {
                cell.hidden = true;
                continue;
            }
            if (vertMerge > 1) {
                let nextIndex = rowIndex + 1;
                while (nextIndex < rows.length) {
                    let canMerge = true;
                    for (let col = cell.colIndex || 0; col < (cell.colIndex || 0) + (cell.colspan || 1); col += 1) {
                        const nextCell = rows[nextIndex].cells[col];
                        if (!nextCell || (nextCell.meta?.vertMerge || 0) !== 1) {
                            canMerge = false;
                            break;
                        }
                    }
                    if (!canMerge)
                        break;
                    for (let col = cell.colIndex || 0; col < (cell.colIndex || 0) + (cell.colspan || 1); col += 1) {
                        rows[nextIndex].cells[col].hidden = true;
                    }
                    cell.rowspan = (cell.rowspan || 1) + 1;
                    nextIndex += 1;
                }
            }
        }
    }
}
function paragraphToBlock(paragraph) {
    return {
        type: 'paragraph',
        id: paragraph.id,
        styleId: paragraph.styleId,
        styleName: paragraph.styleName,
        paraState: paragraph.paraState,
        inlines: paragraph.inlines,
        text: paragraph.text,
    };
}
function buildTableBlock(tableParagraphs) {
    const rows = [];
    let pendingRow = { cells: [] };
    let pendingCellParagraphs = [];
    for (const paragraph of tableParagraphs) {
        pendingCellParagraphs.push(paragraph);
        if (paragraph.terminator === DOC_CONTROL.cellMark) {
            pendingRow.cells.push({
                id: uniqueId('cell'),
                paragraphs: pendingCellParagraphs.map(paragraphToBlock),
                meta: null,
            });
            pendingCellParagraphs = [];
            if (paragraph.paraState.tableRowEnd || paragraph.paraState.innerTableRowEnd) {
                const cellDefs = applyTableStateToCells(paragraph.tableState);
                while (cellDefs.length &&
                    pendingRow.cells.length > cellDefs.length &&
                    pendingRow.cells[pendingRow.cells.length - 1].paragraphs.every((block) => !block.text && !(block.inlines || []).length)) {
                    pendingRow.cells.pop();
                }
                pendingRow.cells.forEach((cell, index) => {
                    cell.meta = cellDefs[index] || { index };
                });
                const gridWidthTwips = cellDefs.length
                    ? ((cellDefs[cellDefs.length - 1].rightBoundary || 0) - (cellDefs[0].leftBoundary || 0))
                    : 0;
                rows.push({
                    id: uniqueId('row'),
                    cells: pendingRow.cells,
                    state: paragraph.tableState,
                    gridWidthTwips,
                });
                pendingRow = { cells: [] };
            }
        }
    }
    if (pendingCellParagraphs.length) {
        pendingRow.cells.push({ id: uniqueId('cell'), paragraphs: pendingCellParagraphs.map(paragraphToBlock), meta: null });
    }
    if (pendingRow.cells.length) {
        rows.push({ id: uniqueId('row'), cells: pendingRow.cells, state: tableParagraphs[0]?.tableState || tablePropsToState([]), gridWidthTwips: 0 });
    }
    finalizeTableGrid(rows);
    const gridWidthTwips = rows.find((row) => row.gridWidthTwips)?.gridWidthTwips || 0;
    const depth = Math.max(...tableParagraphs.map((paragraph) => getTableDepth(paragraph.paraState)), 1);
    return {
        type: 'table',
        id: uniqueId('table'),
        depth,
        rows,
        state: rows[0]?.state || tablePropsToState([]),
        gridWidthTwips,
    };
}
function buildBlocks(paragraphs) {
    const blocks = [];
    let index = 0;
    while (index < paragraphs.length) {
        const paragraph = paragraphs[index];
        const depth = getTableDepth(paragraph.paraState);
        if (depth <= 0) {
            blocks.push(paragraphToBlock(paragraph));
            index += 1;
            continue;
        }
        const tableParagraphs = [];
        while (index < paragraphs.length && getTableDepth(paragraphs[index].paraState) > 0) {
            tableParagraphs.push(paragraphs[index]);
            index += 1;
        }
        blocks.push(buildTableBlock(tableParagraphs));
    }
    return blocks;
}
function collectAssetWarnings(assets, warnings) {
    for (const asset of assets) {
        if (asset.type !== 'image')
            continue;
        if (asset.mime === 'application/octet-stream') {
            pushWarning(warnings, 'Encountered picture data that could not be decoded to a supported image payload', {
                code: 'unsupported-image-payload',
                severity: 'warning',
                offset: asset.meta?.pictureOffset,
                details: { mime: asset.mime, sourceKind: asset.meta?.sourceKind },
            });
            continue;
        }
        if (asset.displayable === false) {
            pushWarning(warnings, 'Encountered an image that was parsed but is not directly browser-displayable', {
                code: asset.meta?.localExternal ? 'linked-local-image' : 'non-displayable-image',
                severity: 'warning',
                offset: asset.meta?.pictureOffset,
                details: { mime: asset.mime, sourceUrl: asset.sourceUrl, linkedPath: asset.meta?.linkedPath },
            });
        }
    }
}
function buildNoteItems(kind, storyCpBase, textEntries, refEntries, parseContent) {
    const items = [];
    const refMap = new Map();
    for (let i = 0; i < textEntries.length; i += 1) {
        const entry = textEntries[i];
        if (entry.cpEnd <= entry.cpStart)
            continue;
        const id = uniqueId(kind);
        const label = String(i + 1);
        const refCp = refEntries[i]?.cpStart;
        if (refCp != null)
            refMap.set(refCp, { kind, id, label });
        const content = parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd);
        items.push({
            id,
            kind,
            index: i,
            label,
            refCp,
            blocks: content.blocks,
            text: content.text,
        });
    }
    return { items, refMap };
}
function buildCommentItems(storyCpBase, textEntries, refEntries, commentAuthors, revisionAuthors, parseContent) {
    const items = [];
    const refMap = new Map();
    for (let i = 0; i < textEntries.length; i += 1) {
        const entry = textEntries[i];
        if (entry.cpEnd <= entry.cpStart)
            continue;
        const meta = parseCommentRefMeta(refEntries[i]?.data || new Uint8Array(0));
        const author = commentAuthors[meta.authorIndex] || revisionAuthors[meta.authorIndex] || meta.initials || undefined;
        const id = uniqueId('comment');
        const label = String(i + 1);
        const refCp = refEntries[i]?.cpStart;
        if (refCp != null)
            refMap.set(refCp, { id, label, author });
        const content = parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd);
        items.push({
            id,
            index: i,
            label,
            refCp,
            author,
            initials: meta.initials || undefined,
            bookmarkId: meta.bookmarkId,
            blocks: content.blocks,
            text: content.text,
        });
    }
    return { items, refMap };
}
function buildTextboxItems(header, storyCpBase, entries, parseContent) {
    const items = [];
    for (const entry of entries) {
        const meta = parseTextboxMeta(entry.data);
        const content = entry.cpEnd > entry.cpStart
            ? parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd)
            : { paragraphs: [], blocks: [], text: '' };
        if (!content.blocks.length && !content.text && meta.reusable)
            continue;
        items.push({
            id: uniqueId(header ? 'hdr-textbox' : 'textbox'),
            index: entry.index,
            label: `${header ? 'Header textbox' : 'Textbox'} ${entry.index + 1}`,
            header,
            reusable: meta.reusable,
            shapeId: meta.shapeId || undefined,
            blocks: content.blocks,
            text: content.text,
        });
    }
    return items;
}
function buildHeaderStories(descriptors, parseContent) {
    const stories = [];
    const latestByRole = new Map();
    for (const descriptor of descriptors) {
        let blocks = [];
        let text = '';
        if (descriptor.cpEnd > descriptor.cpStart) {
            const content = parseContent(descriptor.cpStart, descriptor.cpEnd);
            blocks = content.blocks;
            text = content.text;
        }
        else if (descriptor.inheritedFromSection != null) {
            const inherited = latestByRole.get(descriptor.role);
            if (inherited) {
                blocks = inherited.blocks;
                text = inherited.text;
            }
        }
        if (!blocks.length && !text)
            continue;
        const story = {
            id: uniqueId('header-story'),
            role: descriptor.role,
            roleLabel: descriptor.roleLabel,
            sectionIndex: descriptor.sectionIndex,
            inheritedFromSection: descriptor.inheritedFromSection,
            blocks,
            text,
        };
        stories.push(story);
        if (text || blocks.length)
            latestByRole.set(descriptor.role, story);
    }
    return stories;
}
/**
 * Main MS-DOC entry point.
 * It parses the OLE container, restores text through the piece table, resolves
 * paragraph/character/table properties, and finally produces a normalized AST
 * that the HTML renderer can consume.
 */
export function parseMsDoc(input, options = {}) {
    const warnings = [];
    const cfb = parseCFB(input, options);
    warnings.push(...(cfb.warnings || []));
    const wordBytes = cfb.getStream('/WordDocument');
    if (!wordBytes)
        throw new Error('Missing WordDocument stream');
    const fib = parseFib(wordBytes);
    if (fib.base.wIdent !== 0xA5EC) {
        pushWarning(warnings, `Unexpected FIB identifier: 0x${fib.base.wIdent.toString(16)}`);
    }
    if (fib.base.fEncrypted) {
        throw new Error('Encrypted .doc files are not supported yet');
    }
    const tableBytes = cfb.getStream(fib.base.fWhichTblStm ? '/1Table' : '/0Table');
    if (!tableBytes)
        throw new Error('Missing table stream');
    const dataBytes = cfb.getStream('/Data') || new Uint8Array(0);
    const clx = parseClx(tableBytes, fib.fibRgFcLcb);
    const pieceTexts = buildPieceTextCache(wordBytes, clx);
    const documentText = pieceTexts.join('');
    const storyWindows = buildStoryWindows(fib.fibRgLw, documentText.length);
    const styles = parseStyles(tableBytes, fib.fibRgFcLcb);
    const fonts = parseFonts(tableBytes, fib.fibRgFcLcb);
    const revisionAuthors = parseSttbfRMark(tableBytes, fib.fibRgFcLcb);
    const commentAuthors = parseXstArray(tableBytes, fib.fibRgFcLcb);
    const chpxRuns = readChpxRuns(wordBytes, tableBytes, fib, clx)
        .filter((run) => run.cpStart < documentText.length)
        .map((run) => ({ ...run, cpEnd: Math.min(run.cpEnd, documentText.length) }));
    const papxRuns = readPapxRuns(wordBytes, tableBytes, fib, clx)
        .filter((run) => run.cpStart < documentText.length)
        .map((run) => ({ ...run, cpEnd: Math.min(run.cpEnd, documentText.length) }));
    const objectPool = extractObjectPool(cfb);
    const assets = [];
    const usedAttachmentNames = new Set();
    const assetCache = new Map();
    const resolveAsset = createAssetResolver(dataBytes, objectPool, assets, usedAttachmentNames, assetCache, options);
    const parseContent = (cpStart, cpEnd, inlineContext = {}) => parseIntervalToContent(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors, inlineContext);
    const footnoteTextEntries = readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcffndTxt, fib.fibRgFcLcb.lcbPlcffndTxt, 0);
    const footnoteRefEntries = readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcffndRef, fib.fibRgFcLcb.lcbPlcffndRef, 2).map((entry) => ({ index: entry.index, cpStart: entry.cpStart }));
    const { items: footnoteItems, refMap: footnoteRefMap } = buildNoteItems('footnote', storyWindows.footnote.cpStart, footnoteTextEntries, footnoteRefEntries, parseContent);
    const endnoteTextEntries = readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcfendTxt, fib.fibRgFcLcb.lcbPlcfendTxt, 0);
    const endnoteRefEntries = readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcfendRef, fib.fibRgFcLcb.lcbPlcfendRef, 2).map((entry) => ({ index: entry.index, cpStart: entry.cpStart }));
    const { items: endnoteItems, refMap: endnoteRefMap } = buildNoteItems('endnote', storyWindows.endnote.cpStart, endnoteTextEntries, endnoteRefEntries, parseContent);
    const commentTextEntries = readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcfandTxt, fib.fibRgFcLcb.lcbPlcfandTxt, 0);
    const commentRefEntries = readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcfandRef, fib.fibRgFcLcb.lcbPlcfandRef, 30);
    const { items: commentItems, refMap: commentRefMap } = buildCommentItems(storyWindows.comment.cpStart, commentTextEntries, commentRefEntries, commentAuthors, revisionAuthors, parseContent);
    const noteRefs = new Map();
    for (const [cp, info] of footnoteRefMap.entries())
        noteRefs.set(cp, info);
    for (const [cp, info] of endnoteRefMap.entries())
        noteRefs.set(cp, info);
    const mainContent = parseContent(storyWindows.main.cpStart, storyWindows.main.cpEnd, {
        noteRefs,
        commentRefs: commentRefMap,
    });
    const headerStories = buildHeaderStories(buildHeaderStoryDescriptors(tableBytes, fib.fibRgFcLcb, storyWindows.header), parseContent);
    const textboxItems = buildTextboxItems(false, storyWindows.textbox.cpStart, readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcftxbxTxt, fib.fibRgFcLcb.lcbPlcftxbxTxt, 22), parseContent);
    const headerTextboxItems = buildTextboxItems(true, storyWindows.headerTextbox.cpStart, readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcfHdrtxbxTxt, fib.fibRgFcLcb.lcbPlcfHdrtxbxTxt, 22), parseContent);
    const blocks = [...mainContent.blocks];
    if (footnoteItems.length) {
        const footnotesBlock = { type: 'notes', id: uniqueId('notes-footnote'), kind: 'footnote', items: footnoteItems };
        blocks.push(footnotesBlock);
    }
    if (endnoteItems.length) {
        const endnotesBlock = { type: 'notes', id: uniqueId('notes-endnote'), kind: 'endnote', items: endnoteItems };
        blocks.push(endnotesBlock);
    }
    if (commentItems.length) {
        const commentsBlock = { type: 'comments', id: uniqueId('comments'), items: commentItems };
        blocks.push(commentsBlock);
    }
    if (headerStories.length) {
        const headersBlock = { type: 'headers', id: uniqueId('headers'), stories: headerStories };
        blocks.push(headersBlock);
    }
    if (textboxItems.length) {
        const textboxesBlock = { type: 'textboxes', id: uniqueId('textboxes'), header: false, items: textboxItems };
        blocks.push(textboxesBlock);
    }
    if (headerTextboxItems.length) {
        const headerTextboxesBlock = { type: 'textboxes', id: uniqueId('header-textboxes'), header: true, items: headerTextboxItems };
        blocks.push(headerTextboxesBlock);
    }
    const trailingAttachments = Array.from(objectPool.values())
        .filter((item) => item?.attachment && !usedAttachmentNames.has(item.entry.name))
        .map((item) => item.attachment);
    for (const attachment of trailingAttachments)
        assets.push(attachment);
    if (trailingAttachments.length) {
        blocks.push({ type: 'attachments', id: uniqueId('attachments'), items: trailingAttachments });
    }
    collectAssetWarnings(assets, warnings);
    const countInnerBlocks = (innerBlocks) => innerBlocks.reduce((sum, block) => {
        if (block.type === 'paragraph')
            return sum + 1;
        return sum + block.rows.reduce((rowSum, row) => rowSum + row.cells.reduce((cellSum, cell) => cellSum + cell.paragraphs.length, 0), 0);
    }, 0);
    const paragraphCount = blocks.reduce((sum, block) => {
        if (block.type === 'paragraph' || block.type === 'table') {
            return sum + countInnerBlocks([block]);
        }
        if (block.type === 'notes') {
            return sum + block.items.reduce((itemSum, item) => itemSum + countInnerBlocks(item.blocks), 0);
        }
        if (block.type === 'comments') {
            return sum + block.items.reduce((itemSum, item) => itemSum + countInnerBlocks(item.blocks), 0);
        }
        if (block.type === 'headers') {
            return sum + block.stories.reduce((storySum, story) => storySum + countInnerBlocks(story.blocks), 0);
        }
        if (block.type === 'textboxes') {
            return sum + block.items.reduce((itemSum, item) => itemSum + countInnerBlocks(item.blocks), 0);
        }
        return sum;
    }, 0);
    return {
        kind: 'msdoc',
        version: 1,
        warnings,
        meta: {
            fib: {
                wIdent: fib.base.wIdent,
                nFib: fib.base.nFib,
                fWhichTblStm: fib.base.fWhichTblStm,
                fComplex: fib.base.fComplex,
                fEncrypted: fib.base.fEncrypted,
                ccpText: fib.fibRgLw.ccpText,
            },
            counts: {
                paragraphs: paragraphCount,
                blocks: blocks.length,
                assets: assets.length,
                styles: styles.styles.size,
                fonts: fonts.fonts.length,
                footnotes: footnoteItems.length,
                endnotes: endnoteItems.length,
                comments: commentItems.length,
                headers: headerStories.length,
                textboxes: textboxItems.length,
                headerTextboxes: headerTextboxItems.length,
            },
        },
        fonts: fonts.fonts,
        styles: Array.from(styles.styles.values()).map((style) => ({
            istd: style.istd,
            name: style.name,
            type: style.stdfBase?.stk,
            basedOn: style.stdfBase?.istdBase,
            next: style.stdfBase?.istdNext,
        })),
        blocks,
        assets,
    };
}
//# sourceMappingURL=parser.js.map