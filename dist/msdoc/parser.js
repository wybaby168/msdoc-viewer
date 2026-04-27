import { parseCFB } from '../core/cfb.js';
import { cleanTextControlChars, pushWarning, shallowEqual, uniqueId } from '../core/utils.js';
import { parseBookmarks } from './bookmarks.js';
import { parseClx, buildPieceTextCache, getTextByCp, splitParagraphRanges } from './clx.js';
import { DOC_CONTROL } from './constants.js';
import { parseDop } from './dop.js';
import { parseFib } from './fib.js';
import { readChpxRuns, readPapxRuns } from './fkp.js';
import { parseFonts } from './fonts.js';
import { applyListFormatting, parseLists } from './lists.js';
import { parseDrawingGroup, resolveHeaderAnchorBinding } from './drawings.js';
import { extractObjectPool, extractPictureAsset } from './objects.js';
import { readShapeAnchors } from './shapes.js';
import { findSectionIndex, readSections } from './sections.js';
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
function fieldTokens(value) {
    return value.match(/"[^"]*"|\S+/g)?.map((token) => token.replace(/^"|"$/g, '')) || [];
}
function decodeFieldInstruction(instruction) {
    const normalized = String(instruction || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (!normalized)
        return null;
    const tokens = fieldTokens(normalized);
    const keyword = String(tokens[0] || '').toUpperCase();
    const firstArg = tokens.find((token, index) => index > 0 && !token.startsWith('\\')) || '';
    if (keyword === 'HYPERLINK') {
        const href = firstArg || '';
        return href ? { type: 'hyperlink', href } : null;
    }
    if (keyword === 'INCLUDEPICTURE') {
        return { type: 'includePicture', target: firstArg || '' };
    }
    if (keyword === 'PAGE')
        return { type: 'page', raw: normalized };
    if (keyword === 'NUMPAGES' || keyword === 'SECTIONPAGES')
        return { type: 'numpages', raw: normalized };
    if (keyword === 'REF')
        return { type: 'ref', target: firstArg || '', raw: normalized };
    if (keyword === 'PAGEREF')
        return { type: 'pageref', target: firstArg || '', raw: normalized };
    if (keyword === 'SEQ')
        return { type: 'seq', name: firstArg || 'SEQ', raw: normalized };
    if (keyword === 'DATE' || keyword === 'SAVEDATE' || keyword === 'CREATEDATE' || keyword === 'PRINTDATE') {
        const formatSwitchIndex = tokens.findIndex((token) => token === '\\@');
        return { type: 'date', format: formatSwitchIndex >= 0 ? tokens[formatSwitchIndex + 1] : undefined, raw: normalized };
    }
    if (keyword === 'TIME')
        return { type: 'time', raw: normalized };
    if (keyword === 'TOC')
        return { type: 'toc', raw: normalized };
    if (keyword === 'MERGEFIELD')
        return { type: 'mergefield', name: firstArg || '', raw: normalized };
    if (keyword === 'FORMTEXT')
        return { type: 'formtext', raw: normalized };
    if (keyword === 'EMBED')
        return { type: 'embed', raw: normalized };
    if (keyword === 'LINK')
        return { type: 'link', raw: normalized };
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
function fieldFallbackText(parsed, context) {
    if (!parsed)
        return null;
    switch (parsed.type) {
        case 'page':
            return { text: 'PAGE' };
        case 'numpages':
            return { text: 'NUMPAGES' };
        case 'seq': {
            const counters = context.sequenceCounters || new Map();
            context.sequenceCounters = counters;
            const key = parsed.name || 'SEQ';
            const next = (counters.get(key) || 0) + 1;
            counters.set(key, next);
            return { text: String(next) };
        }
        case 'date':
        case 'time':
            return { text: new Date().toLocaleString() };
        case 'ref': {
            const bookmark = context.bookmarksByName?.get(parsed.target) || context.bookmarksByName?.get(parsed.target.toLowerCase());
            return { text: bookmark?.name || parsed.target, target: bookmark?.id, href: bookmark ? `#${bookmark.id}` : undefined };
        }
        case 'pageref': {
            const bookmark = context.bookmarksByName?.get(parsed.target) || context.bookmarksByName?.get(parsed.target.toLowerCase());
            return { text: bookmark ? 'PAGE' : parsed.target, target: bookmark?.id, href: bookmark ? `#${bookmark.id}` : undefined };
        }
        case 'mergefield':
            return { text: parsed.name ? `«${parsed.name}»` : '«MERGEFIELD»' };
        case 'formtext':
            return { text: '□' };
        case 'toc':
            return { text: '' };
        default:
            return null;
    }
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
function applyDrawingInfoToShapeAnchor(anchor, drawingInfo) {
    const info = drawingInfo.get(anchor.shapeId);
    if (!info)
        return anchor;
    return {
        ...anchor,
        drawingName: info.name,
        drawingDescription: info.description,
        shapeTypeCode: info.shapeTypeCode,
        blipRef: info.blipRef,
        imageAssetId: info.imageAssetId,
        imageAsset: info.imageAsset,
    };
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
        if (node.type === 'field')
            return node.displayText;
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
                if (!nodes.length) {
                    const fallback = fieldFallbackText(current.parsed, context);
                    if (fallback?.text) {
                        nodes = [{
                                type: 'field',
                                fieldType: current.parsed?.type || 'unknown',
                                instruction: current.instruction,
                                displayText: fallback.text,
                                target: fallback.target,
                                href: fallback.href,
                                style: current.resultStyle || segment.state,
                            }];
                    }
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
function buildCharStateAtCp(cp, chpxRuns, styles, baseCharProps, resolveFont, revisionAuthors = [], cursorRef) {
    const overlaps = getOverlappingRuns(chpxRuns, cp, cp + 1, cursorRef);
    const directRun = overlaps.find((run) => run.cpStart <= cp && run.cpEnd > cp);
    const directProps = directRun?.properties || [];
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
    return state;
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
    const markStyle = range.cpEnd > range.cpStart
        ? buildCharStateAtCp(Math.max(range.cpStart, range.cpEnd - 1), chpxRuns, styles, baseCharProps, resolveFont, revisionAuthors, chpxCursor)
        : (segments[segments.length - 1]?.state || buildCharStateAtCp(range.cpStart, chpxRuns, styles, baseCharProps, resolveFont, revisionAuthors, chpxCursor));
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
        markStyle,
        list: undefined,
        bookmarkStarts: undefined,
        bookmarkEnds: undefined,
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
function resolveListLabelExplicitStyle(listCharProps, styles, fonts) {
    if (!listCharProps.length)
        return undefined;
    const directState = charPropsToState(listCharProps);
    const charStyleProps = directState.charStyleId != null ? styles.resolveStyle(directState.charStyleId).charProps : [];
    const finalState = charPropsToState(mergePropertyArrays(charStyleProps, listCharProps));
    const font = fonts.byIndex(finalState.fontFamilyId);
    if (font)
        finalState.fontFamily = font.name || font.altName || undefined;
    return finalState;
}
function assignBookmarksToParagraphs(paragraphs, bookmarks) {
    if (!bookmarks.length || !paragraphs.length)
        return;
    for (const paragraph of paragraphs) {
        const starts = bookmarks.filter((bookmark) => bookmark.cpStart >= paragraph.cpStart && bookmark.cpStart <= paragraph.cpEnd);
        const ends = bookmarks.filter((bookmark) => bookmark.cpEnd >= paragraph.cpStart && bookmark.cpEnd <= paragraph.cpEnd);
        if (starts.length)
            paragraph.bookmarkStarts = starts;
        if (ends.length)
            paragraph.bookmarkEnds = ends;
    }
}
function buildParagraphModelsForInterval(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors = [], inlineContext = {}, resolveSectionIndex = () => 0, lists, bookmarks = []) {
    const ranges = buildRangesForCpInterval(cpStart, cpEnd, documentText, papxRuns);
    const chpxCursor = { index: 0 };
    const models = ranges.map((range) => {
        const rawParagraphText = getTextByCp(wordBytes, clx, pieceTexts, range.cpStart, range.cpEnd);
        const terminatorCandidate = range.terminator === DOC_CONTROL.paragraph || range.terminator === DOC_CONTROL.cellMark ? range.terminator : '';
        const paragraphText = terminatorCandidate && rawParagraphText.endsWith(terminatorCandidate)
            ? rawParagraphText.slice(0, -1)
            : rawParagraphText;
        const model = buildParagraphModel({ ...range, terminator: terminatorCandidate }, paragraphText, styles, fonts, chpxRuns, resolveAsset, chpxCursor, revisionAuthors, inlineContext);
        model.sectionIndex = resolveSectionIndex(range.cpStart);
        return model;
    });
    if (lists) {
        applyListFormatting(models, lists, {
            resolveLabelStyle: (listCharProps) => resolveListLabelExplicitStyle(listCharProps, styles, fonts),
        });
    }
    assignBookmarksToParagraphs(models, bookmarks);
    return models;
}
function parseIntervalToContent(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors = [], inlineContext = {}, resolveSectionIndex = () => 0, lists, bookmarks = []) {
    const paragraphs = buildParagraphModelsForInterval(cpStart, cpEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors, inlineContext, resolveSectionIndex, lists, bookmarks);
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
        cpStart: paragraph.cpStart,
        cpEnd: paragraph.cpEnd,
        sectionIndex: paragraph.sectionIndex,
        styleId: paragraph.styleId,
        styleName: paragraph.styleName,
        paraState: paragraph.paraState,
        markStyle: paragraph.markStyle,
        list: paragraph.list,
        bookmarkStarts: paragraph.bookmarkStarts,
        bookmarkEnds: paragraph.bookmarkEnds,
        inlines: paragraph.inlines,
        text: paragraph.text,
    };
}
function cloneTableStateValue(value) {
    if (Array.isArray(value))
        return value.map((item) => cloneTableStateValue(item));
    if (!value || typeof value !== 'object')
        return value;
    const out = {};
    for (const [key, entry] of Object.entries(value))
        out[key] = cloneTableStateValue(entry);
    return out;
}
function cloneTableState(state) {
    const base = tablePropsToState([]);
    if (!state)
        return base;
    return {
        ...base,
        ...cloneTableStateValue(state),
        operations: [...(state.operations || [])],
    };
}
function isEmptyParagraphBlock(block) {
    if (!block.text && !(block.inlines || []).length)
        return true;
    return !String(block.text || '').trim() && !(block.inlines || []).some((inline) => inline.type !== 'text' || String(inline.text || '').trim());
}
function cellHasRenderableContent(cell) {
    if (!cell)
        return false;
    return cell.paragraphs.some((block) => !isEmptyParagraphBlock(block));
}
function createEmptyCellBlock() {
    return {
        id: uniqueId('cell'),
        paragraphs: [],
        meta: null,
    };
}
function tableStateScore(state) {
    if (!state)
        return -1;
    const cellCount = applyTableStateToCells(state).length;
    let score = cellCount * 10;
    if (state.defTable?.cells?.length)
        score += 1000;
    if (state.tableWidth?.wWidth)
        score += 25;
    score += state.operations?.length || 0;
    return score;
}
function buildEffectiveTableState(templateState, rowState) {
    const effective = cloneTableState(templateState);
    if (!rowState)
        return effective;
    if (rowState.styleId != null)
        effective.styleId = rowState.styleId;
    if (rowState.tableWidth)
        effective.tableWidth = cloneTableStateValue(rowState.tableWidth);
    if (rowState.widthBefore != null)
        effective.widthBefore = cloneTableStateValue(rowState.widthBefore);
    if (rowState.widthAfter != null)
        effective.widthAfter = cloneTableStateValue(rowState.widthAfter);
    if (rowState.cellSpacing)
        effective.cellSpacing = cloneTableStateValue(rowState.cellSpacing);
    if (rowState.defTable)
        effective.defTable = cloneTableStateValue(rowState.defTable);
    if (rowState.leftIndent)
        effective.leftIndent = rowState.leftIndent;
    if (rowState.gapHalf)
        effective.gapHalf = rowState.gapHalf;
    if (rowState.rowHeight)
        effective.rowHeight = rowState.rowHeight;
    if (rowState.absLeft != null)
        effective.absLeft = rowState.absLeft;
    if (rowState.absTop != null)
        effective.absTop = rowState.absTop;
    if (rowState.distanceLeft != null)
        effective.distanceLeft = rowState.distanceLeft;
    if (rowState.distanceTop != null)
        effective.distanceTop = rowState.distanceTop;
    if (rowState.positionCode != null)
        effective.positionCode = rowState.positionCode;
    if (rowState.autoFit != null)
        effective.autoFit = cloneTableStateValue(rowState.autoFit);
    effective.alignment = rowState.alignment;
    effective.cantSplit = rowState.cantSplit || effective.cantSplit;
    effective.header = rowState.header || effective.header;
    effective.rtl = rowState.rtl || effective.rtl;
    effective.operations = [...(templateState?.operations || []), ...(rowState.operations || [])];
    return effective;
}
function inferExpectedColumnCount(rawRows, explicitCellCount) {
    if (explicitCellCount > 0)
        return explicitCellCount;
    let best = 0;
    for (const row of rawRows) {
        let count = row.cells.length;
        while (count > 0 && !cellHasRenderableContent(row.cells[count - 1]))
            count -= 1;
        best = Math.max(best, count || row.cells.length);
    }
    return Math.max(best, 1);
}
function inferColumnMeta(rawRows, expectedColumnCount, tableWidthTwips) {
    const safeCount = Math.max(expectedColumnCount, 1);
    const weights = new Array(safeCount).fill(1);
    for (const row of rawRows) {
        const cells = row.cells.slice(0, safeCount);
        for (let index = 0; index < cells.length; index += 1) {
            const textLength = cells[index].paragraphs.map((block) => block.text.trim().length).join('').length;
            if (textLength > 0)
                weights[index] = Math.max(weights[index], Math.min(textLength, 32));
        }
    }
    const totalWeight = weights.reduce((sum, value) => sum + value, 0) || safeCount;
    const width = tableWidthTwips > 0 ? tableWidthTwips : safeCount * 1440;
    let cursor = 0;
    return weights.map((weight, index) => {
        const nextCursor = index === safeCount - 1 ? width : cursor + Math.round((width * weight) / totalWeight);
        const meta = {
            index,
            width: nextCursor - cursor,
            leftBoundary: cursor,
            rightBoundary: nextCursor,
            borders: {},
        };
        cursor = nextCursor;
        return meta;
    });
}
function applyEdgeVerticalMergeHints(rows) {
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex];
        const contentIndices = row.cells.map((cell, index) => (cellHasRenderableContent(cell) ? index : -1)).filter((index) => index >= 0);
        if (!contentIndices.length || contentIndices.length === row.cells.length)
            continue;
        const firstContent = contentIndices[0];
        const lastContent = contentIndices[contentIndices.length - 1];
        const candidates = [
            ...Array.from({ length: firstContent }, (_, index) => index),
            ...Array.from({ length: Math.max(0, row.cells.length - lastContent - 1) }, (_, index) => lastContent + 1 + index),
        ];
        for (const columnIndex of candidates) {
            const cell = row.cells[columnIndex];
            if (!cell || cellHasRenderableContent(cell))
                continue;
            const above = rows[rowIndex - 1].cells[columnIndex];
            if (!above)
                continue;
            if ((above.meta?.vertMerge || 0) === 1) {
                cell.meta = { ...(cell.meta || { index: columnIndex }), vertMerge: 1 };
                continue;
            }
            if (!cellHasRenderableContent(above))
                continue;
            above.meta = { ...(above.meta || { index: columnIndex }), vertMerge: Math.max(above.meta?.vertMerge || 0, 2) };
            cell.meta = { ...(cell.meta || { index: columnIndex }), vertMerge: 1 };
        }
    }
}
function isTableCandidateParagraph(paragraph) {
    return paragraph.terminator === DOC_CONTROL.cellMark
        || paragraph.paraState.inTable
        || paragraph.paraState.innerTableCell
        || paragraph.paraState.tableRowEnd
        || paragraph.paraState.innerTableRowEnd;
}
function buildTableBlock(tableParagraphs) {
    const rawRows = [];
    let pendingRowCells = [];
    let pendingRowParagraphs = [];
    let pendingCellParagraphs = [];
    const flushCell = () => {
        if (!pendingCellParagraphs.length)
            return;
        pendingRowCells.push({
            id: uniqueId('cell'),
            paragraphs: pendingCellParagraphs.map(paragraphToBlock),
            meta: null,
        });
        pendingCellParagraphs = [];
    };
    const flushRow = (rowEndParagraph) => {
        if (pendingCellParagraphs.length)
            flushCell();
        if (!pendingRowCells.length && !pendingRowParagraphs.length)
            return;
        rawRows.push({
            cells: pendingRowCells,
            rowEndParagraph,
            paragraphs: pendingRowParagraphs,
        });
        pendingRowCells = [];
        pendingRowParagraphs = [];
    };
    for (const paragraph of tableParagraphs) {
        pendingRowParagraphs.push(paragraph);
        pendingCellParagraphs.push(paragraph);
        const endsCell = paragraph.terminator === DOC_CONTROL.cellMark || paragraph.paraState.tableRowEnd || paragraph.paraState.innerTableRowEnd;
        const endsRow = paragraph.paraState.tableRowEnd || paragraph.paraState.innerTableRowEnd;
        if (endsCell)
            flushCell();
        if (endsRow)
            flushRow(paragraph);
    }
    flushRow(null);
    const templateState = rawRows.reduce((best, row) => {
        const candidate = row.rowEndParagraph?.tableState || null;
        return tableStateScore(candidate) >= tableStateScore(best) ? candidate : best;
    }, null);
    const templateCellMeta = templateState ? applyTableStateToCells(templateState) : [];
    const expectedColumnCount = inferExpectedColumnCount(rawRows, templateCellMeta.length);
    const inferredCellMeta = inferColumnMeta(rawRows, expectedColumnCount, templateState?.tableWidth?.wWidth || 0);
    const rows = rawRows.map((rawRow) => {
        const rowState = buildEffectiveTableState(templateState, rawRow.rowEndParagraph?.tableState || null);
        let cells = rawRow.cells.map((cell) => ({ ...cell, meta: cell.meta ? cloneTableStateValue(cell.meta) : null }));
        while (cells.length > expectedColumnCount && !cellHasRenderableContent(cells[cells.length - 1]))
            cells.pop();
        while (cells.length < expectedColumnCount)
            cells.push(createEmptyCellBlock());
        const rowCellMeta = applyTableStateToCells(rowState);
        const cellMetaSource = rowCellMeta.length ? rowCellMeta : templateCellMeta.length ? templateCellMeta : inferredCellMeta;
        cells.forEach((cell, index) => {
            cell.meta = cloneTableStateValue(cellMetaSource[index] || inferredCellMeta[index] || { index });
        });
        const gridWidthTwips = cellMetaSource.length
            ? Math.max(0, (cellMetaSource[cellMetaSource.length - 1].rightBoundary || 0) - (cellMetaSource[0].leftBoundary || 0))
            : (rowState.tableWidth?.wWidth || 0);
        return {
            id: uniqueId('row'),
            cells,
            state: rowState,
            gridWidthTwips,
        };
    });
    applyEdgeVerticalMergeHints(rows);
    finalizeTableGrid(rows);
    const firstParagraph = tableParagraphs[0];
    const lastParagraph = tableParagraphs[tableParagraphs.length - 1];
    const gridWidthTwips = rows.find((row) => row.gridWidthTwips)?.gridWidthTwips || templateState?.tableWidth?.wWidth || 0;
    const depthCandidates = rawRows.map((row) => getTableDepth(row.rowEndParagraph?.paraState || tableParagraphs[0].paraState)).filter((value) => value > 0);
    const depth = depthCandidates.length ? Math.max(...depthCandidates) : 1;
    return {
        type: 'table',
        id: uniqueId('table'),
        cpStart: firstParagraph?.cpStart || 0,
        cpEnd: lastParagraph?.cpEnd || firstParagraph?.cpEnd || 0,
        sectionIndex: firstParagraph?.sectionIndex,
        depth,
        rows,
        state: rows[0]?.state || templateState || tablePropsToState([]),
        gridWidthTwips,
    };
}
function buildBlocks(paragraphs) {
    const blocks = [];
    let index = 0;
    while (index < paragraphs.length) {
        const paragraph = paragraphs[index];
        if (!isTableCandidateParagraph(paragraph)) {
            blocks.push(paragraphToBlock(paragraph));
            index += 1;
            continue;
        }
        const tableParagraphs = [];
        while (index < paragraphs.length && isTableCandidateParagraph(paragraphs[index])) {
            tableParagraphs.push(paragraphs[index]);
            index += 1;
        }
        if (tableParagraphs.some((item) => item.terminator === DOC_CONTROL.cellMark || item.paraState.tableRowEnd || item.paraState.innerTableRowEnd)) {
            blocks.push(buildTableBlock(tableParagraphs));
        }
        else {
            for (const item of tableParagraphs)
                blocks.push(paragraphToBlock(item));
        }
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
function buildTextboxItems(header, storyCpBase, entries, parseContent, shapeById = new Map()) {
    const items = [];
    for (const entry of entries) {
        const meta = parseTextboxMeta(entry.data);
        const content = entry.cpEnd > entry.cpStart
            ? parseContent(storyCpBase + entry.cpStart, storyCpBase + entry.cpEnd)
            : { paragraphs: [], blocks: [], text: '' };
        if (!content.blocks.length && !content.text && meta.reusable)
            continue;
        const item = {
            id: uniqueId(header ? 'hdr-textbox' : 'textbox'),
            index: entry.index,
            label: `${header ? 'Header textbox' : 'Textbox'} ${entry.index + 1}`,
            header,
            reusable: meta.reusable,
            shapeId: meta.shapeId || undefined,
            shape: meta.shapeId ? shapeById.get(meta.shapeId) : undefined,
            sectionIndex: meta.shapeId ? shapeById.get(meta.shapeId)?.sectionIndex : undefined,
            blocks: content.blocks,
            text: content.text,
        };
        if (item.shape)
            item.shape.matchedTextboxId = item.id;
        items.push(item);
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
    const sections = readSections(wordBytes, tableBytes, fib.fibRgFcLcb, storyWindows.main.length);
    const resolveMainSectionIndex = (cp) => findSectionIndex(sections, cp);
    const styles = parseStyles(tableBytes, fib.fibRgFcLcb);
    const fonts = parseFonts(tableBytes, fib.fibRgFcLcb);
    const documentProperties = parseDop(tableBytes, fib.fibRgFcLcb, fib);
    const lists = parseLists(tableBytes, fib.fibRgFcLcb);
    const bookmarks = parseBookmarks(tableBytes, fib.fibRgFcLcb);
    const bookmarksByName = new Map();
    for (const bookmark of bookmarks) {
        bookmarksByName.set(bookmark.name, bookmark);
        bookmarksByName.set(bookmark.name.toLowerCase(), bookmark);
    }
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
    const drawingGroup = parseDrawingGroup(tableBytes, wordBytes, fib.fibRgFcLcb, dataBytes);
    warnings.push(...drawingGroup.warnings);
    for (const asset of drawingGroup.assets)
        assets.push(asset);
    const usedAttachmentNames = new Set();
    const assetCache = new Map();
    const resolveAsset = createAssetResolver(dataBytes, objectPool, assets, usedAttachmentNames, assetCache, options);
    const parseContent = (cpIntervalStart, cpIntervalEnd, inlineContext = {}) => {
        const resolveSectionIndex = cpIntervalStart >= storyWindows.main.cpStart && cpIntervalEnd <= storyWindows.main.cpEnd
            ? resolveMainSectionIndex
            : () => 0;
        return parseIntervalToContent(cpIntervalStart, cpIntervalEnd, documentText, wordBytes, clx, pieceTexts, styles, fonts, papxRuns, chpxRuns, resolveAsset, revisionAuthors, { ...inlineContext, bookmarksByName, sequenceCounters: inlineContext.sequenceCounters || new Map() }, resolveSectionIndex, lists, bookmarks);
    };
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
    const headerStoryDescriptors = buildHeaderStoryDescriptors(tableBytes, fib.fibRgFcLcb, storyWindows.header);
    const mainShapeAnchors = readShapeAnchors(tableBytes, fib.fibRgFcLcb, storyWindows.main.cpStart, 'main')
        .map((anchor) => ({ ...anchor, sectionIndex: resolveMainSectionIndex(anchor.anchorCp) }))
        .map((anchor) => applyDrawingInfoToShapeAnchor(anchor, drawingGroup.shapes));
    const headerShapeAnchors = readShapeAnchors(tableBytes, fib.fibRgFcLcb, storyWindows.header.cpStart, 'header')
        .map((anchor) => {
        const binding = resolveHeaderAnchorBinding(headerStoryDescriptors, anchor.anchorCp);
        return {
            ...applyDrawingInfoToShapeAnchor(anchor, drawingGroup.shapes),
            sectionIndex: binding.sectionIndex ?? anchor.sectionIndex,
            headerKind: binding.kind,
            headerRole: binding.role,
        };
    });
    const mainShapeById = new Map(mainShapeAnchors.map((anchor) => [anchor.shapeId, anchor]));
    const headerShapeById = new Map(headerShapeAnchors.map((anchor) => [anchor.shapeId, anchor]));
    const headerStories = buildHeaderStories(headerStoryDescriptors, parseContent);
    const textboxItems = buildTextboxItems(false, storyWindows.textbox.cpStart, readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcftxbxTxt, fib.fibRgFcLcb.lcbPlcftxbxTxt, 22), parseContent, mainShapeById);
    const headerTextboxItems = buildTextboxItems(true, storyWindows.headerTextbox.cpStart, readFixedPlc(tableBytes, fib.fibRgFcLcb.fcPlcfHdrtxbxTxt, fib.fibRgFcLcb.lcbPlcfHdrtxbxTxt, 22), parseContent, headerShapeById);
    const floatingShapes = mainShapeAnchors.filter((anchor) => !anchor.matchedTextboxId);
    const headerFloatingShapes = headerShapeAnchors.filter((anchor) => !anchor.matchedTextboxId);
    const unresolvedFloatingShapes = floatingShapes.filter((anchor) => !anchor.imageAsset);
    const unresolvedHeaderFloatingShapes = headerFloatingShapes.filter((anchor) => !anchor.imageAsset);
    if (unresolvedFloatingShapes.length || unresolvedHeaderFloatingShapes.length) {
        pushWarning(warnings, 'Some floating shapes were parsed but still require metadata-card fallback because neither a textbox story nor a resolvable image BLIP was available', {
            code: 'floating-shapes-partial-render',
            severity: 'info',
            details: { mainShapes: unresolvedFloatingShapes.length, headerShapes: unresolvedHeaderFloatingShapes.length },
        });
    }
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
    if (floatingShapes.length) {
        const shapesBlock = { type: 'shapes', id: uniqueId('shapes'), header: false, items: floatingShapes };
        blocks.push(shapesBlock);
    }
    if (headerFloatingShapes.length) {
        const headerShapesBlock = { type: 'shapes', id: uniqueId('header-shapes'), header: true, items: headerFloatingShapes };
        blocks.push(headerShapesBlock);
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
                shapes: mainShapeAnchors.length,
                headerShapes: headerShapeAnchors.length,
                sections: sections.length,
                lists: lists.definitions.length,
                bookmarks: bookmarks.length,
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
        sections,
        documentProperties,
        lists,
        bookmarks,
        blocks,
        assets,
    };
}
//# sourceMappingURL=parser.js.map