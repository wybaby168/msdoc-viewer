import { cleanTextControlChars, escapeHtml, slugify, twipsToPx } from '../core/utils.js';
import { HIGHLIGHT_COLORS } from '../msdoc/constants.js';
import { cssTextAlign, cssUnderline, cssVerticalAlign } from '../msdoc/properties.js';
const COLOR_INDEX_MAP = {
    1: '#000000',
    2: '#0000ff',
    3: '#00ffff',
    4: '#00ff00',
    5: '#ff00ff',
    6: '#ff0000',
    7: '#ffff00',
    8: '#ffffff',
    9: '#000080',
    10: '#008080',
    11: '#008000',
    12: '#800080',
    13: '#800000',
    14: '#808000',
    15: '#808080',
    16: '#c0c0c0',
};
function styleObjectToCss(style) {
    return Object.entries(style)
        .filter(([, value]) => value != null && value !== '')
        .map(([key, value]) => `${key}:${value}`)
        .join(';');
}
// Inline pictures use PICMID goal dimensions instead of intrinsic bitmap size.
// This matters most for header/footer images, where Word often stores a large
// raster but asks the renderer to display it much smaller on the page.
function inlineImageDisplaySizePx(asset) {
    const widthPx = twipsToPx(asset.meta?.displayWidthTwips);
    const heightPx = twipsToPx(asset.meta?.displayHeightTwips);
    return {
        widthPx: widthPx && widthPx > 0 ? widthPx : undefined,
        heightPx: heightPx && heightPx > 0 ? heightPx : undefined,
    };
}
function applyInlineImageDisplaySize(style, asset) {
    const { widthPx, heightPx } = inlineImageDisplaySizePx(asset);
    if (!widthPx && !heightPx)
        return;
    if (widthPx)
        style.width = `${widthPx}px`;
    if (heightPx)
        style.height = `${heightPx}px`;
}
function sanitizeLinkHref(href) {
    const value = String(href || '').trim();
    if (!value)
        return null;
    if (/^(?:https?:|mailto:|tel:|ftp:)/i.test(value))
        return value;
    if (value.startsWith('#'))
        return value;
    return null;
}
function sanitizeImageSource(src) {
    const value = String(src || '').trim();
    if (!value)
        return null;
    if (/^data:image\//i.test(value))
        return value;
    if (/^(?:https?:|blob:)/i.test(value))
        return value;
    return null;
}
function sanitizeAssetHref(href) {
    const value = String(href || '').trim();
    if (!value)
        return null;
    if (/^(?:data:|blob:|https?:)/i.test(value))
        return value;
    return null;
}
function renderExternalRef(href) {
    if (!href)
        return '';
    return `<a class="msdoc-link msdoc-external-ref" href="${escapeHtml(href)}" target="_blank" rel="noreferrer noopener">↗</a>`;
}
function joinWithExternalRef(content, href) {
    if (!href)
        return content;
    return `<span class="msdoc-inline-group">${content}${renderExternalRef(href)}</span>`;
}
function normalizeRenderableText(text) {
    return cleanTextControlChars(text)
        .replace(/\t/g, ' ')
        .replace(/[\u0000-\u0008\u000e-\u001f\u007f]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}
function normalizeDisplayText(text) {
    return cleanTextControlChars(text)
        .replace(/\t/g, '    ')
        .replace(/[\u0000-\u0008\u000e-\u001f\u007f]/g, '');
}
function paragraphHasRenderableContent(block) {
    if (normalizeRenderableText(block.text))
        return true;
    return (block.inlines || []).some((node) => {
        if (node.type === 'text')
            return Boolean(normalizeRenderableText(node.text));
        if (node.type === 'lineBreak' || node.type === 'pageBreak')
            return false;
        return true;
    });
}
function blockListHasRenderableContent(blocks) {
    return blocks.some((block) => {
        if (block.type === 'table')
            return block.rows.some((row) => row.cells.some((cell) => cell.paragraphs.some(paragraphHasRenderableContent)));
        return paragraphHasRenderableContent(block);
    });
}
function blockListText(blocks) {
    return normalizeRenderableText(blocks.map((block) => {
        if (block.type === 'table') {
            return block.rows
                .map((row) => row.cells.filter((cell) => !cell.hidden).map((cell) => cell.paragraphs.map((paragraph) => paragraph.text).join(' ')).join(' | '))
                .join(' || ');
        }
        return block.text;
    }).join(' '));
}
function isPageNumberLike(text) {
    const normalized = normalizeRenderableText(text);
    if (!normalized)
        return false;
    return /^[-–—\s\d]+$/.test(normalized) || /\b(?:PAGE|NUMPAGES|SECTIONPAGES)\b/i.test(String(text || ''));
}
function signatureForBlocks(blocks, text) {
    const imageCount = blocks.reduce((count, block) => count + (block.type === 'paragraph' ? block.inlines.filter((node) => node.type === 'image').length : 0), 0);
    const tableCount = blocks.filter((block) => block.type === 'table').length;
    return `${normalizeRenderableText(text).replace(/\d+/g, '#')}|img:${imageCount}|tbl:${tableCount}`;
}
function mergeChromeEntry(target, entry) {
    const signature = entry.source === 'shape'
        ? `${entry.kind}|shape:${entry.shape?.shapeId || entry.text || entry.labels.join('|')}`
        : `${entry.kind}|${signatureForBlocks(entry.blocks, entry.text)}`;
    const current = target.get(signature);
    if (!current) {
        target.set(signature, {
            ...entry,
            labels: [...entry.labels],
            sections: [...entry.sections],
        });
        return;
    }
    for (const label of entry.labels) {
        if (!current.labels.includes(label))
            current.labels.push(label);
    }
    for (const section of entry.sections) {
        if (!current.sections.includes(section))
            current.sections.push(section);
    }
    current.pageNumberLike = current.pageNumberLike || entry.pageNumberLike;
    if (!current.html && entry.html)
        current.html = entry.html;
    if (!current.shape && entry.shape)
        current.shape = entry.shape;
}
function storyKind(role) {
    if (role.endsWith('Header'))
        return 'header';
    if (role.endsWith('Footer'))
        return 'footer';
    return null;
}
function headerShapeKind(shape) {
    if (shape.headerKind)
        return shape.headerKind;
    if (shape.headerRole)
        return shape.headerRole.endsWith('Footer') ? 'footer' : 'header';
    return shape.boundsTwips.top > 7920 ? 'footer' : 'header';
}
function getImageAsset(assetById, shape) {
    if (!shape?.imageAssetId)
        return null;
    return assetById.get(shape.imageAssetId) || null;
}
function collectChromeEntries(headersBlock, headerTextboxes, headerShapes, assetById) {
    const map = new Map();
    const usedShapeIds = new Set();
    for (const story of headersBlock?.stories || []) {
        const kind = storyKind(story.role);
        if (!kind)
            continue;
        if (!blockListHasRenderableContent(story.blocks))
            continue;
        const text = blockListText(story.blocks) || story.text;
        mergeChromeEntry(map, {
            kind,
            source: 'story',
            blocks: story.blocks,
            text,
            labels: [story.roleLabel],
            sections: story.sectionIndex != null ? [story.sectionIndex] : [],
            pageNumberLike: isPageNumberLike(text),
        });
    }
    for (const item of headerTextboxes) {
        if (!blockListHasRenderableContent(item.blocks))
            continue;
        const text = blockListText(item.blocks) || item.text;
        const kind = item.shape?.headerKind
            || (item.shape?.headerRole ? (item.shape.headerRole.endsWith('Footer') ? 'footer' : 'header') : undefined)
            || (isPageNumberLike(item.text || text) ? 'footer' : 'header');
        mergeChromeEntry(map, {
            kind,
            source: 'textbox',
            blocks: item.blocks,
            text,
            labels: [item.label],
            sections: item.sectionIndex != null ? [item.sectionIndex + 1] : [],
            pageNumberLike: isPageNumberLike(item.text || text),
            shape: item.shape,
        });
    }
    for (const shape of headerShapes) {
        const asset = getImageAsset(assetById, shape);
        if (!asset)
            continue;
        const kind = headerShapeKind(shape);
        usedShapeIds.add(shape.id);
        mergeChromeEntry(map, {
            kind,
            source: 'shape',
            blocks: [],
            text: `${shape.drawingName || shape.drawingDescription || 'shape'} ${shape.shapeId}`,
            labels: [shape.drawingName || shape.drawingDescription || `shape ${shape.shapeId}`],
            sections: shape.sectionIndex != null ? [shape.sectionIndex + 1] : [],
            pageNumberLike: false,
            shape,
            html: renderAssetImage(asset, shape, 'msdoc-page-chrome-asset'),
        });
    }
    const values = [...map.values()];
    return {
        headers: values.filter((entry) => entry.kind === 'header'),
        footers: values.filter((entry) => entry.kind === 'footer'),
        usedShapeIds,
    };
}
function borderToCss(border) {
    if (!border)
        return null;
    const raw = border.raw;
    const borderType = border.borderType;
    const lineWidth = border.lineWidth;
    if (raw === 0xffffffff || borderType == null || borderType === 0xff || lineWidth == null || lineWidth === 0xff || borderType === 0) {
        return null;
    }
    const width = Math.max(1, Math.min(12, Math.round((lineWidth / 8) * 1.3333)));
    const style = borderType === 6 ? 'double' : borderType === 3 ? 'dotted' : borderType === 2 ? 'dashed' : 'solid';
    const colorIndex = border.color;
    const color = colorIndex != null && COLOR_INDEX_MAP[colorIndex] ? COLOR_INDEX_MAP[colorIndex] : '#666';
    return `${width}px ${style} ${color}`;
}
function paragraphStyleToCss(paraState) {
    const style = {
        'text-align': cssTextAlign(paraState.alignment),
    };
    const marginTop = twipsToPx(paraState.spacingBefore);
    const marginBottom = twipsToPx(paraState.spacingAfter);
    const marginLeft = twipsToPx(paraState.leftIndent);
    const marginRight = twipsToPx(paraState.rightIndent);
    const textIndent = twipsToPx(paraState.firstLineIndent);
    if (marginTop)
        style['margin-top'] = `${marginTop}px`;
    if (marginBottom)
        style['margin-bottom'] = `${marginBottom}px`;
    if (marginLeft)
        style['margin-left'] = `${marginLeft}px`;
    if (marginRight)
        style['margin-right'] = `${marginRight}px`;
    if (textIndent)
        style['text-indent'] = `${textIndent}px`;
    if (paraState.lineSpacing) {
        if (paraState.lineSpacingRule === 'multiple' || paraState.lineSpacingMultiple) {
            const lineHeight = Math.abs(paraState.lineSpacing) / 240;
            if (lineHeight)
                style['line-height'] = String(Math.max(1, lineHeight));
        }
        else {
            const lineHeightPx = twipsToPx(paraState.lineSpacingTwips || Math.abs(paraState.lineSpacing));
            if (lineHeightPx)
                style['line-height'] = `${Math.max(1, lineHeightPx)}px`;
        }
    }
    if (paraState.keepLines)
        style['break-inside'] = 'avoid';
    if (paraState.keepNext)
        style['break-after'] = 'avoid';
    if (paraState.pageBreakBefore)
        style['break-before'] = 'page';
    if (paraState.rtlPara)
        style.direction = 'rtl';
    const top = borderToCss(paraState.borders?.top);
    const right = borderToCss(paraState.borders?.right);
    const bottom = borderToCss(paraState.borders?.bottom);
    const left = borderToCss(paraState.borders?.left);
    if (top)
        style['border-top'] = top;
    if (right)
        style['border-right'] = right;
    if (bottom)
        style['border-bottom'] = bottom;
    if (left)
        style['border-left'] = left;
    return style;
}
function buildUnderlineStyle(underline) {
    const wordStyle = cssUnderline(underline);
    if (!underline || wordStyle === 'none')
        return {};
    const css = { 'text-decoration-line': 'underline' };
    if (wordStyle === 'double' || wordStyle === 'wavy-double')
        css['text-decoration-style'] = 'double';
    else if (wordStyle.includes('dot') || wordStyle === 'dotted-heavy')
        css['text-decoration-style'] = 'dotted';
    else if (wordStyle.includes('dash'))
        css['text-decoration-style'] = 'dashed';
    else if (wordStyle.includes('wave') || wordStyle.includes('wavy'))
        css['text-decoration-style'] = 'wavy';
    else
        css['text-decoration-style'] = 'solid';
    return css;
}
function inlineStyleToCss(styleState) {
    const style = {};
    if (styleState.bold || styleState.boldBi)
        style['font-weight'] = '700';
    if (styleState.italic || styleState.italicBi)
        style['font-style'] = 'italic';
    if (styleState.strike || styleState.doubleStrike)
        style['text-decoration-line'] = `${style['text-decoration-line'] ? `${style['text-decoration-line']} ` : ''}line-through`;
    Object.assign(style, buildUnderlineStyle(styleState.underline));
    if (styleState.fontSizeHalfPoints)
        style['font-size'] = `${styleState.fontSizeHalfPoints / 2}pt`;
    if (styleState.fontFamily)
        style['font-family'] = `'${String(styleState.fontFamily).replace(/'/g, "\\'")}', sans-serif`;
    if (styleState.colorIndex && COLOR_INDEX_MAP[styleState.colorIndex])
        style.color = COLOR_INDEX_MAP[styleState.colorIndex];
    const highlightIndex = typeof styleState.highlight === 'number' ? styleState.highlight : styleState.highlight?.index;
    if (highlightIndex && HIGHLIGHT_COLORS[highlightIndex]) {
        style['background-color'] = HIGHLIGHT_COLORS[highlightIndex];
    }
    if (styleState.smallCaps)
        style['font-variant-caps'] = 'small-caps';
    if (styleState.caps)
        style['text-transform'] = 'uppercase';
    if (styleState.scale && styleState.scale !== 100) {
        style.display = 'inline-block';
        style.transform = `scaleX(${styleState.scale / 100})`;
        style['transform-origin'] = 'left center';
    }
    if (styleState.positionHalfPoints > 0)
        style['vertical-align'] = 'super';
    if (styleState.positionHalfPoints < 0)
        style['vertical-align'] = 'sub';
    if (styleState.outline)
        style['text-shadow'] = '0 0 0.02em currentColor';
    if (styleState.shadow || styleState.emboss || styleState.imprint) {
        style['text-shadow'] = style['text-shadow'] ? `${style['text-shadow']}, 0.06em 0.06em 0.08em rgba(0,0,0,.25)` : '0.06em 0.06em 0.08em rgba(0,0,0,.25)';
    }
    if (styleState.revisionInsert) {
        style['background-color'] = style['background-color'] || 'rgba(22,163,74,.10)';
        style['text-decoration-color'] = '#16a34a';
    }
    if (styleState.revisionDelete) {
        style.color = style.color || '#b91c1c';
        style['background-color'] = style['background-color'] || 'rgba(185,28,28,.08)';
    }
    if (styleState.rtl)
        style.direction = 'rtl';
    return style;
}
function paragraphMarkStyleToCss(styleState) {
    if (!styleState)
        return {};
    const style = {};
    if (styleState.fontSizeHalfPoints)
        style['font-size'] = `${styleState.fontSizeHalfPoints / 2}pt`;
    if (styleState.fontFamily)
        style['font-family'] = `'${String(styleState.fontFamily).replace(/'/g, "\'")}', sans-serif`;
    if (styleState.bold || styleState.boldBi)
        style['font-weight'] = '700';
    if (styleState.italic || styleState.italicBi)
        style['font-style'] = 'italic';
    if (styleState.rtl)
        style.direction = 'rtl';
    return style;
}
function revisionTitle(styleState) {
    const parts = [];
    if (styleState.revisionInsert)
        parts.push('Inserted text');
    if (styleState.revisionDelete)
        parts.push('Deleted text');
    if (styleState.revisionAuthor)
        parts.push(`by ${styleState.revisionAuthor}`);
    return parts.join(' ');
}
function wrapRevision(inner, styleState) {
    const title = revisionTitle(styleState);
    const titleAttr = title ? ` title="${escapeHtml(title)}"` : '';
    if (styleState.revisionDelete) {
        return `<del class="msdoc-revision msdoc-revision-delete"${titleAttr}>${inner}</del>`;
    }
    if (styleState.revisionInsert) {
        return `<ins class="msdoc-revision msdoc-revision-insert"${titleAttr}>${inner}</ins>`;
    }
    return inner;
}
function renderTextNode(node) {
    const displayText = normalizeDisplayText(node.text);
    if (!displayText)
        return '';
    const content = escapeHtml(displayText);
    const inlineStyle = inlineStyleToCss(node.style);
    inlineStyle['white-space'] = 'break-spaces';
    const style = styleObjectToCss(inlineStyle);
    const span = `<span${style ? ` style="${style}"` : ''}>${content}</span>`;
    const inner = wrapRevision(span, node.style);
    const href = sanitizeLinkHref(node.href);
    if (href) {
        return `<a class="msdoc-link" href="${escapeHtml(href)}" target="_blank" rel="noreferrer noopener">${inner}</a>`;
    }
    return inner;
}
function renderImageNode(node) {
    const src = sanitizeImageSource(node.asset.sourceUrl) || sanitizeImageSource(node.asset.dataUrl);
    const baseStyle = inlineStyleToCss(node.style);
    baseStyle['max-width'] = '100%';
    baseStyle.height = 'auto';
    applyInlineImageDisplaySize(baseStyle, node.asset);
    const href = sanitizeLinkHref(node.href);
    if (!src || node.asset.displayable === false) {
        const fallbackHref = sanitizeAssetHref(node.asset.dataUrl) || sanitizeAssetHref(node.asset.sourceUrl);
        const label = escapeHtml(String(node.asset.meta?.linkedPath || node.asset.mime || 'image'));
        const inner = fallbackHref
            ? `<a class="msdoc-attachment msdoc-image-fallback" href="${escapeHtml(fallbackHref)}" target="_blank" rel="noreferrer noopener">🖼 ${label}</a>`
            : `<span class="msdoc-image-fallback">🖼 ${label}</span>`;
        return joinWithExternalRef(inner, href);
    }
    const img = `<img class="msdoc-image" src="${escapeHtml(src)}" alt="" style="${styleObjectToCss(baseStyle)}">`;
    if (href) {
        return `<a class="msdoc-link" href="${escapeHtml(href)}" target="_blank" rel="noreferrer noopener">${img}</a>`;
    }
    return img;
}
function renderStandaloneImageAsset(asset) {
    const src = sanitizeImageSource(asset.sourceUrl) || sanitizeImageSource(asset.dataUrl);
    if (!src || asset.displayable === false) {
        const fallbackHref = sanitizeAssetHref(asset.dataUrl) || sanitizeAssetHref(asset.sourceUrl);
        const label = escapeHtml(String(asset.meta?.linkedPath || asset.mime || 'image'));
        return fallbackHref
            ? `<a class="msdoc-attachment msdoc-image-fallback" href="${escapeHtml(fallbackHref)}" target="_blank" rel="noreferrer noopener">🖼 ${label}</a>`
            : `<span class="msdoc-image-fallback">🖼 ${label}</span>`;
    }
    const style = { display: 'block', 'max-width': '100%', height: 'auto', margin: '0 auto' };
    applyInlineImageDisplaySize(style, asset);
    return `<img class="msdoc-image" src="${escapeHtml(src)}" alt="" style="${styleObjectToCss(style)}">`;
}
function renderAttachmentNode(node) {
    const label = escapeHtml(node.asset.name || 'embedded-file');
    const inner = `<a class="msdoc-attachment" href="${escapeHtml(node.asset.dataUrl)}" download="${label}">📎 ${label}</a>`;
    return joinWithExternalRef(inner, sanitizeLinkHref(node.href));
}
function renderNoteRefNode(node) {
    const href = `#${escapeHtml(node.refId)}`;
    return `<sup class="msdoc-note-ref msdoc-note-ref-${node.noteType}"><a class="msdoc-link" href="${href}">${escapeHtml(node.label)}</a></sup>`;
}
function renderCommentRefNode(node) {
    const href = `#${escapeHtml(node.commentId)}`;
    const title = node.author ? ` title="${escapeHtml(node.author)}"` : '';
    return `<sup class="msdoc-comment-ref"${title}><a class="msdoc-link" href="${href}">💬${escapeHtml(node.label)}</a></sup>`;
}
function renderFieldNode(node) {
    const text = escapeHtml(normalizeDisplayText(node.displayText || ''));
    const style = styleObjectToCss(inlineStyleToCss(node.style));
    const classes = [`msdoc-field`, `msdoc-field-${slugify(node.fieldType)}`];
    const attrs = [
        ` data-field-type="${escapeHtml(node.fieldType)}"`,
        node.instruction ? ` data-field-instruction="${escapeHtml(node.instruction)}"` : '',
        node.target ? ` data-field-target="${escapeHtml(node.target)}"` : '',
    ].join('');
    const span = `<span class="${classes.join(' ')}"${attrs}${style ? ` style="${style}"` : ''}>${text}</span>`;
    const href = sanitizeLinkHref(node.href);
    if (href)
        return `<a class="msdoc-link msdoc-field-link" href="${escapeHtml(href)}">${span}</a>`;
    return span;
}
function renderBookmarkAnchors(block) {
    return (block.bookmarkStarts || [])
        .map((bookmark) => `<span id="${escapeHtml(bookmark.id)}" class="msdoc-bookmark-anchor" data-bookmark-name="${escapeHtml(bookmark.name)}"></span>`)
        .join('');
}
function renderListLabel(block) {
    if (!block.list)
        return '';
    const label = block.list.label || '';
    const level = Math.max(0, block.list.level || 0);
    const gap = block.list.follow === 'space' ? '&nbsp;' : block.list.follow === 'none' ? '' : '&nbsp;&nbsp;';
    return `<span class="msdoc-list-label msdoc-list-label-level-${level}" aria-hidden="true">${escapeHtml(label)}${gap}</span>`;
}
function renderInlineNodes(nodes) {
    return nodes.map((node) => {
        if (node.type === 'text')
            return renderTextNode(node);
        if (node.type === 'image')
            return renderImageNode(node);
        if (node.type === 'attachment')
            return renderAttachmentNode(node);
        if (node.type === 'noteRef')
            return renderNoteRefNode(node);
        if (node.type === 'commentRef')
            return renderCommentRefNode(node);
        if (node.type === 'field')
            return renderFieldNode(node);
        if (node.type === 'lineBreak')
            return '<br>';
        if (node.type === 'pageBreak')
            return '<span class="msdoc-page-break"></span>';
        return '';
    }).join('');
}
function renderParagraphBlock(block, options = {}) {
    const tag = options.inline ? 'div' : 'p';
    const paraStyle = paragraphStyleToCss(block.paraState);
    if (options.tableCell) {
        delete paraStyle['text-indent'];
        paraStyle['margin-top'] = '0';
        paraStyle['margin-bottom'] = '0';
    }
    const inlineBody = renderInlineNodes(block.inlines || []);
    if (!inlineBody)
        Object.assign(paraStyle, paragraphMarkStyleToCss(block.markStyle));
    const style = styleObjectToCss(paraStyle);
    const bookmarkAnchors = renderBookmarkAnchors(block);
    const listLabel = renderListLabel(block);
    const body = `${bookmarkAnchors}${listLabel}${inlineBody || '<br>'}`;
    const classNames = ['msdoc-paragraph'];
    if (!options.inline)
        classNames.push('msdoc-flow-block');
    if (block.list)
        classNames.push('msdoc-list-paragraph');
    if (block.styleName)
        classNames.push(`msdoc-style-${slugify(block.styleName)}`);
    const attrs = [];
    if (!options.inline) {
        attrs.push(` data-section-index="${block.sectionIndex ?? 0}"`);
        attrs.push(` data-cp-start="${block.cpStart}"`);
        attrs.push(` data-cp-end="${block.cpEnd}"`);
        if (block.paraState.keepNext)
            attrs.push(' data-keep-next="1"');
        if (block.styleName)
            attrs.push(` data-style-name="${escapeHtml(block.styleName)}"`);
        if (block.list) {
            attrs.push(` data-list-id="${block.list.listId}"`);
            attrs.push(` data-list-level="${block.list.level}"`);
            attrs.push(` data-list-label="${escapeHtml(block.list.label)}"`);
        }
        if (block.bookmarkStarts?.length)
            attrs.push(` data-bookmark-starts="${escapeHtml(block.bookmarkStarts.map((bookmark) => bookmark.name).join(','))}"`);
        const headingLike = /标题|题|Title|Heading/i.test(block.styleName || '')
            || (Boolean(normalizeRenderableText(block.text)) && normalizeRenderableText(block.text).length <= 40 && block.paraState.alignment === 1 && (block.markStyle?.bold || block.inlines.some((node) => node.type === 'text' && (node.style.bold || node.style.boldBi))));
        if (headingLike)
            attrs.push(' data-heading-like="1"');
        if (block.paraState.pageBreakBefore)
            attrs.push(' data-page-break-before="1"');
    }
    return `<${tag} class="${classNames.join(' ')}"${style ? ` style="${style}"` : ''}${attrs.join('')}>${body}</${tag}>`;
}
function cellStyle(cell) {
    const style = {};
    const widthTwips = (cell.meta?.rightBoundary != null && cell.meta?.leftBoundary != null)
        ? cell.meta.rightBoundary - cell.meta.leftBoundary
        : cell.meta?.width;
    const widthPx = twipsToPx(widthTwips);
    if (widthPx)
        style.width = `${widthPx}px`;
    if (cell.meta?.noWrap)
        style['white-space'] = 'nowrap';
    if (cell.meta?.fitText)
        style['text-align'] = 'justify';
    if (cell.meta?.vertAlign != null)
        style['vertical-align'] = cssVerticalAlign(cell.meta.vertAlign);
    const borderAll = borderToCss(cell.meta?.borders?.all);
    const top = borderToCss(cell.meta?.borders?.top) || borderAll;
    const right = borderToCss(cell.meta?.borders?.right) || borderAll;
    const bottom = borderToCss(cell.meta?.borders?.bottom) || borderAll;
    const left = borderToCss(cell.meta?.borders?.left) || borderAll;
    if (top)
        style['border-top'] = top;
    if (right)
        style['border-right'] = right;
    if (bottom)
        style['border-bottom'] = bottom;
    if (left)
        style['border-left'] = left;
    const shadingColor = typeof cell.meta?.shading === 'object' && cell.meta?.shading && 'color' in cell.meta.shading
        ? cell.meta.shading.color
        : undefined;
    if (shadingColor && COLOR_INDEX_MAP[shadingColor]) {
        style['background-color'] = COLOR_INDEX_MAP[shadingColor];
    }
    if (cell.meta?.padding) {
        const topPadding = twipsToPx(cell.meta.padding.top);
        const rightPadding = twipsToPx(cell.meta.padding.right);
        const bottomPadding = twipsToPx(cell.meta.padding.bottom);
        const leftPadding = twipsToPx(cell.meta.padding.left);
        if (topPadding != null)
            style['padding-top'] = `${Math.max(topPadding, 0)}px`;
        if (rightPadding != null)
            style['padding-right'] = `${Math.max(rightPadding, 0)}px`;
        if (bottomPadding != null)
            style['padding-bottom'] = `${Math.max(bottomPadding, 0)}px`;
        if (leftPadding != null)
            style['padding-left'] = `${Math.max(leftPadding, 0)}px`;
    }
    return style;
}
function tableStyle(block) {
    const style = {};
    const widthPx = twipsToPx(block.gridWidthTwips || block.state?.tableWidth?.wWidth);
    if (widthPx)
        style.width = `${widthPx}px`;
    else
        style.width = '100%';
    const marginLeft = twipsToPx(block.state?.leftIndent);
    if ((block.state?.alignment || 0) === 1) {
        style['margin-left'] = 'auto';
        style['margin-right'] = 'auto';
    }
    else if ((block.state?.alignment || 0) === 2) {
        style['margin-left'] = 'auto';
        style['margin-right'] = '0';
    }
    else if (marginLeft != null) {
        style['margin-left'] = `${marginLeft}px`;
    }
    const spacingPx = twipsToPx(block.state?.cellSpacing?.wWidth || block.state?.cellSpacing?.width);
    if (spacingPx) {
        style['border-collapse'] = 'separate';
        style['border-spacing'] = `${Math.max(spacingPx, 0)}px`;
    }
    else {
        style['border-collapse'] = 'collapse';
        style['border-spacing'] = '0';
    }
    style['table-layout'] = 'fixed';
    return style;
}
function renderTableBlock(block) {
    const rows = block.rows.map((row) => {
        const rowHeight = row.state?.rowHeight ? twipsToPx(Math.abs(row.state.rowHeight)) : null;
        const rowStyle = rowHeight ? ` style="height:${rowHeight}px"` : '';
        const cells = row.cells
            .filter((cell) => !cell.hidden)
            .map((cell) => {
            const attrs = [];
            if ((cell.colspan ?? 1) > 1)
                attrs.push(` colspan="${cell.colspan}"`);
            if ((cell.rowspan ?? 1) > 1)
                attrs.push(` rowspan="${cell.rowspan}"`);
            const style = styleObjectToCss(cellStyle(cell));
            const body = cell.paragraphs.map((paragraph) => renderParagraphBlock(paragraph, { inline: true, tableCell: true })).join('');
            return `<td class="msdoc-cell"${attrs.join('')}${style ? ` style="${style}"` : ''}>${body || '<div class="msdoc-paragraph"><br></div>'}</td>`;
        })
            .join('');
        return `<tr class="msdoc-row"${rowStyle}>${cells}</tr>`;
    }).join('');
    return `<table class="msdoc-table msdoc-flow-block msdoc-table-depth-${block.depth}" style="${styleObjectToCss(tableStyle(block))}" data-section-index="${block.sectionIndex ?? 0}" data-cp-start="${block.cpStart}" data-cp-end="${block.cpEnd}"><tbody>${rows}</tbody></table>`;
}
export function renderBlockList(blocks) {
    return blocks.map((block) => block.type === 'paragraph' ? renderParagraphBlock(block) : renderTableBlock(block)).join('');
}
function renderNotesBlock(block) {
    const title = block.kind === 'footnote' ? 'Footnotes' : 'Endnotes';
    const items = block.items.map((item) => `
    <li id="${escapeHtml(item.id)}" class="msdoc-note-item">
      <div class="msdoc-note-label">${escapeHtml(item.label)}</div>
      <div class="msdoc-note-body">${renderBlockList(item.blocks)}</div>
    </li>
  `).join('');
    return `<section class="msdoc-section msdoc-notes msdoc-notes-${block.kind}"><div class="msdoc-section-title">${title}</div><ol class="msdoc-note-list">${items}</ol></section>`;
}
function renderCommentsBlock(block) {
    const items = block.items.map((item) => {
        const meta = [item.author, item.initials].filter(Boolean).join(' · ');
        return `
      <li id="${escapeHtml(item.id)}" class="msdoc-comment-item">
        <div class="msdoc-comment-header">
          <span class="msdoc-comment-label">Comment ${escapeHtml(item.label)}</span>
          ${meta ? `<span class="msdoc-comment-meta">${escapeHtml(meta)}</span>` : ''}
        </div>
        <div class="msdoc-comment-body">${renderBlockList(item.blocks)}</div>
      </li>
    `;
    }).join('');
    return `<section class="msdoc-section msdoc-comments"><div class="msdoc-section-title">Comments</div><ol class="msdoc-comment-list">${items}</ol></section>`;
}
function renderAttachmentsBlock(block) {
    const items = block.items.map((item) => `<li><a class="msdoc-attachment" href="${escapeHtml(item.dataUrl)}" download="${escapeHtml(item.name || 'embedded-file')}">📎 ${escapeHtml(item.name || 'embedded-file')}</a></li>`).join('');
    return `<section class="msdoc-attachments"><div class="msdoc-attachments-title">Embedded attachments</div><ul>${items}</ul></section>`;
}
function formatPxValue(value) {
    return value == null ? 'auto' : `${value}px`;
}
function renderShapeMeta(shape) {
    const widthPx = twipsToPx(shape.boundsTwips.width);
    const heightPx = twipsToPx(shape.boundsTwips.height);
    const leftPx = twipsToPx(shape.boundsTwips.left);
    const topPx = twipsToPx(shape.boundsTwips.top);
    const badges = [
        shape.behindText ? 'behind text' : 'in front of text',
        shape.anchorLocked ? 'anchor locked' : '',
        shape.matchedTextboxId ? 'textbox linked' : '',
        shape.imageAssetId ? 'image bound' : '',
    ].filter(Boolean).map((label) => `<span class="msdoc-badge">${escapeHtml(label)}</span>`).join(' ');
    const rows = [
        ['Shape ID', String(shape.shapeId)],
        ['Anchor CP', String(shape.anchorCp)],
        ['Origin', `${shape.anchorX} / ${shape.anchorY}`],
        ['Wrap', `${shape.wrapStyle} / ${shape.wrapSide}`],
        ['Bounds', `${formatPxValue(leftPx)} × ${formatPxValue(topPx)} → ${formatPxValue(widthPx)} × ${formatPxValue(heightPx)}`],
        shape.shapeTypeCode != null ? ['Type', String(shape.shapeTypeCode)] : null,
        shape.headerRole ? ['Role', shape.headerRole] : null,
        shape.drawingName ? ['Name', shape.drawingName] : null,
        shape.drawingDescription ? ['Description', shape.drawingDescription] : null,
        shape.blipRef ? ['BLIP', `${shape.blipRef.kind} #${shape.blipRef.index}`] : null,
    ].filter((row) => Boolean(row)).map(([label, value]) => `<div class="msdoc-shape-meta-row"><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd></div>`).join('');
    return `${badges ? `<div class="msdoc-shape-badges">${badges}</div>` : ''}<dl class="msdoc-shape-meta">${rows}</dl>`;
}
function floatingStyle(shape) {
    const style = {
        width: `${Math.max(twipsToPx(shape.boundsTwips.width) || 0, 120)}px`,
        'max-width': '100%',
    };
    const height = twipsToPx(shape.boundsTwips.height);
    if (height)
        style['min-height'] = `${Math.max(height, 48)}px`;
    if (shape.wrapStyle === 'topBottom' || shape.wrapStyle === 'none') {
        style.display = 'block';
        style.margin = '12px auto';
    }
    else if (shape.wrapSide === 'left') {
        style.float = 'left';
        style.margin = '4px 16px 12px 0';
    }
    else {
        style.float = 'right';
        style.margin = '4px 0 12px 16px';
    }
    if (shape.behindText)
        style.opacity = '0.92';
    return style;
}
function isImageOnlyParagraph(block) {
    const nonImageContent = (block.inlines || []).filter((node) => {
        if (node.type === 'image')
            return false;
        if (node.type === 'text')
            return Boolean(normalizeRenderableText(node.text));
        return node.type !== 'lineBreak' && node.type !== 'pageBreak';
    });
    return block.inlines.some((node) => node.type === 'image') && nonImageContent.length === 0;
}
function renderFloatingImage(shape, block) {
    return `<figure class="msdoc-floating msdoc-floating-image msdoc-flow-block" style="${styleObjectToCss(floatingStyle(shape))}" data-msdoc-floating="1" data-shape-id="${shape.shapeId}" data-section-index="${shape.sectionIndex ?? block.sectionIndex ?? 0}" data-anchor-cp="${shape.anchorCp}">${renderInlineNodes(block.inlines)}</figure>`;
}
function renderAssetImage(asset, shape, extraClass = '') {
    const src = sanitizeImageSource(asset.sourceUrl) || sanitizeImageSource(asset.dataUrl);
    if (!src || asset.displayable === false) {
        const fallbackHref = sanitizeAssetHref(asset.dataUrl) || sanitizeAssetHref(asset.sourceUrl);
        const label = escapeHtml(String(asset.meta?.fbseName || asset.meta?.linkedPath || asset.mime || 'image'));
        const inner = fallbackHref
            ? `<a class="msdoc-attachment msdoc-image-fallback" href="${escapeHtml(fallbackHref)}" target="_blank" rel="noreferrer noopener">🖼 ${label}</a>`
            : `<span class="msdoc-image-fallback">🖼 ${label}</span>`;
        return `<figure class="msdoc-floating msdoc-floating-image ${extraClass}" style="${styleObjectToCss(floatingStyle(shape))}" data-msdoc-floating="1" data-shape-id="${shape.shapeId}" data-section-index="${shape.sectionIndex ?? 0}" data-anchor-cp="${shape.anchorCp}">${inner}</figure>`;
    }
    const img = `<img class="msdoc-image" src="${escapeHtml(src)}" alt="${escapeHtml(shape.drawingDescription || shape.drawingName || '')}" style="max-width:100%;height:auto;display:block;margin:0 auto">`;
    return `<figure class="msdoc-floating msdoc-floating-image ${extraClass}" style="${styleObjectToCss(floatingStyle(shape))}" data-msdoc-floating="1" data-shape-id="${shape.shapeId}" data-section-index="${shape.sectionIndex ?? 0}" data-anchor-cp="${shape.anchorCp}">${img}</figure>`;
}
function renderFloatingTextbox(shape, item, asset) {
    const preview = asset ? `<div class="msdoc-story-card-extra">${renderAssetImage(asset, shape)}</div>` : '';
    return `<aside class="msdoc-floating msdoc-floating-textbox msdoc-textboxes msdoc-flow-block" style="${styleObjectToCss(floatingStyle(shape))}" data-msdoc-floating="1" data-shape-id="${shape.shapeId}" data-section-index="${shape.sectionIndex ?? item.sectionIndex ?? 0}" data-anchor-cp="${shape.anchorCp}"><div class="msdoc-floating-title">Textbox</div><div class="msdoc-story-card-meta">textbox linked</div>${preview}<div class="msdoc-floating-body">${renderBlockList(item.blocks)}</div></aside>`;
}
function renderFloatingShapePlaceholder(shape, assetById) {
    const asset = shape.imageAssetId ? assetById.get(shape.imageAssetId) : undefined;
    if (asset) {
        return renderAssetImage(asset, shape, 'msdoc-flow-block');
    }
    return `<aside class="msdoc-floating msdoc-floating-shape-placeholder msdoc-shapes msdoc-flow-block" style="${styleObjectToCss(floatingStyle(shape))}" data-msdoc-floating="1" data-shape-id="${shape.shapeId}" data-section-index="${shape.sectionIndex ?? 0}" data-anchor-cp="${shape.anchorCp}"><div class="msdoc-floating-title">Floating shape</div>${renderShapeMeta(shape)}</aside>`;
}
function findAnchorIndex(blocks, anchorCp) {
    const index = blocks.findIndex((block) => block.cpEnd >= anchorCp);
    return index >= 0 ? index : Math.max(0, blocks.length - 1);
}
function findNearestImageParagraphIndex(blocks, anchorCp, usedIds, imageAssetId) {
    let bestIndex = null;
    let bestScore = Number.POSITIVE_INFINITY;
    const anchorIndex = findAnchorIndex(blocks, anchorCp);
    for (let index = 0; index < blocks.length; index += 1) {
        const block = blocks[index];
        if (block.type !== 'paragraph' || usedIds.has(block.id) || !isImageOnlyParagraph(block))
            continue;
        if (imageAssetId) {
            const imageIds = block.inlines.filter((node) => node.type === 'image').map((node) => node.asset.id);
            if (!imageIds.includes(imageAssetId))
                continue;
        }
        const cpDistance = anchorCp < block.cpStart
            ? block.cpStart - anchorCp
            : anchorCp > block.cpEnd
                ? anchorCp - block.cpEnd
                : 0;
        const indexDistance = Math.abs(index - anchorIndex);
        if (cpDistance > 2048 && indexDistance > 8)
            continue;
        const score = cpDistance + indexDistance * 32;
        if (score < bestScore) {
            bestScore = score;
            bestIndex = index;
        }
    }
    return bestIndex;
}
function renderFloatingPlaceholders(title, items) {
    if (!items.length)
        return '';
    return `<section class="msdoc-section msdoc-floating-leftovers"><div class="msdoc-section-title">${escapeHtml(title)}</div><div class="msdoc-floating-grid">${items.join('')}</div></section>`;
}
function placeFloatingArtifacts(mainBlocks, mainTextboxes, mainShapes, assetById) {
    const beforeMap = new Map();
    const skipBlockIds = new Set();
    const leftoverHtml = [];
    const remainingTextboxes = new Map();
    const textboxByShapeId = new Map();
    for (const item of mainTextboxes) {
        remainingTextboxes.set(item.id, item);
        if (item.shapeId != null)
            textboxByShapeId.set(item.shapeId, item);
    }
    const pushBefore = (index, html) => {
        const list = beforeMap.get(index) || [];
        list.push(html);
        beforeMap.set(index, list);
    };
    for (const shape of mainShapes) {
        const insertionIndex = findAnchorIndex(mainBlocks, shape.anchorCp);
        const textbox = (shape.matchedTextboxId ? remainingTextboxes.get(shape.matchedTextboxId) : null)
            || textboxByShapeId.get(shape.shapeId)
            || null;
        const boundAsset = shape.imageAssetId ? assetById.get(shape.imageAssetId) : undefined;
        if (textbox && blockListHasRenderableContent(textbox.blocks)) {
            pushBefore(insertionIndex, renderFloatingTextbox(shape, textbox, boundAsset));
            remainingTextboxes.delete(textbox.id);
            continue;
        }
        if (boundAsset) {
            pushBefore(insertionIndex, renderAssetImage(boundAsset, shape, 'msdoc-flow-block'));
            continue;
        }
        if (shape.imageAssetId) {
            const exactImageIndex = findNearestImageParagraphIndex(mainBlocks, shape.anchorCp, skipBlockIds, shape.imageAssetId);
            if (exactImageIndex != null) {
                const imageBlock = mainBlocks[exactImageIndex];
                if (imageBlock && imageBlock.type === 'paragraph') {
                    skipBlockIds.add(imageBlock.id);
                    pushBefore(insertionIndex, renderFloatingImage(shape, imageBlock));
                    continue;
                }
            }
            const asset = assetById.get(shape.imageAssetId);
            if (asset) {
                pushBefore(insertionIndex, renderAssetImage(asset, shape, 'msdoc-flow-block'));
                continue;
            }
        }
        const imageIndex = findNearestImageParagraphIndex(mainBlocks, shape.anchorCp, skipBlockIds);
        if (imageIndex != null) {
            const imageBlock = mainBlocks[imageIndex];
            if (imageBlock && imageBlock.type === 'paragraph') {
                skipBlockIds.add(imageBlock.id);
                pushBefore(insertionIndex, renderFloatingImage(shape, imageBlock));
                continue;
            }
        }
        leftoverHtml.push(renderFloatingShapePlaceholder(shape, assetById));
    }
    for (const item of remainingTextboxes.values()) {
        if (!blockListHasRenderableContent(item.blocks))
            continue;
        const meta = item.shape ? '<div class="msdoc-story-card-meta">textbox linked</div>' : '';
        leftoverHtml.push(`<article class="msdoc-story-card msdoc-textboxes"><div class="msdoc-story-card-title">${escapeHtml(item.label)}</div>${meta}<div class="msdoc-story-card-body">${renderBlockList(item.blocks)}</div></article>`);
    }
    return { beforeMap, skipBlockIds, leftoverHtml };
}
function renderMainContent(mainBlocks, mainTextboxes, mainShapes, assetById) {
    const placement = placeFloatingArtifacts(mainBlocks, mainTextboxes, mainShapes, assetById);
    const parts = [];
    for (let index = 0; index < mainBlocks.length; index += 1) {
        const before = placement.beforeMap.get(index);
        if (before?.length)
            parts.push(before.join(''));
        const block = mainBlocks[index];
        if (placement.skipBlockIds.has(block.id))
            continue;
        parts.push(block.type === 'paragraph' ? renderParagraphBlock(block) : renderTableBlock(block));
    }
    const tail = placement.beforeMap.get(mainBlocks.length);
    if (tail?.length)
        parts.push(tail.join(''));
    return { html: parts.join(''), leftoverHtml: placement.leftoverHtml };
}
function renderChromeBand(kind, entries) {
    if (!entries.length)
        return '';
    const items = entries.map((entry) => {
        const metaParts = [];
        if (entry.sections.length)
            metaParts.push(`section ${entry.sections.join(', ')}`);
        if (entry.labels.length)
            metaParts.push(entry.labels.join(' · '));
        if (entry.source === 'textbox')
            metaParts.push('textbox');
        if (entry.source === 'shape')
            metaParts.push('shape');
        const meta = metaParts.join(' · ');
        const classes = ['msdoc-page-chrome-entry', `msdoc-page-chrome-entry-${kind}`];
        if (entry.pageNumberLike)
            classes.push('msdoc-page-chrome-entry-page');
        const body = entry.html || renderBlockList(entry.blocks);
        return `<div class="${classes.join(' ')}">${meta ? `<div class="msdoc-page-chrome-meta">${escapeHtml(meta)}</div>` : ''}<div class="msdoc-page-chrome-body">${body}</div></div>`;
    }).join('');
    return `<section class="msdoc-page-chrome msdoc-page-chrome-${kind} ${kind === 'header' ? 'msdoc-headers' : 'msdoc-footers'}">${items}</section>`;
}
function renderHeadersAppendix(block) {
    const items = block.stories
        .filter((story) => !storyKind(story.role) && blockListHasRenderableContent(story.blocks))
        .map((story) => {
        const section = story.sectionIndex != null ? `Section ${story.sectionIndex}` : 'Shared';
        const inherited = story.inheritedFromSection != null ? `<span class="msdoc-badge">inherits section ${story.inheritedFromSection}</span>` : '';
        return `
        <article class="msdoc-story-card">
          <div class="msdoc-story-card-title">${escapeHtml(story.roleLabel)}</div>
          <div class="msdoc-story-card-meta">${escapeHtml(section)} ${inherited}</div>
          <div class="msdoc-story-card-body">${renderBlockList(story.blocks)}</div>
        </article>
      `;
    })
        .join('');
    if (!items)
        return '';
    return `<section class="msdoc-section msdoc-headers"><div class="msdoc-section-title">Additional header/footer stories</div><div class="msdoc-story-grid">${items}</div></section>`;
}
function renderTextboxesAppendix(title, items, assetById) {
    if (!items.length)
        return '';
    const html = items.map((item) => {
        const meta = [item.reusable ? 'reusable' : '', item.shapeId != null ? `shape ${item.shapeId}` : ''].filter(Boolean).join(' · ');
        const asset = item.shape?.imageAssetId ? assetById.get(item.shape.imageAssetId) : undefined;
        const preview = asset && item.shape ? `<div class="msdoc-story-card-extra">${renderAssetImage(asset, item.shape)}</div>` : '';
        const shapeMeta = item.shape ? `<div class="msdoc-story-card-extra">${renderShapeMeta(item.shape)}</div>` : '';
        return `
      <article class="msdoc-story-card">
        <div class="msdoc-story-card-title">${escapeHtml(item.label)}</div>
        ${meta ? `<div class="msdoc-story-card-meta">${escapeHtml(meta)}</div>` : ''}
        ${preview}
        ${shapeMeta}
        <div class="msdoc-story-card-body">${renderBlockList(item.blocks)}</div>
      </article>
    `;
    }).join('');
    return `<section class="msdoc-section msdoc-textboxes"><div class="msdoc-section-title">${escapeHtml(title)}</div><div class="msdoc-story-grid">${html}</div></section>`;
}
function renderShapesAppendix(title, shapes, assetById) {
    if (!shapes.length)
        return '';
    const items = shapes.map((shape, index) => {
        const asset = shape.imageAssetId ? assetById.get(shape.imageAssetId) : undefined;
        const preview = asset ? `<div class="msdoc-story-card-extra">${renderAssetImage(asset, shape)}</div>` : '';
        return `
      <article class="msdoc-story-card msdoc-shape-card">
        <div class="msdoc-story-card-title">${escapeHtml(`Shape ${index + 1}`)}</div>
        <div class="msdoc-story-card-meta">${escapeHtml(`${shape.story} story`)}</div>
        ${preview}
        <div class="msdoc-story-card-body">${renderShapeMeta(shape)}</div>
      </article>
    `;
    }).join('');
    return `<section class="msdoc-section msdoc-shapes"><div class="msdoc-section-title">${escapeHtml(title)}</div><div class="msdoc-story-grid">${items}</div></section>`;
}
export function defaultMsDocCss() {
    return `
.msdoc-root{box-sizing:border-box;max-width:100%;padding:24px;background:#fff;color:#111;font:14px/1.2 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}
.msdoc-root *{box-sizing:border-box}
.msdoc-body{position:relative}
.msdoc-paragraph{margin:0;white-space:normal;word-break:break-word;overflow-wrap:anywhere}
.msdoc-paragraph:last-child{margin-bottom:0}
.msdoc-list-paragraph{display:block}
.msdoc-list-label{display:inline-block;min-width:2.4em;margin-right:.25em;white-space:pre;text-align:right}
.msdoc-list-label-level-1{min-width:3.2em}.msdoc-list-label-level-2{min-width:4em}.msdoc-list-label-level-3{min-width:4.8em}
.msdoc-bookmark-anchor{position:relative;top:-.25em}
.msdoc-field{background:rgba(96,165,250,.08);border-radius:2px;padding:0 .05em}
.msdoc-field-page,.msdoc-field-numpages{background:transparent;padding:0}
.msdoc-table{margin:0;max-width:100%;border-collapse:collapse;border-spacing:0}
.msdoc-cell{padding:0;vertical-align:top;word-break:break-word;overflow-wrap:anywhere}
.msdoc-link{color:#1a73e8;text-decoration:none}
.msdoc-link:hover{text-decoration:underline}
.msdoc-inline-group{display:inline-flex;align-items:center;gap:6px;vertical-align:middle;max-width:100%}
.msdoc-external-ref{font-size:.9em}
.msdoc-image{display:inline-block;vertical-align:middle}
.msdoc-image-fallback{display:inline-flex;align-items:center;gap:6px}
.msdoc-attachment{display:inline-flex;align-items:center;gap:6px;padding:4px 8px;border:1px solid #d0d7de;border-radius:6px;background:#f6f8fa;color:#0969da;text-decoration:none}
.msdoc-page-break{display:block;height:0;border-top:1px dashed #cbd5e1;margin:16px 0}
.msdoc-page-chrome{margin:0 0 16px;padding:10px 0;border-bottom:1px solid #e5e7eb}
.msdoc-page-chrome-footer{margin-top:24px;margin-bottom:0;border-top:1px solid #e5e7eb;border-bottom:0;padding-top:10px;padding-bottom:0}
.msdoc-page-chrome-entry{position:relative}
.msdoc-page-chrome-entry + .msdoc-page-chrome-entry{margin-top:10px}
.msdoc-page-chrome-entry-page{display:flex;justify-content:flex-end}
.msdoc-page-chrome-meta{font-size:.82em;color:#6b7280;margin-bottom:4px}
.msdoc-page-chrome-body .msdoc-paragraph{margin-bottom:0}
.msdoc-page-chrome-body .msdoc-table{margin:8px 0}
.msdoc-floating{position:relative;border:1px solid #e5e7eb;border-radius:10px;background:#fff;box-shadow:0 2px 8px rgba(15,23,42,.06);padding:8px}
.msdoc-floating::after{content:"";display:block;clear:both}
.msdoc-floating-image{display:block;text-align:center}
.msdoc-floating-image .msdoc-image{display:block;margin:0 auto;max-width:100%;height:auto}
.msdoc-floating-textbox{background:#fafafa}
.msdoc-floating-body .msdoc-paragraph:last-child{margin-bottom:0}
.msdoc-floating-title{font-weight:600;margin-bottom:8px}
.msdoc-floating-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}
.msdoc-attachments{margin-top:20px;padding-top:12px;border-top:1px solid #e5e7eb}
.msdoc-attachments-title{font-weight:600;margin-bottom:8px}
.msdoc-section{margin-top:24px;padding-top:12px;border-top:1px solid #e5e7eb}
.msdoc-section-title{font-weight:700;margin-bottom:12px;font-size:1rem}
.msdoc-note-list,.msdoc-comment-list{margin:0;padding-left:20px}
.msdoc-note-item,.msdoc-comment-item{margin:0 0 12px}
.msdoc-note-item:last-child,.msdoc-comment-item:last-child{margin-bottom:0}
.msdoc-note-label,.msdoc-comment-label{font-weight:600}
.msdoc-note-body,.msdoc-comment-body{margin-top:4px}
.msdoc-comment-header{display:flex;gap:8px;align-items:baseline;flex-wrap:wrap}
.msdoc-comment-meta{color:#6b7280;font-size:.92em}
.msdoc-story-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px}
.msdoc-story-card{border:1px solid #e5e7eb;border-radius:10px;padding:12px;background:#fafafa}
.msdoc-story-card-title{font-weight:600;margin-bottom:4px}
.msdoc-story-card-meta{color:#6b7280;font-size:.92em;margin-bottom:8px}
.msdoc-story-card-extra{margin-bottom:10px}
.msdoc-badge{display:inline-block;padding:2px 6px;border-radius:999px;background:#eef2ff;color:#4338ca;font-size:.82em}
.msdoc-shape-badges{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:8px}
.msdoc-shape-meta{margin:0;display:grid;gap:6px}
.msdoc-shape-meta-row{display:grid;grid-template-columns:88px minmax(0,1fr);gap:8px}
.msdoc-shape-meta-row dt{font-weight:600;margin:0}
.msdoc-shape-meta-row dd{margin:0;color:#374151}
.msdoc-note-ref,.msdoc-comment-ref{font-size:.78em;vertical-align:super;line-height:1}
.msdoc-comment-ref a,.msdoc-note-ref a{text-decoration:none}
.msdoc-revision{border-radius:2px;padding:0 1px}
.msdoc-revision-insert{text-decoration:none}
.msdoc-revision-delete{text-decoration-thickness:1px}
@media (max-width: 768px){
  .msdoc-root{padding:16px}
  .msdoc-floating{float:none !important;width:auto !important;max-width:100%;margin:12px 0 !important}
}
`;
}
/**
 * Converts the parsed AST into HTML and a companion CSS string.
 * Keeping rendering separate from parsing makes it easier for downstream apps
 * to customize styles or consume the AST directly.
 */
export function renderMsDoc(parsed, options = {}) {
    const css = options.css ?? defaultMsDocCss();
    const assetById = new Map(parsed.assets.filter((asset) => asset.type === 'image').map((asset) => [asset.id, asset]));
    const mainBlocks = parsed.blocks.filter((block) => block.type === 'paragraph' || block.type === 'table');
    const attachmentsBlocks = parsed.blocks.filter((block) => block.type === 'attachments');
    const notesBlocks = parsed.blocks.filter((block) => block.type === 'notes');
    const commentsBlocks = parsed.blocks.filter((block) => block.type === 'comments');
    const headersBlock = parsed.blocks.find((block) => block.type === 'headers') || null;
    const mainTextboxes = parsed.blocks.filter((block) => block.type === 'textboxes' && !block.header).flatMap((block) => block.items);
    const headerTextboxes = parsed.blocks.filter((block) => block.type === 'textboxes' && block.header).flatMap((block) => block.items);
    const mainShapes = parsed.blocks.filter((block) => block.type === 'shapes' && !block.header).flatMap((block) => block.items);
    const headerShapes = parsed.blocks.filter((block) => block.type === 'shapes' && block.header).flatMap((block) => block.items);
    const chromeEntries = collectChromeEntries(headersBlock, headerTextboxes, headerShapes, assetById);
    const mainContent = renderMainContent(mainBlocks, mainTextboxes, mainShapes, assetById);
    const html = [
        renderChromeBand('header', chromeEntries.headers),
        `<div class="msdoc-body">${mainContent.html}</div>`,
        renderChromeBand('footer', chromeEntries.footers),
        ...attachmentsBlocks.map(renderAttachmentsBlock),
        ...notesBlocks.map(renderNotesBlock),
        ...commentsBlocks.map(renderCommentsBlock),
        renderFloatingPlaceholders('Unplaced floating content', mainContent.leftoverHtml),
        renderTextboxesAppendix('Additional textboxes', headerTextboxes.filter((item) => !blockListHasRenderableContent(item.blocks) ? false : !isPageNumberLike(item.text || blockListText(item.blocks))), assetById),
        renderShapesAppendix('Additional header shapes', headerShapes.filter((shape) => !chromeEntries.usedShapeIds.has(shape.id)), assetById),
        headersBlock ? renderHeadersAppendix(headersBlock) : '',
    ].join('');
    return {
        html,
        css,
        warnings: parsed.warnings || [],
        meta: parsed.meta,
        assets: parsed.assets || [],
        parsed,
    };
}
//# sourceMappingURL=html.js.map