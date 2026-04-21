import assert from 'assert/strict';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { parseMsDoc, renderMsDoc } from '../dist/index.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const sampleCandidates = [
  path.resolve(__dirname, './test.doc'),
  path.resolve(__dirname, './fixtures/image-embedded.doc'),
  path.resolve(__dirname, './fixtures/image-linked.doc'),
  path.resolve(__dirname, '../../work/file-viewer3/public/example/test.doc'),
  path.resolve(__dirname, '../../work/file-viewer3/test/test.doc'),
].filter((file) => fs.existsSync(file));

if (!sampleCandidates.length) {
  throw new Error('No sample .doc file found for smoke test');
}

const sampleResults = sampleCandidates.map((file) => {
  const buffer = fs.readFileSync(file);
  const parsed = parseMsDoc(buffer);
  const rendered = renderMsDoc(parsed);
  return { file, parsed, rendered };
});

const defaultSample = sampleResults[0];
if (!defaultSample) throw new Error('Missing default sample result');

const imageFixture = sampleResults.find((item) => item.file.endsWith(path.join('fixtures', 'image-embedded.doc')));
if (!imageFixture) {
  throw new Error('Missing embedded image fixture for smoke test');
}

const imageAssets = imageFixture.parsed.assets.filter((asset) => asset.type === 'image');
if (imageAssets.length < 2) {
  throw new Error(`Expected at least two images in embedded image fixture, got ${imageAssets.length}`);
}
if (!imageAssets.every((asset) => asset.displayable !== false && /^(image\/png|image\/jpeg)$/i.test(asset.mime))) {
  throw new Error(`Embedded image fixture did not resolve to browser-displayable raster assets: ${imageAssets.map((asset) => asset.mime).join(', ')}`);
}

const linkedFixture = sampleResults.find((item) => item.file.endsWith(path.join('fixtures', 'image-linked.doc')));
if (!linkedFixture) {
  throw new Error('Missing linked image fixture for smoke test');
}
const linkedImages = linkedFixture.parsed.assets.filter((asset) => asset.type === 'image');
assert.ok(linkedImages.length >= 2, 'Expected linked image fixture to expose linked image assets');
assert.ok(linkedImages.every((asset) => asset.displayable === false), 'Local external linked images should not be marked as displayable');
assert.ok(linkedImages.every((asset) => asset.meta?.localExternal === true), 'Local external linked images should be tagged in asset metadata');
assert.ok(!/href="file:\/\//i.test(linkedFixture.rendered.html), 'Rendered HTML must not emit clickable file:// links for local external images');

const securityRendered = renderMsDoc({
  kind: 'msdoc',
  version: 1,
  warnings: [],
  meta: {
    fib: {
      wIdent: 0xA5EC,
      nFib: 0,
      fWhichTblStm: 0,
      fComplex: false,
      fEncrypted: false,
      ccpText: 0,
    },
    counts: {
      paragraphs: 1,
      blocks: 1,
      assets: 2,
      styles: 0,
      fonts: 0,
    },
  },
  fonts: [],
  styles: [],
  assets: [
    {
      id: 'asset-img-unsafe',
      type: 'image',
      mime: 'image/png',
      bytes: new Uint8Array(0),
      dataUrl: '',
      sourceUrl: 'file:///tmp/unsafe.png',
      displayable: false,
      meta: { pictureOffset: -1, lcb: 0, cbHeader: 0, sourceKind: 'linked', localExternal: true },
    },
    {
      id: 'asset-attachment-1',
      type: 'attachment',
      name: 'embedded.bin',
      mime: 'application/octet-stream',
      bytes: new Uint8Array(0),
      dataUrl: 'data:application/octet-stream;base64,AA==',
    },
  ],
  blocks: [
    {
      type: 'paragraph',
      id: 'p1',
      styleId: 0,
      styleName: '',
      text: 'unsafe',
      paraState: {
        styleId: 0,
        alignment: 0,
        spacingBefore: 0,
        spacingAfter: 0,
        lineSpacing: 0,
        leftIndent: 0,
        rightIndent: 0,
        firstLineIndent: 0,
        keepLines: false,
        keepNext: false,
        pageBreakBefore: false,
        widowControl: false,
        inTable: false,
        tableRowEnd: false,
        innerTableCell: false,
        innerTableRowEnd: false,
        itap: 0,
        dtap: 0,
        rtlPara: false,
        adjustRight: false,
        borders: {},
      },
      inlines: [
        {
          type: 'text',
          text: 'bad link',
          href: 'javascript:alert(1)',
          style: {
            bold: false,
            italic: false,
            strike: false,
            underline: 0,
            spacing: 0,
            positionHalfPoints: 0,
            scale: 100,
            hidden: false,
            smallCaps: false,
            caps: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            rtl: false,
            rtlChar: false,
            vanish: false,
            revision: false,
            data: false,
            object: false,
            ole2: false,
            symbol: false,
            fieldVanish: false,
            boldBi: false,
            italicBi: false,
            doubleStrike: false,
          },
        },
        {
          type: 'image',
          asset: {
            id: 'unsafe-image',
            type: 'image',
            mime: 'image/png',
            bytes: new Uint8Array(0),
            dataUrl: '',
            sourceUrl: 'file:///tmp/local.png',
            displayable: false,
            meta: { pictureOffset: -1, lcb: 0, cbHeader: 0, sourceKind: 'linked', localExternal: true },
          },
          href: 'javascript:alert(2)',
          style: {
            bold: false,
            italic: false,
            strike: false,
            underline: 0,
            spacing: 0,
            positionHalfPoints: 0,
            scale: 100,
            hidden: false,
            smallCaps: false,
            caps: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            rtl: false,
            rtlChar: false,
            vanish: false,
            revision: false,
            data: false,
            object: false,
            ole2: false,
            symbol: false,
            fieldVanish: false,
            boldBi: false,
            italicBi: false,
            doubleStrike: false,
          },
        },
        {
          type: 'attachment',
          asset: {
            id: 'attachment-inline',
            type: 'attachment',
            name: 'inline.bin',
            mime: 'application/octet-stream',
            bytes: new Uint8Array(0),
            dataUrl: 'data:application/octet-stream;base64,AA==',
          },
          href: 'https://example.com/spec',
          style: {
            bold: false,
            italic: false,
            strike: false,
            underline: 0,
            spacing: 0,
            positionHalfPoints: 0,
            scale: 100,
            hidden: false,
            smallCaps: false,
            caps: false,
            outline: false,
            shadow: false,
            emboss: false,
            imprint: false,
            rtl: false,
            rtlChar: false,
            vanish: false,
            revision: false,
            data: false,
            object: false,
            ole2: false,
            symbol: false,
            fieldVanish: false,
            boldBi: false,
            italicBi: false,
            doubleStrike: false,
          },
        },
      ],
    },
  ],
});
assert.ok(!/javascript:alert\(/i.test(securityRendered.html), 'Rendered HTML must strip javascript: links');
assert.ok(!/href="file:\/\//i.test(securityRendered.html), 'Rendered HTML must not emit clickable local file URLs');
assert.ok(!/<a[^>]*>\s*<a/i.test(securityRendered.html), 'Rendered HTML must not contain nested anchors');

console.log('sample:', defaultSample.file);
console.log('counts:', defaultSample.parsed.meta.counts);
console.log('warnings:', defaultSample.parsed.warnings.length);
console.log('image fixture:', imageFixture.file);
console.log('image assets:', imageAssets.map((asset) => ({ mime: asset.mime, displayable: asset.displayable })));
console.log('linked image warnings:', linkedFixture.parsed.warnings.map((warning) => warning.code || warning.message));

fs.writeFileSync(
  new URL('./rendered-sample.html', import.meta.url),
  `<!doctype html><meta charset="utf-8"><style>${defaultSample.rendered.css}</style><div class="msdoc-root">${defaultSample.rendered.html}</div>`
);
fs.writeFileSync(
  new URL('./rendered-image-sample.html', import.meta.url),
  `<!doctype html><meta charset="utf-8"><style>${imageFixture.rendered.css}</style><div class="msdoc-root">${imageFixture.rendered.html}</div>`
);
