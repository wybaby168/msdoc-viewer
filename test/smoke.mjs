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

console.log('sample:', defaultSample.file);
console.log('counts:', defaultSample.parsed.meta.counts);
console.log('warnings:', defaultSample.parsed.warnings.length);
console.log('image fixture:', imageFixture.file);
console.log('image assets:', imageAssets.map((asset) => ({ mime: asset.mime, displayable: asset.displayable })));

fs.writeFileSync(
  new URL('./rendered-sample.html', import.meta.url),
  `<!doctype html><meta charset="utf-8"><style>${defaultSample.rendered.css}</style><div class="msdoc-root">${defaultSample.rendered.html}</div>`
);
fs.writeFileSync(
  new URL('./rendered-image-sample.html', import.meta.url),
  `<!doctype html><meta charset="utf-8"><style>${imageFixture.rendered.css}</style><div class="msdoc-root">${imageFixture.rendered.html}</div>`
);
