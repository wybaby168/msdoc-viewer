import fs from 'fs';
import path from 'path';
import { parseMsDoc, renderMsDoc } from '../src/index.js';

const sampleCandidates = [
  path.resolve('../work/file-viewer3/public/example/test.doc'),
  path.resolve('../work/file-viewer3/test/test.doc'),
  path.resolve('./test.doc'),
].filter((file) => fs.existsSync(file));

if (!sampleCandidates.length) {
  throw new Error('No sample .doc file found for smoke test');
}

const buffer = fs.readFileSync(sampleCandidates[0]);
const parsed = parseMsDoc(buffer);
const rendered = renderMsDoc(parsed);

console.log('counts:', parsed.meta.counts);
console.log('warnings:', parsed.warnings.length);

fs.writeFileSync(
  new URL('./rendered-sample.html', import.meta.url),
  `<!doctype html><meta charset="utf-8"><style>${rendered.css}</style><div class="msdoc-root">${rendered.html}</div>`
);
