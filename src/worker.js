import { parseMsDoc } from './msdoc/parser.js';
import { renderMsDoc } from './render/html.js';

self.addEventListener('message', (event) => {
  const { id, type, buffer, parsed, options } = event.data || {};
  try {
    let result;
    if (type === 'parse') {
      result = parseMsDoc(buffer, options || {});
    } else if (type === 'render') {
      result = renderMsDoc(parsed, options || {});
    } else if (type === 'parseToHtml') {
      const parseOptions = options?.parseOptions || options || {};
      const renderOptions = options?.renderOptions || {};
      result = renderMsDoc(parseMsDoc(buffer, parseOptions), renderOptions);
    } else {
      throw new Error(`Unsupported worker request type: ${type}`);
    }
    self.postMessage({ id, ok: true, result });
  } catch (error) {
    self.postMessage({ id, ok: false, error: error?.message || String(error) });
  }
});
