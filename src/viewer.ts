import { parseMsDoc } from './msdoc/parser.js';
import { renderMsDoc } from './render/html.js';
import type {
  MsDocParseToHtmlOptions,
  MsDocRenderResult,
  MsDocViewer,
  MsDocViewerConfig,
  MsDocViewerLoadOptions,
  ViewerInput,
} from './types.js';

async function normalizeInput(input: ViewerInput): Promise<ArrayBuffer> {
  if (input instanceof ArrayBuffer) return input;
  if (ArrayBuffer.isView(input)) {
    return input.byteOffset === 0 && input.byteLength === input.buffer.byteLength
      ? input.buffer.slice(0)
      : input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
  }
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    return await input.arrayBuffer();
  }
  if (typeof input === 'string') {
    const response = await fetch(input);
    if (!response.ok) throw new Error(`Failed to fetch document: ${response.status}`);
    return await response.arrayBuffer();
  }
  throw new TypeError('Unsupported input type');
}

export function mountMsDoc(container: HTMLElement, rendered: MsDocRenderResult): HTMLElement {
  if (!container) throw new Error('A container element is required');
  container.innerHTML = `<style data-msdoc>${rendered.css}</style><div class="msdoc-root">${rendered.html}</div>`;
  return container;
}

export async function parseMsDocToHtml(input: ViewerInput, options: MsDocParseToHtmlOptions = {}): Promise<MsDocRenderResult> {
  const buffer = await normalizeInput(input);
  if (options.workerClient) {
    return options.workerClient.parseToHtml(buffer, {
      parseOptions: options.parseOptions || {},
      renderOptions: options.renderOptions || {},
    });
  }
  const parsed = parseMsDoc(buffer, options.parseOptions || {});
  return renderMsDoc(parsed, options.renderOptions || {});
}

/**
 * Small DOM-oriented helper that keeps browser integration trivial.
 * Apps can either use it directly or consume the lower-level parse/render APIs.
 */
export function createMsDocViewer(container: HTMLElement, config: MsDocViewerConfig = {}): MsDocViewer {
  let current: MsDocRenderResult | null = null;
  return {
    async load(input: ViewerInput, options: MsDocViewerLoadOptions = {}): Promise<MsDocRenderResult> {
      const rendered = await parseMsDocToHtml(input, {
        workerClient: options.workerClient || config.workerClient,
        parseOptions: { ...(config.parseOptions || {}), ...(options.parseOptions || {}) },
        renderOptions: { ...(config.renderOptions || {}), ...(options.renderOptions || {}) },
      });
      mountMsDoc(container, rendered);
      current = rendered;
      return rendered;
    },
    mount(rendered: MsDocRenderResult): HTMLElement {
      current = rendered;
      return mountMsDoc(container, rendered);
    },
    clear(): void {
      container.innerHTML = '';
      current = null;
    },
    destroy(): void {
      this.clear();
    },
    get value(): MsDocRenderResult | null {
      return current;
    },
  };
}
