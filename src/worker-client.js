function toArrayBuffer(input) {
  if (input instanceof ArrayBuffer) return input.slice(0);
  if (ArrayBuffer.isView(input)) {
    return input.byteOffset === 0 && input.byteLength === input.buffer.byteLength
      ? input.buffer.slice(0)
      : input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
  }
  throw new TypeError('Expected ArrayBuffer or TypedArray');
}

export class MsDocWorkerClient {
  constructor(worker) {
    this.worker = worker;
    this.pending = new Map();
    this.seq = 0;
    this.handleMessage = this.handleMessage.bind(this);
    this.handleError = this.handleError.bind(this);
    worker.addEventListener('message', this.handleMessage);
    worker.addEventListener('error', this.handleError);
  }

  static create(url = new URL('./worker.js', import.meta.url), workerOptions = {}) {
    return new MsDocWorkerClient(new Worker(url, { type: 'module', ...workerOptions }));
  }

  handleMessage(event) {
    const { id, ok, result, error } = event.data || {};
    if (!this.pending.has(id)) return;
    const { resolve, reject } = this.pending.get(id);
    this.pending.delete(id);
    if (ok) resolve(result);
    else reject(new Error(error || 'Worker request failed'));
  }

  handleError(event) {
    for (const { reject } of this.pending.values()) reject(event.error || new Error(event.message || 'Worker crashed'));
    this.pending.clear();
  }

  request(type, payload, transfer = []) {
    const id = ++this.seq;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.worker.postMessage({ id, type, ...payload }, transfer);
    });
  }

  parse(input, options = {}) {
    const buffer = toArrayBuffer(input);
    return this.request('parse', { buffer, options }, [buffer]);
  }

  parseToHtml(input, options = {}) {
    const buffer = toArrayBuffer(input);
    return this.request('parseToHtml', { buffer, options }, [buffer]);
  }

  render(parsed, options = {}) {
    return this.request('render', { parsed, options });
  }

  destroy() {
    this.worker.removeEventListener('message', this.handleMessage);
    this.worker.removeEventListener('error', this.handleError);
    this.worker.terminate();
    this.pending.clear();
  }
}
