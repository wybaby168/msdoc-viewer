import type {
  MsDocParseOptions,
  MsDocParseResult,
  MsDocParseToHtmlOptions,
  MsDocRenderOptions,
  MsDocRenderResult,
  MsDocWorkerClientLike,
  WorkerLike,
  WorkerRequestMap,
  WorkerRequestType,
  WorkerResponse,
} from './types.js';

function toArrayBuffer(input: ArrayBuffer | ArrayBufferView): ArrayBuffer {
  if (input instanceof ArrayBuffer) return input.slice(0);
  if (ArrayBuffer.isView(input)) {
    return input.byteOffset === 0 && input.byteLength === input.buffer.byteLength
      ? input.buffer.slice(0)
      : input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
  }
  throw new TypeError('Expected ArrayBuffer or TypedArray');
}

interface PendingRequest<TResult> {
  resolve(value: TResult): void;
  reject(reason?: unknown): void;
}

/**
 * Thin typed wrapper around a module worker so applications can offload
 * expensive binary parsing without having to manage message correlation.
 */
export class MsDocWorkerClient implements MsDocWorkerClientLike {
  private readonly worker: WorkerLike;
  private readonly pending = new Map<number, PendingRequest<unknown>>();
  private seq = 0;

  constructor(worker: WorkerLike) {
    this.worker = worker;
    this.handleMessage = this.handleMessage.bind(this);
    this.handleError = this.handleError.bind(this);
    worker.addEventListener('message', this.handleMessage);
    worker.addEventListener('error', this.handleError);
  }

  static create(url = new URL('./worker.js', import.meta.url), workerOptions: WorkerOptions = {}): MsDocWorkerClient {
    return new MsDocWorkerClient(new Worker(url, { type: 'module', ...workerOptions }));
  }

  private handleMessage(event: MessageEvent): void {
    const data = (event.data ?? {}) as Partial<WorkerResponse<unknown>>;
    const id = typeof data.id === 'number' ? data.id : -1;
    if (!this.pending.has(id)) return;
    const current = this.pending.get(id) as PendingRequest<unknown>;
    this.pending.delete(id);
    if (data.ok) current.resolve(data.result);
    else current.reject(new Error(data.error || 'Worker request failed'));
  }

  private handleError(event: ErrorEvent): void {
    for (const { reject } of this.pending.values()) {
      reject(event.error || new Error(event.message || 'Worker crashed'));
    }
    this.pending.clear();
  }

  private request<TKey extends WorkerRequestType, TResult>(
    type: TKey,
    payload: WorkerRequestMap[TKey],
    transfer: Transferable[] = [],
  ): Promise<TResult> {
    const id = ++this.seq;
    return new Promise<TResult>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as PendingRequest<unknown>['resolve'], reject });
      this.worker.postMessage({ id, type, ...payload }, transfer);
    });
  }

  parse(input: ArrayBuffer | ArrayBufferView, options: MsDocParseOptions = {}): Promise<MsDocParseResult> {
    const buffer = toArrayBuffer(input);
    return this.request<'parse', MsDocParseResult>('parse', { buffer, options }, [buffer]);
  }

  parseToHtml(input: ArrayBuffer | ArrayBufferView, options: MsDocParseToHtmlOptions = {}): Promise<MsDocRenderResult> {
    const buffer = toArrayBuffer(input);
    return this.request<'parseToHtml', MsDocRenderResult>('parseToHtml', { buffer, options }, [buffer]);
  }

  render(parsed: MsDocParseResult, options: MsDocRenderOptions = {}): Promise<MsDocRenderResult> {
    return this.request<'render', MsDocRenderResult>('render', { parsed, options });
  }

  destroy(): void {
    this.worker.removeEventListener('message', this.handleMessage);
    this.worker.removeEventListener('error', this.handleError);
    this.worker.terminate();
    this.pending.clear();
  }
}
