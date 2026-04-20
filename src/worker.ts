import { parseMsDoc } from './msdoc/parser.js';
import { renderMsDoc } from './render/html.js';
import type {
  MsDocParseResult,
  MsDocRenderResult,
  WorkerRequestMap,
  WorkerRequestType,
} from './types.js';

type ParseWorkerMessage = { id: number; type: 'parse' } & WorkerRequestMap['parse'];
type RenderWorkerMessage = { id: number; type: 'render' } & WorkerRequestMap['render'];
type ParseToHtmlWorkerMessage = { id: number; type: 'parseToHtml' } & WorkerRequestMap['parseToHtml'];

type WorkerMessage = ParseWorkerMessage | RenderWorkerMessage | ParseToHtmlWorkerMessage;
type WorkerEnvelope = Partial<WorkerMessage> & { id?: number; type?: WorkerRequestType };

const workerScope = self as DedicatedWorkerGlobalScope;

workerScope.addEventListener('message', (event: MessageEvent<WorkerEnvelope>) => {
  const message = event.data ?? {};
  const id = typeof message.id === 'number' ? message.id : -1;
  const type = message.type;

  try {
    let result: MsDocParseResult | MsDocRenderResult;

    if (type === 'parse') {
      const buffer = (message as Partial<ParseWorkerMessage>).buffer;
      const options = (message as Partial<ParseWorkerMessage>).options;
      result = parseMsDoc(buffer as ArrayBuffer, options ?? {});
    } else if (type === 'render') {
      const parsed = (message as Partial<RenderWorkerMessage>).parsed;
      const options = (message as Partial<RenderWorkerMessage>).options;
      result = renderMsDoc(parsed as MsDocParseResult, options ?? {});
    } else if (type === 'parseToHtml') {
      const buffer = (message as Partial<ParseToHtmlWorkerMessage>).buffer;
      const options = (message as Partial<ParseToHtmlWorkerMessage>).options;
      const parseOptions = options?.parseOptions ?? {};
      const renderOptions = options?.renderOptions ?? {};
      result = renderMsDoc(parseMsDoc(buffer as ArrayBuffer, parseOptions), renderOptions);
    } else {
      throw new Error(`Unsupported worker request type: ${String(type)}`);
    }

    workerScope.postMessage({ id, ok: true, result });
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    workerScope.postMessage({ id, ok: false, error: errorMessage });
  }
});
