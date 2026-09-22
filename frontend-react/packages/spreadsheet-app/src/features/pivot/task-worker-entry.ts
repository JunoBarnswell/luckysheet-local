import { evaluatePivotTask } from './engine';
import { assertPivotSourceIndex, type PivotSourceIndex } from './source-index';
import {
  assertPivotTaskRequest,
  PIVOT_TASK_PROTOCOL,
  PIVOT_TASK_VERSION,
  pivotTaskFailure,
  type PivotCalculateRequest,
  type PivotTaskRequest,
  type PivotTaskResult,
} from './task-protocol';

interface RegisteredPivotSource {
  revision: string;
  source: PivotSourceIndex;
}

export interface PivotWorkerScope {
  onmessage: ((event: { readonly data: unknown }) => void) | null;
  postMessage(message: PivotTaskResult): void;
}

/** Stateful evaluator owned by one worker or the inline test port. */
export class PivotTaskEvaluator {
  private readonly sources = new Map<string, RegisteredPivotSource>();

  consume(payload: unknown): PivotTaskResult {
    let request: PivotTaskRequest;
    try {
      assertPivotTaskRequest(payload);
      request = payload;
    } catch (error) {
      return protocolFailure(payload, error);
    }
    if (request.kind === 'cancel') {
      return envelope(request, { status: 'cancelled' });
    }
    if (request.kind === 'source-register') {
      try {
        assertPivotSourceIndex(request.source);
        this.sources.set(request.sourceIdentity, { revision: request.sourceRevision, source: request.source });
        return envelope(request, { status: 'accepted', sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision });
      } catch (error) {
        return sourceFailure(request, error, 'PIVOT_SOURCE_INVALID');
      }
    }
    if (request.kind === 'source-release') {
      const current = this.sources.get(request.sourceIdentity);
      if (current?.revision === request.sourceRevision) this.sources.delete(request.sourceIdentity);
      return envelope(request, { status: 'accepted', sourceIdentity: request.sourceIdentity, sourceRevision: request.sourceRevision });
    }
    const registered = this.sources.get(request.sourceIdentity);
    if (!registered) return pivotTaskFailure(request, new Error(`Pivot source ${request.sourceIdentity} is unavailable`), 'PIVOT_SOURCE_UNAVAILABLE');
    if (registered.revision !== request.revisions.sourceRevision) {
      return pivotTaskFailure(request, new Error(`Pivot source revision mismatch: ${registered.revision} != ${request.revisions.sourceRevision}`), 'PIVOT_TASK_REVISION_MISMATCH');
    }
    try {
      const result = evaluatePivotTask({
        definition: request.definition,
        source: registered.source,
        controls: request.controls,
        revisions: request.revisions,
        targetBounds: request.targetBounds,
      });
      return envelope(request, {
        status: 'completed',
        sourceIdentity: request.sourceIdentity,
        sourceRevision: request.revisions.sourceRevision,
        result,
      });
    } catch (error) {
      return pivotTaskFailure(request, error);
    }
  }
}

export function installPivotWorkerEntry(scope: PivotWorkerScope): () => void {
  const previous = scope.onmessage;
  const evaluator = new PivotTaskEvaluator();
  scope.onmessage = (event) => scope.postMessage(evaluator.consume(event.data));
  return () => { scope.onmessage = previous; };
}

function sourceFailure(
  request: Extract<PivotTaskRequest, { kind: 'source-register' }>,
  error: unknown,
  code: 'PIVOT_SOURCE_INVALID',
): PivotTaskResult {
  const message = error instanceof Error ? error.message : 'Pivot source registration failed';
  return {
    ...baseEnvelope(request),
    status: 'failed',
    error: {
      code,
      message,
      pivotId: 'unbound',
      sourceIdentity: request.sourceIdentity,
      sourceRevision: request.sourceRevision,
      recovery: 'fix-source',
    },
  };
}

function protocolFailure(payload: unknown, error: unknown): PivotTaskResult {
  const request = isRecord(payload) ? payload : {};
  return {
    protocol: PIVOT_TASK_PROTOCOL,
    version: PIVOT_TASK_VERSION,
    taskId: typeof request.taskId === 'string' && request.taskId ? request.taskId : 'invalid-task',
    generation: Number.isSafeInteger(request.generation) && Number(request.generation) >= 0 ? Number(request.generation) : 0,
    status: 'failed',
    error: {
      code: 'PIVOT_TASK_PROTOCOL_ERROR',
      message: error instanceof Error ? error.message : 'Pivot task protocol is invalid',
      pivotId: 'unbound',
      sourceIdentity: typeof request.sourceIdentity === 'string' ? request.sourceIdentity : 'unknown',
      sourceRevision: 'unknown',
      recovery: 'retry',
    },
  };
}

function baseEnvelope(request: Pick<PivotTaskRequest, 'taskId' | 'generation'>) {
  return { protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId: request.taskId, generation: request.generation } as const;
}

function envelope<T extends Omit<PivotTaskResult, keyof ReturnType<typeof baseEnvelope>>>(
  request: Pick<PivotTaskRequest, 'taskId' | 'generation'>,
  result: T,
): PivotTaskResult {
  return { ...baseEnvelope(request), ...result } as PivotTaskResult;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
