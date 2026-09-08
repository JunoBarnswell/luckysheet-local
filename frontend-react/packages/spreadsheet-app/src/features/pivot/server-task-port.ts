import type { PivotDefinition, PivotResultTree, RangeRef } from '@react-sheets/core-model';
import type { WorkbookApiClient } from '@react-sheets/protocol';
import type { PivotTaskControl, PivotRevisionKey } from './engine';
import {
  createPivotAnalyticsRequest,
  pivotTreeFromSparseResult,
  type PivotAnalyticsSparseResult,
} from './rust-analytics';

export type PivotTaskErrorCode =
  | 'PIVOT_SOURCE_INVALID'
  | 'PIVOT_SOURCE_UNAVAILABLE'
  | 'PIVOT_PERMISSION_DENIED'
  | 'PIVOT_TASK_CANCELLED'
  | 'PIVOT_TASK_FAILED'
  | 'PIVOT_TASK_PROTOCOL_ERROR'
  | 'PIVOT_TASK_REVISION_MISMATCH';

export interface PivotTaskError {
  code: PivotTaskErrorCode;
  message: string;
  pivotId: string;
  sourceIdentity: string;
  sourceRevision: string;
  recovery: 'retry' | 'fix-source';
}

export interface ServerPivotRunRequest {
  taskId: string;
  generation: number;
  definition: PivotDefinition;
  controls: readonly PivotTaskControl[];
  revisions: PivotRevisionKey;
  source: RangeRef;
  targetBounds: { rowCount: number; columnCount: number };
}

export interface ServerPivotTaskPortOptions {
  unitId: string;
  revision: () => number;
  api: Pick<WorkbookApiClient, 'prepareAnalytics' | 'executeAnalytics' | 'cancelAnalytics'>;
}

interface ActiveTask {
  pivotId: string;
  generation: number;
  revision: number;
  controller: AbortController;
}

/** Server-owned Pivot task boundary. The browser sends only a revision-pinned
 * analytics descriptor; the server loads canonical pages and runs Rust. */
export class ServerPivotTaskPort {
  private readonly active = new Map<string, ActiveTask>();
  private readonly latestGeneration = new Map<string, number>();
  private disposed = false;

  constructor(private readonly options: ServerPivotTaskPortOptions) {}

  async run(request: ServerPivotRunRequest): Promise<PivotResultTree> {
    this.assertRequest(request);
    const revision = this.options.revision();
    const currentGeneration = this.latestGeneration.get(request.definition.id) ?? 0;
    if (request.generation < currentGeneration) throw new Error('PIVOT_TASK_REVISION_MISMATCH: generation is stale');
    for (const [taskId, active] of this.active) {
      if (active.pivotId === request.definition.id && active.generation < request.generation) this.cancel(taskId, 'superseded');
    }
    const active: ActiveTask = {
      pivotId: request.definition.id,
      generation: request.generation,
      revision,
      controller: new AbortController(),
    };
    this.latestGeneration.set(request.definition.id, request.generation);
    this.active.set(request.taskId, active);
    const analyticsRequest = createPivotAnalyticsRequest({
      unitId: this.options.unitId,
      revision,
      source: request.source,
      definition: request.definition,
    });
    try {
      const prepared = await this.options.api.prepareAnalytics(this.options.unitId, {
        queryId: request.taskId,
        revision,
        request: analyticsRequest as unknown as Record<string, unknown>,
      }, { signal: active.controller.signal });
      this.assertActive(request, active);
      if (prepared.queryId !== request.taskId || prepared.sourceRevision !== revision || !prepared.executionToken.trim()) {
        throw new Error('PIVOT_TASK_PROTOCOL_ERROR: prepare response identity is invalid');
      }
      const raw = await this.options.api.executeAnalytics(this.options.unitId, request.taskId, {
        executionToken: prepared.executionToken,
        request: analyticsRequest as unknown as Record<string, unknown>,
      }, 'execute', { signal: active.controller.signal });
      this.assertActive(request, active);
      const result = parseSparseResult(raw, request.taskId, prepared.executionToken, revision);
      const tree = pivotTreeFromSparseResult(request.definition, result, request.revisions);
      if (tree.sourceRevision !== request.revisions.sourceRevision
        || tree.layoutRevision !== request.revisions.layoutRevision
        || tree.filterRevision !== request.revisions.filterRevision) {
        throw new Error('PIVOT_TASK_REVISION_MISMATCH: result metadata is stale');
      }
      return tree;
    } finally {
      if (this.active.get(request.taskId) === active) this.active.delete(request.taskId);
    }
  }

  cancel(taskId: string, _reason: 'cancelled' | 'superseded' | 'disposed' = 'cancelled'): void {
    const active = this.active.get(taskId);
    if (!active) return;
    this.active.delete(taskId);
    active.controller.abort();
    void this.options.api.cancelAnalytics(this.options.unitId, taskId).catch(() => undefined);
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const taskId of [...this.active.keys()]) this.cancel(taskId, 'disposed');
    this.latestGeneration.clear();
  }

  private assertRequest(request: ServerPivotRunRequest): void {
    if (this.disposed) throw new Error('PIVOT_TASK_FAILED: task port is disposed');
    if (!request.taskId.trim() || !request.definition.id.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: task identity is missing');
    if (!Number.isSafeInteger(request.generation) || request.generation <= 0) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: generation is invalid');
    if (request.revisions.pivotId !== request.definition.id) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: revision identity is invalid');
    if (request.targetBounds.rowCount <= 0 || request.targetBounds.columnCount <= 0) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: target bounds are invalid');
  }

  private assertActive(request: ServerPivotRunRequest, active: ActiveTask): void {
    if (active.controller.signal.aborted || this.active.get(request.taskId) !== active) throw new Error('PIVOT_TASK_CANCELLED: task was cancelled');
    if (active.revision !== this.options.revision() || this.latestGeneration.get(active.pivotId) !== request.generation) {
      throw new Error('PIVOT_TASK_REVISION_MISMATCH: workbook revision or generation changed');
    }
  }
}

function parseSparseResult(raw: unknown, taskId: string, executionToken: string, revision: number): PivotAnalyticsSparseResult {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: analytics response is not an object');
  const value = raw as Record<string, unknown>;
  if (value.queryId !== taskId || value.executionToken !== executionToken || value.sourceRevision !== revision
    || value.kind !== 'pivot' || value.revision !== revision || !Array.isArray(value.rows)
    || !Array.isArray(value.columns) || !Array.isArray(value.cells)) {
    throw new Error('PIVOT_TASK_REVISION_MISMATCH: analytics response identity is invalid');
  }
  return value as unknown as PivotAnalyticsSparseResult;
}
