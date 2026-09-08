import type {
  PivotDefinition,
  PivotResultTree,
  PivotSource,
  RangeRef,
} from '@react-sheets/core-model';
import type { PivotTaskControl } from './engine';
import {
  pivotTreeFromSparseResult,
  type PivotAnalyticsSparseResult,
} from './rust-analytics';

/**
 * The server owns the analytics task.  A source registration identifies the
 * persisted source and its revision; it never transfers a worksheet snapshot
 * or a TypeScript SourceRow array to the browser.
 */
export interface ServerPivotSourceRegisterRequest {
  unitId: string;
  taskId: string;
  generation: number;
  sourceIdentity: string;
  sourceRevision: string;
  source: PivotSource;
  ranges: RangeRef[];
}

export interface ServerPivotPrepareRequest {
  unitId: string;
  taskId: string;
  generation: number;
  revision: number;
  sourceIdentity: string;
  sourceRevision: string;
  layoutRevision: string;
  filterRevision: string;
  definition: PivotDefinition;
  controls: PivotTaskControl[];
  targetBounds: { rowCount: number; columnCount: number };
}

export interface ServerPivotExecuteRequest extends Omit<ServerPivotPrepareRequest, 'controls' | 'targetBounds' | 'definition'> {
  prepareToken: string;
}

export interface ServerPivotViewportRequest extends ServerPivotExecuteRequest {
  viewport: { rowOffset: number; columnOffset: number; rowLimit: number; columnLimit: number };
}

export interface ServerPivotDrilldownRequest extends ServerPivotExecuteRequest {
  drilldown: { rowKeys: import('@react-sheets/core-model').PivotScalar[]; columnKeys: import('@react-sheets/core-model').PivotScalar[]; offset: number; limit: number };
}

export interface ServerPivotCancelRequest {
  unitId: string;
  taskId: string;
  generation: number;
  reason: 'cancelled' | 'superseded' | 'disposed';
}

export interface ServerPivotAcceptedResponse {
  status: 'accepted';
  taskId: string;
  generation: number;
  revision: number;
  sourceIdentity: string;
  sourceRevision: string;
}

export interface ServerPivotPreparedResponse extends ServerPivotAcceptedResponse {
  prepareToken: string;
}

export interface ServerPivotCompletedResponse {
  status: 'completed';
  taskId: string;
  generation: number;
  revision: number;
  sourceIdentity: string;
  sourceRevision: string;
  result: PivotAnalyticsSparseResult | PivotResultTree;
}

export interface ServerPivotDrilldownResponse {
  status: 'completed';
  taskId: string;
  generation: number;
  revision: number;
  sourceIdentity: string;
  sourceRevision: string;
  drilldown: {
    sourceRows: number[];
    values: import('@react-sheets/core-model').PivotScalar[][];
    columns: number[];
    total: number;
    offset: number;
  };
}

export interface ServerPivotTaskTransport {
  registerSource(request: ServerPivotSourceRegisterRequest, signal: AbortSignal): Promise<ServerPivotAcceptedResponse>;
  prepare(request: ServerPivotPrepareRequest, signal: AbortSignal): Promise<ServerPivotPreparedResponse>;
  execute(request: ServerPivotExecuteRequest, signal: AbortSignal): Promise<ServerPivotCompletedResponse>;
  viewport(request: ServerPivotViewportRequest, signal: AbortSignal): Promise<ServerPivotCompletedResponse>;
  drilldown(request: ServerPivotDrilldownRequest, signal: AbortSignal): Promise<ServerPivotDrilldownResponse>;
  cancel(request: ServerPivotCancelRequest): Promise<void>;
}

export type ServerPivotPhaseResult =
  | ServerPivotAcceptedResponse
  | ServerPivotPreparedResponse
  | ServerPivotCompletedResponse
  | ServerPivotDrilldownResponse;

export interface ServerPivotTaskPortOptions {
  unitId: string;
  revision: () => number;
  transport: ServerPivotTaskTransport;
}

interface ActiveTask {
  generation: number;
  pivotId: string;
  sourceIdentity?: string;
  revision?: number;
  controller: AbortController;
  definition?: PivotDefinition;
  sourceRevision?: string;
  layoutRevision?: string;
  filterRevision?: string;
  prepare?: ServerPivotPreparedResponse;
}

/**
 * Canonical frontend/server task boundary for Pivot analytics.
 *
 * Every phase carries the same workbook revision and source revision. A
 * response from an older generation or revision is rejected before it can be
 * converted into a presentation tree. Cancellation aborts the local request,
 * asks the server to cancel the native task, and leaves no result to publish.
 */
export class ServerPivotTaskPort {
  private readonly active = new Map<string, ActiveTask>();
  private readonly latestGeneration = new Map<string, number>();
  private disposed = false;

  constructor(private readonly options: ServerPivotTaskPortOptions) {}

  registerSource(request: ServerPivotSourceRegisterRequest): Promise<ServerPivotAcceptedResponse> {
    this.assertUnit(request.unitId);
    this.assertSourceRegistrationRequest(request);
    const active = this.begin(request.taskId, request.generation, request.sourceIdentity, request.sourceIdentity, undefined, request.sourceRevision);
    return this.options.transport.registerSource(request, active.controller.signal)
      .then((response) => this.assertAccepted(request.taskId, request.generation, request.sourceIdentity, request.sourceRevision, response))
      .finally(() => this.finish(request.taskId));
  }

  prepare(request: ServerPivotPrepareRequest): Promise<ServerPivotPreparedResponse> {
    this.assertUnit(request.unitId);
    this.assertPrepareRequest(request);
    const active = this.begin(request.taskId, request.generation, request.definition.id, request.sourceIdentity, request.revision, request.sourceRevision);
    return this.options.transport.prepare(request, active.controller.signal)
      .then((response) => {
        const accepted = this.assertAccepted(request.taskId, request.generation, request.sourceIdentity, request.sourceRevision, response);
        if (!response.prepareToken.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: server returned an empty prepare token');
        active.definition = structuredClone(request.definition);
        active.sourceRevision = request.sourceRevision;
        active.layoutRevision = request.layoutRevision;
        active.filterRevision = request.filterRevision;
        active.prepare = response;
        return accepted as ServerPivotPreparedResponse;
      })
      .catch((error) => {
        this.finish(request.taskId);
        throw error;
      });
  }

  execute(request: ServerPivotExecuteRequest): Promise<ServerPivotCompletedResponse & { result: PivotResultTree }> {
    this.assertUnit(request.unitId);
    const active = this.requirePrepared(request.taskId, request.generation, request.revision, request.sourceIdentity, request.sourceRevision);
    if (request.prepareToken !== active.prepare!.prepareToken) throw new Error('PIVOT_TASK_REVISION_MISMATCH: prepare token does not match the active task');
    if (request.layoutRevision !== active.layoutRevision || request.filterRevision !== active.filterRevision) throw new Error('PIVOT_TASK_REVISION_MISMATCH: layout or filter revision changed within a task');
    return this.options.transport.execute(request, active.controller.signal)
      .then((response) => this.completed(request, response));
  }

  viewport(request: ServerPivotViewportRequest): Promise<ServerPivotCompletedResponse & { result: PivotResultTree }> {
    this.assertUnit(request.unitId);
    const active = this.requirePrepared(request.taskId, request.generation, request.revision, request.sourceIdentity, request.sourceRevision);
    if (request.prepareToken !== active.prepare!.prepareToken) throw new Error('PIVOT_TASK_REVISION_MISMATCH: prepare token does not match the active task');
    if (request.layoutRevision !== active.layoutRevision || request.filterRevision !== active.filterRevision) throw new Error('PIVOT_TASK_REVISION_MISMATCH: layout or filter revision changed within a task');
    return this.options.transport.viewport(request, active.controller.signal)
      .then((response) => this.completed(request, response));
  }

  drilldown(request: ServerPivotDrilldownRequest): Promise<ServerPivotDrilldownResponse> {
    this.assertUnit(request.unitId);
    const active = this.requirePrepared(request.taskId, request.generation, request.revision, request.sourceIdentity, request.sourceRevision);
    if (request.prepareToken !== active.prepare!.prepareToken) throw new Error('PIVOT_TASK_REVISION_MISMATCH: prepare token does not match the active task');
    if (request.layoutRevision !== active.layoutRevision || request.filterRevision !== active.filterRevision) throw new Error('PIVOT_TASK_REVISION_MISMATCH: layout or filter revision changed within a task');
    return this.options.transport.drilldown(request, active.controller.signal)
      .then((response) => {
        this.assertPhase(request.taskId, request.generation, request.sourceIdentity, request.sourceRevision, response);
        return response;
      });
  }

  cancel(taskId: string, reason: ServerPivotCancelRequest['reason'] = 'cancelled'): void {
    const active = this.active.get(taskId);
    if (!active) return;
    this.active.delete(taskId);
    active.controller.abort(reason);
    void this.options.transport.cancel({ unitId: this.options.unitId, taskId, generation: active.generation, reason }).catch(() => {
      // The caller has already received the cancellation boundary. A server
      // cancellation failure must be observable by the next task, not turn a
      // cancelled task into a successful result.
    });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    for (const taskId of [...this.active.keys()]) this.cancel(taskId, 'disposed');
    this.latestGeneration.clear();
  }

  private begin(
    taskId: string,
    generation: number,
    pivotId: string,
    sourceIdentity?: string,
    revision?: number,
    sourceRevision?: string,
  ): ActiveTask {
    if (this.disposed) throw new Error('PIVOT_TASK_FAILED: server Pivot task port has been disposed');
    const current = this.latestGeneration.get(pivotId) ?? 0;
    if (!Number.isSafeInteger(generation) || generation <= 0) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: generation must be a positive safe integer');
    if (generation < current) throw new Error('PIVOT_TASK_REVISION_MISMATCH: Pivot task generation is stale');
    const existing = this.active.get(taskId);
    if (existing) throw new Error(`PIVOT_TASK_PROTOCOL_ERROR: task already exists: ${taskId}`);
    if (generation > current) {
      for (const [id, task] of this.active) {
        if (task.pivotId === pivotId && task.generation < generation) this.cancel(id, 'superseded');
      }
    }
    this.latestGeneration.set(pivotId, generation);
    const active = { generation, pivotId, sourceIdentity, revision, sourceRevision, controller: new AbortController() } satisfies ActiveTask;
    this.active.set(taskId, active);
    return active;
  }

  private requirePrepared(taskId: string, generation: number, revision: number, sourceIdentity: string, sourceRevision: string): ActiveTask {
    const active = this.active.get(taskId);
    if (!active) throw new Error('PIVOT_TASK_CANCELLED: Pivot task is no longer active');
    this.assertRequest(taskId, generation, revision, sourceIdentity, sourceRevision, active);
    if (!active.prepare?.prepareToken) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: execute requires a successful prepare phase');
    if (!active.layoutRevision || !active.filterRevision) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: prepared revisions are unavailable');
    return active;
  }

  private completed(
    request: ServerPivotExecuteRequest | ServerPivotViewportRequest,
    response: ServerPivotCompletedResponse,
  ): ServerPivotCompletedResponse & { result: PivotResultTree } {
    this.assertPhase(request.taskId, request.generation, request.sourceIdentity, request.sourceRevision, response);
    const active = this.active.get(request.taskId);
    const sourceDefinition = active?.definition;
    if (!sourceDefinition) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: prepared definition is unavailable');
    const result = isPivotResultTree(response.result)
      ? structuredClone(response.result)
      : pivotTreeFromSparseResult(sourceDefinition, response.result, {
        sourceRevision: active.sourceRevision ?? request.sourceRevision,
        layoutRevision: active.layoutRevision ?? '',
        filterRevision: active.filterRevision ?? '',
      });
    if (result.sourceRevision !== (active.sourceRevision ?? request.sourceRevision)
      || result.layoutRevision !== (active.layoutRevision ?? request.layoutRevision)
      || result.filterRevision !== (active.filterRevision ?? request.filterRevision)) {
      throw new Error('PIVOT_TASK_REVISION_MISMATCH: result metadata does not match the prepared task');
    }
    return { ...response, result };
  }

  private assertAccepted(
    taskId: string,
    generation: number,
    sourceIdentity: string,
    sourceRevision: string,
    response: ServerPivotAcceptedResponse | ServerPivotPreparedResponse,
  ): ServerPivotAcceptedResponse | ServerPivotPreparedResponse {
    this.assertPhase(taskId, generation, sourceIdentity, sourceRevision, response);
    return response;
  }

  private assertPhase(taskId: string, generation: number, sourceIdentity: string, sourceRevision: string, response: ServerPivotPhaseResult): void {
    const active = this.active.get(taskId);
    if (!active) throw new Error('PIVOT_TASK_CANCELLED: response arrived after cancellation');
    this.assertRequest(taskId, generation, active.revision ?? -1, sourceIdentity, sourceRevision, active);
    if (response.status !== 'accepted' && response.status !== 'completed') throw new Error('PIVOT_TASK_PROTOCOL_ERROR: server returned a non-terminal phase status');
    if (response.taskId !== taskId || response.generation !== generation || response.sourceIdentity !== sourceIdentity || response.sourceRevision !== sourceRevision || response.revision !== this.options.revision()) {
      throw new Error('PIVOT_TASK_REVISION_MISMATCH: server response does not match the pinned task');
    }
  }

  private assertRequest(taskId: string, generation: number, revision: number, sourceIdentity: string, sourceRevision: string, active: ActiveTask): void {
    if (active.generation !== generation || this.latestGeneration.get(active.pivotId) !== generation) throw new Error('PIVOT_TASK_REVISION_MISMATCH: Pivot task generation is stale');
    if (active.revision !== undefined && (revision !== active.revision || revision !== this.options.revision())) throw new Error('PIVOT_TASK_REVISION_MISMATCH: task revision is stale');
    if (active.sourceIdentity !== undefined && active.sourceIdentity !== sourceIdentity) throw new Error('PIVOT_TASK_REVISION_MISMATCH: source identity changed within a task');
    if (active.sourceRevision !== undefined && active.sourceRevision !== sourceRevision) throw new Error('PIVOT_TASK_REVISION_MISMATCH: source revision changed within a task');
    if (!sourceIdentity.trim() || !sourceRevision.trim() || !taskId.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: task identity is incomplete');
  }

  private assertUnit(unitId: string): void {
    if (unitId !== this.options.unitId) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: request workbook identity does not match the task port');
  }

  private assertSourceRegistrationRequest(request: ServerPivotSourceRegisterRequest): void {
    if (!request.taskId.trim() || !request.sourceIdentity.trim() || !request.sourceRevision.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: source registration identity is incomplete');
    if (!request.ranges.length && request.source.kind !== 'data-source') throw new Error('PIVOT_SOURCE_INVALID: source registration requires at least one range');
  }

  private assertPrepareRequest(request: ServerPivotPrepareRequest): void {
    if (!request.taskId.trim() || !request.sourceIdentity.trim() || !request.sourceRevision.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: prepare identity is incomplete');
    if (!Number.isSafeInteger(request.revision) || request.revision < 0 || request.revision !== this.options.revision()) throw new Error('PIVOT_TASK_REVISION_MISMATCH: prepare revision is not current');
    if (!request.layoutRevision.trim() || !request.filterRevision.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: layout and filter revisions are required');
    if (!request.definition.id.trim()) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: Pivot definition id is required');
    if (!Number.isSafeInteger(request.targetBounds.rowCount) || request.targetBounds.rowCount <= 0 || !Number.isSafeInteger(request.targetBounds.columnCount) || request.targetBounds.columnCount <= 0) throw new Error('PIVOT_TASK_PROTOCOL_ERROR: target bounds are invalid');
  }

  private finish(taskId: string): void {
    this.active.delete(taskId);
  }
}

function isPivotResultTree(value: unknown): value is PivotResultTree {
  return Boolean(value && typeof value === 'object' && (value as { schema?: unknown }).schema === 'PivotResultTree');
}
