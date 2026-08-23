import type { CellAddress } from './ast';
import type { FormulaDependency } from './range-index';
import type { ResolvedSpill } from './spill-resolver';
import type { FormulaValue } from './values';

/** Stable wire identity for the calculation task transport. */
export const CALCULATION_TASK_PROTOCOL = 'react-sheets.formula-calculation' as const;
export const CALCULATION_TASK_VERSION = 1 as const;

export type CalculationTaskKind = 'recalculate';

/**
 * Serializable request sent to a calculation host.  The request contains no
 * engine instance, workbook object or callback, so it can be transferred to
 * a Worker/process without changing its shape.  `revision` lets the host
 * discard a stale result after a newer mutation has committed.
 */
export interface CalculationTaskRequest {
  readonly protocol: typeof CALCULATION_TASK_PROTOCOL;
  readonly version: typeof CALCULATION_TASK_VERSION;
  readonly taskId: string;
  readonly kind: CalculationTaskKind;
  readonly revision: number;
  readonly roots?: readonly CellAddress[];
}

export interface CalculationCellResult {
  readonly address: CellAddress;
  readonly value: FormulaValue;
  readonly formula?: string;
  readonly dependencies: readonly FormulaDependency[];
}

export interface CalculationTaskReport {
  readonly recalculated: readonly CellAddress[];
  readonly results: readonly CalculationCellResult[];
  /** Current spill projection after this task, used to replace stale spills. */
  readonly spills?: readonly ResolvedSpill[];
  /** Dirty roots that remain after the task under manual calculation mode. */
  readonly pendingRoots?: readonly CellAddress[];
}

export interface CalculationTaskResult {
  readonly protocol: typeof CALCULATION_TASK_PROTOCOL;
  readonly version: typeof CALCULATION_TASK_VERSION;
  readonly taskId: string;
  readonly revision: number;
  readonly status: 'completed' | 'cancelled' | 'failed';
  readonly report?: CalculationTaskReport;
  readonly error?: { readonly code: string; readonly message: string };
}

export interface CalculationTaskPort {
  readonly protocol: typeof CALCULATION_TASK_PROTOCOL;
  readonly version: typeof CALCULATION_TASK_VERSION;
  submit(request: CalculationTaskRequest): Promise<CalculationTaskResult>;
  cancel(taskId: string): void;
  dispose?(): void;
}

/** A cancellation message is safe to structured-clone to the browser Worker. */
export interface CalculationTaskCancellation {
  readonly protocol: typeof CALCULATION_TASK_PROTOCOL;
  readonly version: typeof CALCULATION_TASK_VERSION;
  readonly kind: 'cancel';
  readonly taskId: string;
}

export function assertCalculationTaskRequest(request: CalculationTaskRequest): void {
  if (request.protocol !== CALCULATION_TASK_PROTOCOL) {
    throw new Error(`Unsupported calculation task protocol: ${request.protocol}`);
  }
  if (request.version !== CALCULATION_TASK_VERSION) {
    throw new Error(`Unsupported calculation task version: ${request.version}`);
  }
  if (!request.taskId || typeof request.taskId !== 'string') {
    throw new Error('Calculation task requires a taskId');
  }
  if (!Number.isSafeInteger(request.revision) || request.revision < 0) {
    throw new Error('Calculation task revision must be a non-negative integer');
  }
  if (request.kind !== 'recalculate') {
    throw new Error(`Unsupported calculation task kind: ${request.kind}`);
  }
  if (request.roots !== undefined && !request.roots.every(isCellAddress)) {
    throw new Error('Calculation task roots must be valid cell addresses');
  }
}

export function isCalculationTaskCancellation(value: unknown): value is CalculationTaskCancellation {
  if (!isRecord(value)) return false;
  return value.protocol === CALCULATION_TASK_PROTOCOL
    && value.version === CALCULATION_TASK_VERSION
    && value.kind === 'cancel'
    && typeof value.taskId === 'string'
    && value.taskId.length > 0;
}

export function assertCalculationTaskResult(value: unknown): asserts value is CalculationTaskResult {
  if (!isRecord(value)) throw new Error('Calculation worker returned a non-object result');
  if (value.protocol !== CALCULATION_TASK_PROTOCOL || value.version !== CALCULATION_TASK_VERSION) {
    throw new Error('Calculation worker returned an unsupported result protocol');
  }
  if (typeof value.taskId !== 'string' || value.taskId.length === 0) {
    throw new Error('Calculation worker result requires a taskId');
  }
  if (typeof value.revision !== 'number' || !Number.isSafeInteger(value.revision) || value.revision < 0) {
    throw new Error('Calculation worker result requires a non-negative revision');
  }
  if (value.status !== 'completed' && value.status !== 'cancelled' && value.status !== 'failed') {
    throw new Error('Calculation worker result has an invalid status');
  }
  if (value.status === 'completed' && !isCalculationTaskReport(value.report)) {
    throw new Error('Completed calculation worker result requires a report');
  }
  if (value.status === 'failed' && (!isRecord(value.error) || typeof value.error.code !== 'string' || typeof value.error.message !== 'string')) {
    throw new Error('Failed calculation worker result requires an error');
  }
}

function isCellAddress(value: CellAddress): boolean {
  return Boolean(value)
    && typeof value.sheetId === 'string'
    && value.sheetId.length > 0
    && Number.isSafeInteger(value.row)
    && value.row >= 0
    && Number.isSafeInteger(value.column)
    && value.column >= 0;
}

function isCalculationTaskReport(value: unknown): value is CalculationTaskReport {
  if (!isRecord(value) || !Array.isArray(value.recalculated) || !Array.isArray(value.results)) return false;
  return value.recalculated.every((address) => isCellAddress(address as CellAddress))
    && value.results.every(isCalculationCellResult)
    && (value.spills === undefined || Array.isArray(value.spills))
    && (value.pendingRoots === undefined || (Array.isArray(value.pendingRoots) && value.pendingRoots.every((address) => isCellAddress(address as CellAddress))));
}

function isCalculationCellResult(value: unknown): value is CalculationCellResult {
  if (!isRecord(value) || !isCellAddress(value.address as CellAddress) || !Array.isArray(value.dependencies)) return false;
  return typeof value.formula === 'undefined' || typeof value.formula === 'string';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Explicit synchronous fallback for environments without a browser Worker.
 * It does not claim thread isolation; browser callers use BrowserCalculationTaskPort.
 */
export class InlineCalculationTaskPort implements CalculationTaskPort {
  readonly protocol = CALCULATION_TASK_PROTOCOL;
  readonly version = CALCULATION_TASK_VERSION;

  private readonly cancelled = new Set<string>();

  constructor(
    private readonly execute: (request: CalculationTaskRequest) => CalculationTaskReport,
  ) {}

  submit(request: CalculationTaskRequest): Promise<CalculationTaskResult> {
    assertCalculationTaskRequest(request);
    if (this.cancelled.delete(request.taskId)) {
      return Promise.resolve({
        protocol: CALCULATION_TASK_PROTOCOL,
        version: CALCULATION_TASK_VERSION,
        taskId: request.taskId,
        revision: request.revision,
        status: 'cancelled',
      });
    }
    try {
      const report = this.execute(request);
      if (this.cancelled.delete(request.taskId)) {
        return Promise.resolve({
          protocol: CALCULATION_TASK_PROTOCOL,
          version: CALCULATION_TASK_VERSION,
          taskId: request.taskId,
          revision: request.revision,
          status: 'cancelled',
        });
      }
      return Promise.resolve({
        protocol: CALCULATION_TASK_PROTOCOL,
        version: CALCULATION_TASK_VERSION,
        taskId: request.taskId,
        revision: request.revision,
        status: 'completed',
        report,
      });
    } catch (error) {
      return Promise.resolve({
        protocol: CALCULATION_TASK_PROTOCOL,
        version: CALCULATION_TASK_VERSION,
        taskId: request.taskId,
        revision: request.revision,
        status: 'failed',
        error: {
          code: 'CALCULATION_TASK_FAILED',
          message: error instanceof Error ? error.message : 'Calculation task failed',
        },
      });
    }
  }

  cancel(taskId: string): void {
    if (taskId) this.cancelled.add(taskId);
  }
}
