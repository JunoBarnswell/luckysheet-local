import type { CellAddress } from './ast';
import type { FormulaDependency } from './range-index';
import type { FormulaValue } from './values';

/** Stable wire/protocol identity for the future calculation worker. */
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

function isCellAddress(value: CellAddress): boolean {
  return Boolean(value)
    && typeof value.sheetId === 'string'
    && value.sheetId.length > 0
    && Number.isSafeInteger(value.row)
    && value.row >= 0
    && Number.isSafeInteger(value.column)
    && value.column >= 0;
}

/**
 * A real synchronous host for environments which do not have a Worker yet.
 * It implements the same versioned contract and never pretends to provide
 * thread isolation; a Worker host can replace it later without changing the
 * request/result protocol.
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

