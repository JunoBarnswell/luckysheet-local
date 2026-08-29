import type { CellAddress } from './ast';
import type { FormulaVisibilitySnapshot } from './reference-cursor';
import type { FormulaCalculationBootstrap, FormulaSpillSpaceSnapshot } from './calculation-state';
import type { WorkbookCalculationSettings } from './calculation-settings';
import type { FormulaDefinedName } from './defined-names';
import type { FormulaDependency } from './range-index';
import type { SheetTableRef } from './sheet-table-resolver';
import type { ResolvedSpill } from './spill-resolver';
import type { FormulaValue, ScalarValue } from './values';

/**
 * Canonical calculation session wire contract.
 *
 * A session is opened once with a bootstrap. All subsequent messages carry
 * only changed inputs and dirty roots; sending a workbook snapshot for every
 * calculation is deliberately not a supported operation.
 */
export const CALCULATION_DELTA_PROTOCOL = 'react-sheets.formula-delta' as const;
export const CALCULATION_DELTA_VERSION = 1 as const;

export type CalculationSessionRequestKind = 'session.open' | 'calculation.delta' | 'calculation.cancel' | 'session.close';

export type FormulaInputDelta =
  | { readonly kind: 'set-value'; readonly address: CellAddress; readonly value: ScalarValue }
  | { readonly kind: 'set-formula'; readonly address: CellAddress; readonly formula: string }
  | { readonly kind: 'clear'; readonly address: CellAddress };

export interface FormulaCalculationDelta {
  readonly cells?: readonly FormulaInputDelta[];
  readonly definedNameModels?: readonly FormulaDefinedName[];
  readonly sheetTables?: readonly SheetTableRef[];
  readonly spillSpaces?: readonly FormulaSpillSpaceSnapshot[];
  readonly visibility?: FormulaVisibilitySnapshot;
  readonly calculationSettings?: Partial<WorkbookCalculationSettings>;
}

export interface CalculationSessionOpenRequest {
  readonly protocol: typeof CALCULATION_DELTA_PROTOCOL;
  readonly version: typeof CALCULATION_DELTA_VERSION;
  readonly kind: 'session.open';
  readonly sessionId: string;
  readonly taskId: string;
  readonly revision: number;
  readonly generation: number;
  readonly bootstrap: FormulaCalculationBootstrap;
}

export interface CalculationDeltaRequest {
  readonly protocol: typeof CALCULATION_DELTA_PROTOCOL;
  readonly version: typeof CALCULATION_DELTA_VERSION;
  readonly kind: 'calculation.delta';
  readonly sessionId: string;
  readonly taskId: string;
  readonly revision: number;
  readonly generation: number;
  readonly delta: FormulaCalculationDelta;
  readonly roots?: readonly CellAddress[];
  readonly forceRecalculate?: boolean;
}

export interface CalculationCancelRequest {
  readonly protocol: typeof CALCULATION_DELTA_PROTOCOL;
  readonly version: typeof CALCULATION_DELTA_VERSION;
  readonly kind: 'calculation.cancel';
  readonly sessionId: string;
  readonly taskId: string;
  readonly revision: number;
  readonly generation: number;
}

export interface CalculationSessionCloseRequest {
  readonly protocol: typeof CALCULATION_DELTA_PROTOCOL;
  readonly version: typeof CALCULATION_DELTA_VERSION;
  readonly kind: 'session.close';
  readonly sessionId: string;
  readonly taskId: string;
  readonly revision: number;
  readonly generation: number;
}

export type CalculationSessionRequest =
  | CalculationSessionOpenRequest
  | CalculationDeltaRequest
  | CalculationCancelRequest
  | CalculationSessionCloseRequest;

export interface CalculationCellResult {
  readonly address: CellAddress;
  readonly value: FormulaValue;
  readonly formula?: string;
  readonly dependencies: readonly FormulaDependency[];
}

export interface CalculationDeltaReport {
  readonly recalculated: readonly CellAddress[];
  readonly results: readonly CalculationCellResult[];
  readonly spills: readonly ResolvedSpill[];
  readonly pendingRoots: readonly CellAddress[];
}

export type CalculationSessionResultStatus = 'ready' | 'completed' | 'cancelled' | 'closed' | 'failed';

export interface CalculationSessionResult {
  readonly protocol: typeof CALCULATION_DELTA_PROTOCOL;
  readonly version: typeof CALCULATION_DELTA_VERSION;
  readonly kind: 'session.ready' | 'calculation.result' | 'session.closed' | 'calculation.failed';
  readonly sessionId: string;
  readonly taskId: string;
  readonly revision: number;
  readonly generation: number;
  readonly status: CalculationSessionResultStatus;
  readonly report?: CalculationDeltaReport;
  readonly error?: { readonly code: string; readonly message: string; readonly recovery?: string };
}

export interface CalculationSessionPort {
  readonly protocol: typeof CALCULATION_DELTA_PROTOCOL;
  readonly version: typeof CALCULATION_DELTA_VERSION;
  submit(request: CalculationSessionRequest): Promise<CalculationSessionResult>;
  cancel(taskId: string): void;
  dispose(): void;
}

export function assertCalculationSessionRequest(value: unknown): asserts value is CalculationSessionRequest {
  if (!isRecord(value)) throw new Error('Calculation session request must be an object');
  if (value.protocol !== CALCULATION_DELTA_PROTOCOL || value.version !== CALCULATION_DELTA_VERSION) {
    throw new Error('Unsupported calculation delta protocol');
  }
  if ('snapshot' in value || 'roots' in value && value.kind !== 'calculation.delta') throw new Error('Legacy calculation snapshot/task fields are not accepted by the delta protocol');
  if (!isNonEmptyString(value.sessionId) || !isNonEmptyString(value.taskId)) {
    throw new Error('Calculation session request requires sessionId and taskId');
  }
  if (!isRevision(value.revision) || !isRevision(value.generation)) {
    throw new Error('Calculation session revision and generation must be non-negative safe integers');
  }
  if (value.kind === 'session.open') {
    if (!isRecord(value.bootstrap)) throw new Error('Calculation session.open requires a bootstrap');
    return;
  }
  if (value.kind === 'calculation.delta') {
    if (!isRecord(value.delta)) throw new Error('Calculation delta requires a delta object');
    if (value.roots !== undefined && (!Array.isArray(value.roots) || !value.roots.every(isCellAddress))) {
      throw new Error('Calculation delta roots must be valid cell addresses');
    }
    if (value.forceRecalculate !== undefined && typeof value.forceRecalculate !== 'boolean') throw new Error('Calculation delta forceRecalculate must be boolean');
    assertCalculationDelta(value.delta);
    return;
  }
  if (value.kind !== 'calculation.cancel' && value.kind !== 'session.close') {
    throw new Error(`Unsupported calculation session request kind: ${String(value.kind)}`);
  }
}

export function assertCalculationSessionResult(value: unknown): asserts value is CalculationSessionResult {
  if (!isRecord(value)) throw new Error('Calculation session result must be an object');
  if (value.protocol !== CALCULATION_DELTA_PROTOCOL || value.version !== CALCULATION_DELTA_VERSION) {
    throw new Error('Unsupported calculation delta result protocol');
  }
  if (!isNonEmptyString(value.sessionId) || !isNonEmptyString(value.taskId)) throw new Error('Calculation result requires sessionId and taskId');
  if (!isRevision(value.revision) || !isRevision(value.generation)) throw new Error('Calculation result revision is invalid');
  if (!['session.ready', 'calculation.result', 'session.closed', 'calculation.failed'].includes(String(value.kind))) throw new Error('Calculation result kind is invalid');
  if (!['ready', 'completed', 'cancelled', 'closed', 'failed'].includes(String(value.status))) throw new Error('Calculation result status is invalid');
  if (value.status === 'completed' && !isCalculationDeltaReport(value.report)) throw new Error('Completed calculation result requires a report');
  if (value.status === 'failed' && (!isRecord(value.error) || typeof value.error.code !== 'string' || typeof value.error.message !== 'string')) {
    throw new Error('Failed calculation result requires an error');
  }
}

function assertCalculationDelta(value: Record<string, unknown>): void {
  const allowed = new Set(['cells', 'definedNameModels', 'sheetTables', 'spillSpaces', 'visibility', 'calculationSettings']);
  if (Object.keys(value).some((key) => !allowed.has(key))) throw new Error('Calculation delta contains unsupported fields');
  if (value.cells !== undefined && (!Array.isArray(value.cells) || !value.cells.every(isFormulaInputDelta))) throw new Error('Calculation delta cells are invalid');
  if (value.definedNameModels !== undefined && !Array.isArray(value.definedNameModels)) throw new Error('Calculation delta defined names are invalid');
  if (value.sheetTables !== undefined && !Array.isArray(value.sheetTables)) throw new Error('Calculation delta sheet tables are invalid');
  if (value.spillSpaces !== undefined && !Array.isArray(value.spillSpaces)) throw new Error('Calculation delta spill spaces are invalid');
  if (value.visibility !== undefined && !isRecord(value.visibility)) throw new Error('Calculation delta visibility is invalid');
  if (value.calculationSettings !== undefined && !isRecord(value.calculationSettings)) throw new Error('Calculation delta settings are invalid');
}

function isFormulaInputDelta(value: unknown): value is FormulaInputDelta {
  if (!isRecord(value) || !isCellAddress(value.address)) return false;
  if (value.kind === 'clear') return true;
  if (value.kind === 'set-formula') return typeof value.formula === 'string' && value.formula.length > 0;
  if (value.kind !== 'set-value') return false;
  return isScalarValue(value.value);
}

function isCalculationDeltaReport(value: unknown): value is CalculationDeltaReport {
  if (!isRecord(value) || !Array.isArray(value.recalculated) || !Array.isArray(value.results) || !Array.isArray(value.spills) || !Array.isArray(value.pendingRoots)) return false;
  return value.recalculated.every(isCellAddress)
    && value.pendingRoots.every(isCellAddress)
    && value.results.every((entry) => isRecord(entry) && isCellAddress(entry.address) && Array.isArray(entry.dependencies));
}

function isCellAddress(value: unknown): value is CellAddress {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && value.sheetId.length > 0
    && Number.isSafeInteger(value.row)
    && Number(value.row) >= 0
    && Number.isSafeInteger(value.column)
    && Number(value.column) >= 0;
}

function isScalarValue(value: unknown): value is ScalarValue {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === 'string' && value.trim().length > 0;
}

function isRevision(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}
