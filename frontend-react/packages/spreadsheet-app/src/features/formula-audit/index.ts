import type { RangeRef } from '@react-sheets/core-model';
import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import {
  cellAddressKey,
  formatCellAddress,
  type FormulaCellEntry,
  type CellAddress,
  type FormulaDependency,
  type FormulaEngine,
  type FormulaEvaluationTrace,
  type FormulaEvaluationTraceStep,
  type FormulaError,
  type FormulaValue,
  type RecalculationMode,
  isFormulaError,
} from '@react-sheets/formula-engine';

export type FormulaAuditDirection = 'precedent' | 'dependent';

export interface FormulaAuditArrow {
  readonly id: string;
  readonly direction: FormulaAuditDirection;
  /** The formula cell that owns the relationship. */
  readonly formulaCell: CellAddress;
  /** The cell/range being referenced by the formula cell. */
  readonly target: FormulaDependency;
}

export interface FormulaAuditFormulaProjection {
  readonly address: CellAddress;
  readonly formula: string;
  readonly value: FormulaValue;
  readonly dependencies: readonly FormulaDependency[];
}

export interface FormulaAuditError {
  readonly address: CellAddress;
  readonly formula: string;
  readonly code: FormulaError['code'];
  readonly message: string;
  readonly position?: number;
}

export interface FormulaAuditEvaluationStep extends FormulaEvaluationTraceStep {
  readonly index: number;
}

export interface FormulaAuditEvaluationProjection {
  readonly address: CellAddress;
  readonly formula: string;
  readonly value: FormulaValue;
  readonly steps: readonly FormulaAuditEvaluationStep[];
}

export interface FormulaAuditProjection {
  readonly selectedCell?: CellAddress;
  readonly arrows: readonly FormulaAuditArrow[];
  readonly showFormulas: boolean;
  readonly formulas: readonly FormulaAuditFormulaProjection[];
  readonly errors: readonly FormulaAuditError[];
  readonly evaluation?: FormulaAuditEvaluationProjection;
}

export interface FormulaErrorScanOptions {
  readonly sheetId?: string;
  readonly range?: RangeRef;
}

export interface FormulaAuditControllerOptions {
  readonly showFormulas?: boolean;
}

/**
 * Read-only formula auditing controller. It owns only derived UI projection
 * state; authored formulas and calculation results remain owned by the
 * caller's FormulaEngine.
 */
export class FormulaAuditController {
  private currentFormula: FormulaEngine;
  private selectedCell?: CellAddress;
  private arrows: FormulaAuditArrow[] = [];
  private arrowDirection?: FormulaAuditDirection;
  private showFormulasValue: boolean;
  private errors: readonly FormulaAuditError[] = [];
  private evaluation?: FormulaAuditEvaluationProjection;

  constructor(
    formula: FormulaEngine,
    options: FormulaAuditControllerOptions = {},
  ) {
    this.currentFormula = formula;
    this.showFormulasValue = options.showFormulas ?? false;
  }

  get formula(): FormulaEngine {
    return this.currentFormula;
  }

  /** Rebind the controller after the host rebuilds its formula engine. */
  setFormula(formula: FormulaEngine): void {
    this.currentFormula = formula;
    this.arrows = [];
    this.errors = [];
    this.evaluation = undefined;
  }

  showPrecedents(address: CellAddress): FormulaAuditProjection {
    const cell = cloneAddress(address);
    this.selectedCell = cell;
    this.arrowDirection = 'precedent';
    this.arrows = buildPrecedentArrows(this.formula, cell);
    return this.getProjection();
  }

  showDependents(address: CellAddress): FormulaAuditProjection {
    const cell = cloneAddress(address);
    this.selectedCell = cell;
    this.arrowDirection = 'dependent';
    this.arrows = buildDependentArrows(this.formula, cell);
    return this.getProjection();
  }

  removeArrows(): FormulaAuditProjection {
    this.arrows = [];
    this.arrowDirection = undefined;
    return this.getProjection();
  }

  /** Recompute the active arrow set after a workbook/formula mutation. */
  refresh(): FormulaAuditProjection {
    if (this.selectedCell && this.arrowDirection === 'precedent') {
      this.arrows = buildPrecedentArrows(this.formula, this.selectedCell);
    } else if (this.selectedCell && this.arrowDirection === 'dependent') {
      this.arrows = buildDependentArrows(this.formula, this.selectedCell);
    }
    return this.getProjection();
  }

  setShowFormulas(enabled: boolean): FormulaAuditProjection {
    this.showFormulasValue = enabled;
    return this.getProjection();
  }

  setRecalculationMode(mode: RecalculationMode): RecalculationMode {
    this.formula.setRecalculationMode(mode);
    return this.formula.getRecalculationMode();
  }

  scanErrors(options: FormulaErrorScanOptions = {}): readonly FormulaAuditError[] {
    this.errors = scanFormulaErrors(this.formula, options);
    return this.errors.map(cloneError);
  }

  evaluateStep(address: CellAddress): FormulaAuditEvaluationProjection | undefined {
    const cell = cloneAddress(address);
    const entry = this.formula.getFormulaEntries().find((candidate) => cellAddressKey(candidate.address) === cellAddressKey(cell));
    if (!entry) {
      this.evaluation = undefined;
      return undefined;
    }
    const trace = this.formula.evaluateFormulaWithTrace(cell);
    this.evaluation = trace === undefined
      ? undefined
      : toEvaluationProjection(entry, trace);
    return this.evaluation ? cloneEvaluation(this.evaluation) : undefined;
  }

  getProjection(): FormulaAuditProjection {
    const projection: FormulaAuditProjection = {
      ...(this.selectedCell === undefined ? {} : { selectedCell: cloneAddress(this.selectedCell) }),
      arrows: this.arrows.map(cloneArrow),
      showFormulas: this.showFormulasValue,
      formulas: this.showFormulasValue ? formulaProjection(this.formula) : [],
      errors: this.errors.map(cloneError),
      ...(this.evaluation === undefined ? {} : { evaluation: cloneEvaluation(this.evaluation) }),
    };
    return projection;
  }
}

/** Build the real precedent edges from the engine's dependency index. */
export function getFormulaPrecedents(formula: FormulaEngine, address: CellAddress): readonly FormulaAuditArrow[] {
  return buildPrecedentArrows(formula, address).map(cloneArrow);
}

/** Build the real dependent edges from the engine's reverse dependency index. */
export function getFormulaDependents(formula: FormulaEngine, address: CellAddress): readonly FormulaAuditArrow[] {
  return buildDependentArrows(formula, address).map(cloneArrow);
}

/** Clear a derived arrow projection without mutating workbook or formula state. */
export function removeFormulaAuditArrows(controller: FormulaAuditController): FormulaAuditProjection {
  return controller.removeArrows();
}

/** Project all authored formulas for a Show Formulas surface. */
export function projectFormulaCells(formula: FormulaEngine): readonly FormulaAuditFormulaProjection[] {
  return formulaProjection(formula);
}

/** Scan current formula results and parse/evaluation errors in deterministic order. */
export function scanFormulaErrors(
  formula: FormulaEngine,
  options: FormulaErrorScanOptions = {},
): readonly FormulaAuditError[] {
  return formula.getFormulaEntries()
    .filter((entry) => matchesScan(entry.address, options))
    .flatMap((entry) => {
      const error = firstFormulaError(entry.value);
      return error === undefined ? [] : [{
        address: cloneAddress(entry.address),
        formula: entry.formula,
        code: error.code,
        message: error.message,
        ...(error.position === undefined ? {} : { position: error.position }),
      }];
    });
}

/** Return a real AST evaluation trace for Evaluate Formula UI. */
export function evaluateFormulaStep(
  formula: FormulaEngine,
  address: CellAddress,
): FormulaAuditEvaluationProjection | undefined {
  const entry = formula.getFormulaEntries().find((candidate) => cellAddressKey(candidate.address) === cellAddressKey(address));
  const trace = formula.evaluateFormulaWithTrace(address);
  return entry && trace ? toEvaluationProjection(entry, trace) : undefined;
}

/** Register the UI commands against the existing command registry. */
export function registerFormulaAuditCommands(
  registry: CommandRegistry,
  controller: FormulaAuditController,
): readonly string[] {
  const commandIds = [
    'formula.audit.precedents.show',
    'formula.audit.dependents.show',
    'formula.audit.arrows.remove',
    'formula.audit.formulas.show',
    'formula.audit.errors.scan',
    'formula.audit.evaluate.step',
    'formula.calculation.mode.set',
  ] as const;

  registry.registerCommand<FormulaAuditAddressParams>({
    id: 'formula.audit.precedents.show',
    execute: (params, context) => {
      assertAuditAddress(context, params.address);
      const projection = controller.showPrecedents(params.address);
      return auditResult(context, projection, params.address);
    },
  });
  registry.registerCommand<FormulaAuditAddressParams>({
    id: 'formula.audit.dependents.show',
    execute: (params, context) => {
      assertAuditAddress(context, params.address);
      const projection = controller.showDependents(params.address);
      return auditResult(context, projection, params.address);
    },
  });
  registry.registerCommand<FormulaAuditEmptyParams>({
    id: 'formula.audit.arrows.remove',
    execute: (_params, context) => {
      controller.removeArrows();
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
    },
  });
  registry.registerCommand<FormulaAuditShowFormulasParams>({
    id: 'formula.audit.formulas.show',
    execute: (params, context) => {
      const projection = controller.setShowFormulas(params.enabled);
      return auditResult(context, projection, projection.selectedCell);
    },
  });
  registry.registerCommand<FormulaAuditErrorScanParams>({
    id: 'formula.audit.errors.scan',
    execute: (params, context) => {
      if (params.sheetId) context.workbook.getSheet(params.sheetId);
      if (params.range) context.workbook.getSheet(params.range.sheetId);
      const errors = controller.scanErrors(params);
      return {
        operationId: context.operationId,
        mutationCount: 0,
        affectedRanges: params.range ? [structuredClone(params.range)] : errors.map((error) => cellRange(error.address)),
      };
    },
  });
  registry.registerCommand<FormulaAuditAddressParams>({
    id: 'formula.audit.evaluate.step',
    execute: (params, context) => {
      assertAuditAddress(context, params.address);
      const projection = controller.evaluateStep(params.address);
      if (!projection) throw new Error(`Formula not found at ${formatCellAddress(params.address, true)}`);
      return auditResult(context, projection, params.address);
    },
  });
  registry.registerCommand<FormulaCalculationModeParams>({
    id: 'formula.calculation.mode.set',
    execute: (params, context) => {
      if (params.mode !== 'automatic' && params.mode !== 'manual') {
        throw new Error('Formula calculation mode must be automatic or manual');
      }
      controller.setRecalculationMode(params.mode);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
    },
  });
  return commandIds;
}

export interface FormulaAuditAddressParams {
  readonly address: CellAddress;
}

export interface FormulaAuditShowFormulasParams {
  readonly enabled: boolean;
}

export interface FormulaAuditErrorScanParams extends FormulaErrorScanOptions {}

export interface FormulaCalculationModeParams {
  readonly mode: RecalculationMode;
}

export type FormulaAuditEmptyParams = Record<string, never>;

function assertAuditAddress(context: CommandContext, address: CellAddress): void {
  if (!address || typeof address.sheetId !== 'string' || !Number.isSafeInteger(address.row) || address.row < 0
    || !Number.isSafeInteger(address.column) || address.column < 0) {
    throw new Error('Formula audit requires a non-negative cell address');
  }
  context.workbook.getSheet(address.sheetId);
}

function auditResult(context: CommandContext, projection: FormulaAuditProjection | FormulaAuditEvaluationProjection, address?: CellAddress): CommandResult {
  const affectedRanges = address ? [cellRange(address)] : [];
  return { operationId: context.operationId, mutationCount: 0, affectedRanges };
}

function buildPrecedentArrows(formula: FormulaEngine, address: CellAddress): FormulaAuditArrow[] {
  return formula.getDependencies(address).map((target, index) => ({
    id: `precedent:${cellAddressKey(address)}:${dependencyKey(target)}:${index}`,
    direction: 'precedent',
    formulaCell: cloneAddress(address),
    target: cloneDependency(target),
  }));
}

function buildDependentArrows(formula: FormulaEngine, address: CellAddress): FormulaAuditArrow[] {
  return formula.getDependents(address).map((formulaCell, index) => ({
    id: `dependent:${cellAddressKey(formulaCell)}:${cellAddressKey(address)}:${index}`,
    direction: 'dependent',
    formulaCell: cloneAddress(formulaCell),
    target: { kind: 'cell', address: cloneAddress(address) },
  }));
}

function formulaProjection(formula: FormulaEngine): FormulaAuditFormulaProjection[] {
  return formula.getFormulaEntries().map((entry) => ({
    address: cloneAddress(entry.address),
    formula: entry.formula,
    value: structuredClone(entry.value),
    dependencies: entry.dependencies.map(cloneDependency),
  }));
}

function toEvaluationProjection(entry: FormulaCellEntry, trace: FormulaEvaluationTrace): FormulaAuditEvaluationProjection {
  return {
    address: cloneAddress(entry.address),
    formula: entry.formula,
    value: structuredClone(trace.value),
    steps: trace.steps.map((step, index) => ({
      index,
      node: structuredClone(step.node),
      expression: step.expression,
      value: structuredClone(step.value),
    })),
  };
}

function matchesScan(address: CellAddress, options: FormulaErrorScanOptions): boolean {
  if (options.sheetId !== undefined && address.sheetId !== options.sheetId) return false;
  if (!options.range) return true;
  return address.sheetId === options.range.sheetId
    && address.row >= options.range.startRow && address.row <= options.range.endRow
    && address.column >= options.range.startColumn && address.column <= options.range.endColumn;
}

function dependencyKey(dependency: FormulaDependency): string {
  return dependency.kind === 'cell'
    ? `cell:${cellAddressKey(dependency.address)}`
    : `range:${cellAddressKey(dependency.start)}:${cellAddressKey(dependency.end)}`;
}

function cloneAddress(address: CellAddress): CellAddress {
  return { sheetId: address.sheetId, row: address.row, column: address.column };
}

function cloneDependency(dependency: FormulaDependency): FormulaDependency {
  return dependency.kind === 'cell'
    ? { kind: 'cell', address: cloneAddress(dependency.address) }
    : { kind: 'range', start: cloneAddress(dependency.start), end: cloneAddress(dependency.end) };
}

function cloneArrow(arrow: FormulaAuditArrow): FormulaAuditArrow {
  return {
    id: arrow.id,
    direction: arrow.direction,
    formulaCell: cloneAddress(arrow.formulaCell),
    target: cloneDependency(arrow.target),
  };
}

function cloneError(error: FormulaAuditError): FormulaAuditError {
  return {
    address: cloneAddress(error.address),
    formula: error.formula,
    code: error.code,
    message: error.message,
    ...(error.position === undefined ? {} : { position: error.position }),
  };
}

function cloneEvaluation(evaluation: FormulaAuditEvaluationProjection): FormulaAuditEvaluationProjection {
  return {
    address: cloneAddress(evaluation.address),
    formula: evaluation.formula,
    value: structuredClone(evaluation.value),
    steps: evaluation.steps.map((step) => ({
      index: step.index,
      node: structuredClone(step.node),
      expression: step.expression,
      value: structuredClone(step.value),
    })),
  };
}

function cellRange(address: CellAddress): RangeRef {
  return {
    sheetId: address.sheetId,
    startRow: address.row,
    endRow: address.row,
    startColumn: address.column,
    endColumn: address.column,
  };
}

function firstFormulaError(value: FormulaValue): FormulaError | undefined {
  if (isFormulaError(value)) return value;
  if (!Array.isArray(value)) return undefined;
  for (const row of value) {
    for (const entry of row) {
      const error = firstFormulaError(entry);
      if (error) return error;
    }
  }
  return undefined;
}
