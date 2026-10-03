import type { FormulaEvaluationContext, FormulaCellOverride } from './evaluator';
import type { CellAddress, FormulaAst, FormulaReferenceNode } from './ast';
import type { RangeDependency } from './range-index';
import { isFormulaInputFault, type FormulaInputFault } from './input-fault';
import type { FormulaValue } from './values';

/** One evaluation owns every input it actually consumes, including lazy cursors. */
export class FormulaInputBoundary {
  readonly context: FormulaEvaluationContext;
  failure: FormulaInputFault | undefined;

  constructor(context: FormulaEvaluationContext) {
    this.context = {
      ...context,
      readCell: address => this.observe(context.readCell(address)),
      readRange: range => this.values(context.readRange(range)),
      ...(context.readRangeMatrix ? { readRangeMatrix: (range: RangeDependency) => this.observe(context.readRangeMatrix!(range)) } : {}),
      ...(context.readSparseRange ? { readSparseRange: (range: RangeDependency) => this.values(context.readSparseRange!(range)) } : {}),
      ...(context.readSparseRangeCells ? { readSparseRangeCells: (range: RangeDependency) => {
        const owner = this;
        return (function* () { for (const cell of context.readSparseRangeCells!(range)) {
          owner.observe(cell.value); yield cell;
        } })();
      } } : {}),
      ...(context.readSpillValue ? { readSpillValue: (address: CellAddress) => this.observe(context.readSpillValue!(address)) } : {}),
      ...(context.resolveName ? { resolveName: (name: string) => this.observe(context.resolveName!(name)) } : {}),
      ...(context.resolveReference ? { resolveReference: (reference: FormulaReferenceNode) => this.observe(context.resolveReference!(reference)) } : {}),
      ...(context.resolveTableReference ? { resolveTableReference: (name: string, request: Parameters<NonNullable<FormulaEvaluationContext['resolveTableReference']>>[1]) => this.observe(context.resolveTableReference!(name, request)) } : {}),
      ...(context.random ? { random: (name: string, occurrence?: string, element?: number) => this.observe(context.random!(name, occurrence, element)) } : {}),
      ...(context.evaluateWithCellOverrides ? { evaluateWithCellOverrides: (ast: FormulaAst, overrides: readonly FormulaCellOverride[]) => this.observe(context.evaluateWithCellOverrides!(ast, overrides)) } : {}),
    };
  }

  observe<T>(value: T): T {
    if (isFormulaInputFault(value)) this.failure ??= value;
    else if (Array.isArray(value)) for (const item of value) this.observe(item);
    return value;
  }

  private *values(values: Iterable<FormulaValue>): Iterable<FormulaValue> {
    for (const value of values) yield this.observe(value);
  }
}
