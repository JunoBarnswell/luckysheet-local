import type { FormulaEvaluationContext, FormulaEvaluationValue } from '../evaluator';
import { matchesCriteria, parseCriteria, type CriteriaExpression } from '../criteria';
import { normalizeExcelPrecision } from '../numeric';
import { createRangeView, projectRangeView, sameRangeShape, type FormulaRangeView } from '../range-view';
import { createFormulaError, isFormulaError, type FormulaValue } from '../values';

export const CONDITIONAL_FUNCTIONS = new Set(['COUNTIF', 'COUNTIFS', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS', 'MINIFS', 'MAXIFS']);

export function evaluateConditional(name: string, args: readonly FormulaEvaluationValue[], context?: FormulaEvaluationContext): FormulaValue {
  const countOnly = name.startsWith('COUNT');
  const single = name === 'COUNTIF' || name === 'SUMIF' || name === 'AVERAGEIF';
  if (single ? args.length < 2 || args.length > (countOnly ? 2 : 3) : args.length < (countOnly ? 2 : 3) || args.length % 2 !== (countOnly ? 0 : 1)) return createFormulaError('#VALUE!', `${name} requires complete criteria pairs`);
  const pairs: Array<{ view: FormulaRangeView; criterion: CriteriaExpression }> = [];
  for (let index = single || countOnly ? 0 : 1; index < (single ? 2 : args.length); index += 2) {
    const view = createRangeView(args[index]!, context);
    if (isFormulaError(view)) return view;
    const criterionView = createRangeView(args[index + 1]!, context);
    if (isFormulaError(criterionView)) return criterionView;
    if (criterionView.rows !== 1 || criterionView.columns !== 1) return createFormulaError('#VALUE!', `${name} criteria must be scalar`);
    const criterion = criterionView.read(0, 0);
    if (isFormulaError(criterion)) return criterion;
    pairs.push({ view, criterion: parseCriteria(criterion) });
  }
  const shape = pairs[0]!.view;
  if (pairs.some(({ view }) => !sameRangeShape(shape, view))) return createFormulaError('#VALUE!', 'Criteria ranges must have identical shape');
  for (const { view } of pairs) for (const { value } of view.entries()) if (isFormulaError(value) && value.code === '#BLOCKED!') return value;
  const passes = (row: number, column: number) => pairs.every(({ view, criterion }) => matchesCriteria(view.read(row, column), criterion));
  if (countOnly) {
    const blanksPass = pairs.every(({ criterion }) => matchesCriteria(null, criterion));
    let count = blanksPass ? shape.rows * shape.columns : 0;
    const seen = new Set<number>();
    for (const { view } of pairs) for (const entry of view.entries()) {
      const key = entry.row * shape.columns + entry.column;
      if (seen.has(key)) continue;
      seen.add(key);
      count += Number(passes(entry.row, entry.column)) - Number(blanksPass);
    }
    return count;
  }
  let target = createRangeView(args[single ? (args.length === 3 ? 2 : 0) : 0]!, context);
  if (isFormulaError(target)) return target;
  if (single) target = projectRangeView(target, shape.rows, shape.columns, context);
  else if (!sameRangeShape(target, shape)) return createFormulaError('#VALUE!', 'Target and criteria ranges must have identical shape');
  if (name === 'MINIFS' || name === 'MAXIFS') {
    const candidates = new Set<number>();
    for (const view of [target, ...pairs.map(({ view }) => view)]) for (const entry of view.entries()) candidates.add(entry.row * shape.columns + entry.column);
    let extreme = name === 'MINIFS' ? Infinity : -Infinity;
    let found = false;
    const consume = (value: number) => { found = true; extreme = name === 'MINIFS' ? Math.min(extreme, value) : Math.max(extreme, value); };
    if (pairs.every(({ criterion }) => matchesCriteria(null, criterion)) && candidates.size < shape.rows * shape.columns) consume(0);
    for (const key of candidates) {
      const row = Math.floor(key / shape.columns);
      const column = key % shape.columns;
      if (!passes(row, column)) continue;
      const value = target.read(row, column);
      if (isFormulaError(value)) return value;
      if (value === null || typeof value === 'number') consume(value ?? 0);
    }
    return found ? extreme : 0;
  }
  let total = 0;
  let count = 0;
  for (const entry of target.entries()) {
    if (!passes(entry.row, entry.column)) continue;
    if (isFormulaError(entry.value)) return entry.value;
    if (typeof entry.value !== 'number') continue;
    total += entry.value;
    count++;
  }
  return name.startsWith('AVERAGE') ? count ? normalizeExcelPrecision(total / count) : createFormulaError('#DIV/0!', 'No matching numeric values') : normalizeExcelPrecision(total);
}
