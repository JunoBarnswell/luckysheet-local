import type { PivotModel, WorkbookModel } from '@react-sheets/core-model';
import { computePivotResult, normalizePivotDefinition } from './engine';

export interface PivotWriteResult {
  targetStartRow: number;
  targetStartColumn: number;
  values: Array<Array<{ value: string | number | boolean | null }>>;
}

/**
 * Compatibility/export projection only. The normal worksheet render path uses
 * buildPivotGridProjection and never calls this function or writes its values
 * into WorksheetModel.cells.
 */
export function buildPivotWriteback(pivot: PivotModel, workbook: WorkbookModel): PivotWriteResult {
  const definition = normalizePivotDefinition(workbook, pivot);
  const tree = computePivotResult(workbook, pivot);
  const values = definition.layout.values;
  const output: Array<Array<{ value: string | number | boolean | null }>> = [];
  output.push(tree.columnPaths.flatMap((path) => values.map((field) => ({ value: `${path.map((item) => item == null ? '(blank)' : String(item)).join(' / ')} ${field.displayName ?? field.fieldId ?? field.field ?? ''}`.trim() }))));
  for (const node of tree.rows) output.push([{ value: node.label }, ...node.values.flatMap((item) => item.values.map((value) => ({ value }))) ]);
  if (tree.grandTotal) output.push([{ value: 'Grand Total' }, ...tree.grandTotal.values.map((value) => ({ value }))]);
  return { targetStartRow: definition.target.anchor.row, targetStartColumn: definition.target.anchor.column, values: output };
}
