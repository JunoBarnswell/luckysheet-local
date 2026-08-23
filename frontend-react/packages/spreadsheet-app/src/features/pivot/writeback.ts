import type { PivotModel, WorkbookModel } from '@react-sheets/core-model';
import { computePivotResult } from './engine';

export interface PivotWriteResult { targetStartRow: number; targetStartColumn: number; values: Array<Array<{ value: string | number | boolean | null }>>; }

/** Export-only projection; aggregation remains exclusively in the result-tree engine. */
export function buildPivotWriteback(pivot: PivotModel, workbook: WorkbookModel): PivotWriteResult {
  const tree = computePivotResult(workbook, pivot); const values = pivot.layout.values; const output: Array<Array<{ value: string | number | boolean | null }>> = [];
  output.push(tree.columnPaths.flatMap((path) => values.map((field) => ({ value: `${path.map((item) => item == null ? '(blank)' : String(item)).join(' / ')} ${field.displayName ?? field.field}`.trim() }))));
  for (const node of tree.rows) output.push([{ value: node.label }, ...node.values.flatMap((item) => item.values.map((value) => ({ value })))]);
  if (tree.grandTotal) output.push([{ value: 'Grand Total' }, ...tree.grandTotal.values.map((value) => ({ value }))]);
  const anchor = pivot.targetAnchor ?? { row: pivot.sourceRange.endRow + 2, column: pivot.sourceRange.startColumn }; return { targetStartRow: anchor.row, targetStartColumn: anchor.column, values: output };
}
