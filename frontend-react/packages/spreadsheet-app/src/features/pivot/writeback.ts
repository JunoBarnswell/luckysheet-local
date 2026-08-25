import { formatPivotMember, type PivotModel, type PivotScalar, type WorkbookModel } from '@react-sheets/core-model';
import { computePivotResult, normalizePivotDefinition } from './engine';

export interface PivotWriteResult {
  targetStartRow: number;
  targetStartColumn: number;
  values: Array<Array<{ value: PivotScalar }>>;
}

/**
 * Export projection only. The normal worksheet render path uses
 * buildPivotGridProjection and never calls this function or writes its values
 * into WorksheetModel.cells.
 */
export function buildPivotWriteback(pivot: PivotModel, workbook: WorkbookModel): PivotWriteResult {
  const definition = normalizePivotDefinition(workbook, pivot);
  const tree = computePivotResult(workbook, pivot);
  const values = definition.layout.values;
  const output: Array<Array<{ value: PivotScalar }>> = [];
  output.push([
    ...tree.columnPaths.flatMap((path) => values.map((field) => ({ value: `${path.map(formatPivotMember).join(' / ')} ${field.displayName ?? field.fieldId}`.trim() }))),
    ...(definition.layout.showRowGrandTotals ? values.map((field) => ({ value: `Grand Total ${field.displayName ?? field.fieldId}` })) : []),
  ]);
  for (const node of tree.rows) output.push([
    { value: node.label },
    ...node.values.flatMap((item) => item.values.map((value) => ({ value }))),
    ...(definition.layout.showRowGrandTotals ? (node.rowGrandTotal?.values ?? values.map(() => null)).map((value) => ({ value })) : []),
  ]);
  if (tree.grandTotal && definition.layout.showColumnGrandTotals) output.push([
    { value: 'Grand Total' },
    ...(tree.columnGrandTotals ?? [tree.grandTotal]).flatMap((cell) => cell.values.map((value) => ({ value }))),
    ...(definition.layout.showRowGrandTotals ? tree.grandTotal.values.map((value) => ({ value })) : []),
  ]);
  return { targetStartRow: definition.target.anchor.row, targetStartColumn: definition.target.anchor.column, values: output };
}
