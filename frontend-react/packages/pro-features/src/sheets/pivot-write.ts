import type { PivotModel, WorksheetModel } from '@react-sheets/core-model';

export interface PivotWriteResult {
  targetStartRow: number;
  targetStartColumn: number;
  values: Array<Array<{ value: string | number | boolean | null }>>;
}

function cellNumber(value: unknown): number {
  if (typeof value === 'number') return value;
  const parsed = Number(String(value ?? '').replace(/[$,%]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

/**
 * 从数据源聚合透视表并生成写回矩阵:
 * 行字段分组 × 列字段分组,值字段求和。
 */
export function buildPivotWriteback(pivot: PivotModel, sheet: WorksheetModel): PivotWriteResult {
  const source = pivot.sourceRange;
  const rowsValues: string[] = [];
  const columnsValues: string[] = [];
  const buckets = new Map<string, Map<string, number>>();

  for (let row = source.startRow + 1; row <= source.endRow; row++) {
    const rowKey = pivot.rowFields
      .map((field) => String(sheet.cells.get(row, source.startColumn + field.index)?.value ?? ''))
      .join('|');
    if (!rowKey) continue;
    const columnKey = pivot.columnFields?.length
      ? pivot.columnFields.map((field) => String(sheet.cells.get(row, source.startColumn + field.index)?.value ?? '')).join('|')
      : 'Total';
    const valueCell = sheet.cells.get(row, source.startColumn + pivot.valueField.index);
    if (valueCell == null && !sheet.cells.has(row, source.startColumn + pivot.valueField.index)) continue;
    if (!rowsValues.includes(rowKey)) rowsValues.push(rowKey);
    if (!columnsValues.includes(columnKey)) columnsValues.push(columnKey);
    let rowMap = buckets.get(rowKey);
    if (!rowMap) { rowMap = new Map(); buckets.set(rowKey, rowMap); }
    rowMap.set(columnKey, (rowMap.get(columnKey) ?? 0) + cellNumber(valueCell?.value));
  }

  rowsValues.sort((a, b) => a.localeCompare(b));
  columnsValues.sort((a, b) => a.localeCompare(b));

  const header = [
    ...pivot.rowFields.map((field) => ({ value: field.name })),
    ...(pivot.columnFields?.length ? [] : []),
    ...(columnsValues.length > 1 ? columnsValues : [{ value: pivot.valueField.name }]).map((entry) =>
      typeof entry === 'string' ? { value: entry as string } : entry,
    ),
  ];

  const values: Array<Array<{ value: string | number | boolean | null }>> = [header];
  for (const rowKey of rowsValues) {
    const rowCells: Array<{ value: string | number | boolean | null }> = rowKey
      .split('|')
      .map((part) => ({ value: part }));
    for (const columnKey of columnsValues.length > 1 ? columnsValues : ['Total']) {
      rowCells.push({ value: buckets.get(rowKey)?.get(columnKey) ?? 0 });
    }
    values.push(rowCells);
  }

  return {
    targetStartRow: pivot.targetCell.row,
    targetStartColumn: pivot.targetCell.column,
    values,
  };
}
