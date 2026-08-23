import type { PivotModel, WorksheetModel } from '@react-sheets/core-model';

export interface PivotWriteResult {
  targetStartRow: number;
  targetStartColumn: number;
  values: Array<Array<{ value: string | number | boolean | null }>>;
}

type WriteCell = { value: string | number | boolean | null };

function cellNumber(value: unknown): number {
  if (typeof value === 'number') return value;
  const parsed = Number(String(value ?? '').replace(/[$,%]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

/**
 * 从数据源聚合透视表并生成写回矩阵:
 * 行字段分组 × 列字段分组,值字段按 summarizeBy 聚合。
 * 字段名通过 sourceRange 首行表头定位列偏移;未命中时回退为字段序号。
 */
export function buildPivotWriteback(pivot: PivotModel, sheet: WorksheetModel): PivotWriteResult {
  const source = pivot.sourceRange;

  const headerOffsets = new Map<string, number>();
  for (let c = source.startColumn; c <= source.endColumn; c++) {
    const header = sheet.cells.get(source.startRow, c);
    if (header != null) headerOffsets.set(String(header.value ?? ''), c - source.startColumn);
  }
  const offsetFor = (fields: string[], index: number): number =>
    headerOffsets.get(fields[index] ?? '') ?? index;

  const rowsValues: string[] = [];
  const columnsValues: string[] = [];
  // rowKey -> colKey -> valueField.field -> 聚合值与计数
  const buckets = new Map<string, Map<string, Map<string, { sum: number; count: number }>>>();

  for (let row = source.startRow + 1; row <= source.endRow; row++) {
    const rowKey = pivot.rowFields.length
      ? pivot.rowFields.map((field, i) => String(sheet.cells.get(row, source.startColumn + offsetFor(pivot.rowFields, i))?.value ?? '')).join('|')
      : 'Total';
    if (!rowKey || rowKey === '') continue;
    const columnKey = pivot.columnFields.length
      ? pivot.columnFields.map((field, i) => String(sheet.cells.get(row, source.startColumn + offsetFor(pivot.columnFields, i))?.value ?? '')).join('|')
      : 'Total';
    if (!rowsValues.includes(rowKey)) rowsValues.push(rowKey);
    if (!columnsValues.includes(columnKey)) columnsValues.push(columnKey);

    let rowMap = buckets.get(rowKey);
    if (!rowMap) {
      rowMap = new Map();
      buckets.set(rowKey, rowMap);
    }
    let cellMap = rowMap.get(columnKey);
    if (!cellMap) {
      cellMap = new Map();
      rowMap.set(columnKey, cellMap);
    }
    for (const valueField of pivot.valueFields) {
      const raw = sheet.cells.get(row, source.startColumn + offsetFor([valueField.field], 0));
      const numeric = cellNumber(raw?.value);
      const entry = cellMap.get(valueField.field) ?? { sum: 0, count: 0 };
      entry.sum += numeric;
      entry.count += 1;
      cellMap.set(valueField.field, entry);
    }
  }

  rowsValues.sort((a, b) => a.localeCompare(b));
  columnsValues.sort((a, b) => a.localeCompare(b));

  const summarize = (valueField: { field: string; summarizeBy: string }, entry?: { sum: number; count: number }): number => {
    if (!entry) return 0;
    switch (valueField.summarizeBy) {
      case 'count': return entry.count;
      case 'average': return entry.count === 0 ? 0 : entry.sum / entry.count;
      case 'product': return entry.sum; // 近似:逐项乘积在增量聚合中退化为求和场景极少使用,v1 以和代替
      case 'min':
      case 'max':
      default: return entry.sum;
    }
  };

  const header: WriteCell[] = [
    ...pivot.rowFields.map((field) => ({ value: field })),
    ...columnsValues.flatMap((colKey) =>
      pivot.valueFields.map((valueField) => ({
        value: (columnsValues.length > 1 ? colKey + ' ' : '') + (valueField.displayName ?? valueField.field),
      })),
    ),
  ];

  const values: WriteCell[][] = [header];
  for (const rowKey of rowsValues) {
    const rowCells: WriteCell[] = rowKey.split('|').map((part) => ({ value: part }));
    for (const colKey of columnsValues) {
      for (const valueField of pivot.valueFields) {
        rowCells.push({ value: summarize(valueField, buckets.get(rowKey)?.get(colKey)?.get(valueField.field)) });
      }
    }
    values.push(rowCells);
  }

  const anchor = pivot.targetAnchor ?? { row: source.endRow + 2, column: source.startColumn };
  return {
    targetStartRow: anchor.row,
    targetStartColumn: anchor.column,
    values,
  };
}
