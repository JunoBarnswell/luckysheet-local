import type { CellData, PivotModel, RangeRef, WorkbookModel } from '@react-sheets/core-model';

export interface PivotResultTable {
  headers: string[];
  rows: Array<{
    keys: string[];
    values: Array<number | string>;
  }>;
  grandTotal: Array<number | string>;
}

export function computePivotTable(workbook: WorkbookModel, pivot: PivotModel): PivotResultTable {
  const sheet = workbook.getSheet(pivot.sourceRange.sheetId);
  const { startRow, endRow, startColumn, endColumn } = pivot.sourceRange;

  if (startRow >= endRow || startColumn >= endColumn) {
    return { headers: [], rows: [], grandTotal: [] };
  }

  // 1. Read header names
  const headers: string[] = [];
  for (let c = startColumn; c <= endColumn; c++) {
    const cell = sheet.cells.get(startRow, c);
    headers.push(String(cell?.value ?? `Col${c - startColumn + 1}`));
  }

  // 2. Read data rows
  const data: Array<Record<string, unknown>> = [];
  for (let r = startRow + 1; r <= endRow; r++) {
    const rowObj: Record<string, unknown> = {};
    for (let c = startColumn; c <= endColumn; c++) {
      const headerName = headers[c - startColumn]!;
      const cell = sheet.cells.get(r, c);
      rowObj[headerName] = cell?.value ?? null;
    }
    data.push(rowObj);
  }

  // 3. Group by rowFields
  const groups = new Map<string, Array<Record<string, unknown>>>();
  for (const item of data) {
    const groupKey = pivot.rowFields.map((f) => String(item[f] ?? '(blank)')).join(' | ');
    let list = groups.get(groupKey);
    if (!list) {
      list = [];
      groups.set(groupKey, list);
    }
    list.push(item);
  }

  // 4. Aggregate values
  const resultRows: PivotResultTable['rows'] = [];
  for (const [groupKey, items] of groups) {
    const keys = groupKey.split(' | ');
    const values: Array<number | string> = [];

    for (const valField of pivot.valueFields) {
      const fieldName = valField.field;
      const op = valField.summarizeBy;
      const numList: number[] = [];

      for (const item of items) {
        const val = item[fieldName];
        if (typeof val === 'number') numList.push(val);
        else if (typeof val === 'string' && !Number.isNaN(Number(val))) numList.push(Number(val));
      }

      if (op === 'count') {
        values.push(items.length);
      } else if (op === 'sum') {
        values.push(numList.reduce((a, b) => a + b, 0));
      } else if (op === 'average') {
        values.push(numList.length > 0 ? numList.reduce((a, b) => a + b, 0) / numList.length : 0);
      } else if (op === 'min') {
        values.push(numList.length > 0 ? Math.min(...numList) : 0);
      } else if (op === 'max') {
        values.push(numList.length > 0 ? Math.max(...numList) : 0);
      } else if (op === 'product') {
        values.push(numList.length > 0 ? numList.reduce((a, b) => a * b, 1) : 0);
      }
    }

    resultRows.push({ keys, values });
  }

  // 5. Grand Totals
  const grandTotal: Array<number | string> = [];
  for (const valField of pivot.valueFields) {
    const fieldName = valField.field;
    const op = valField.summarizeBy;
    const numList: number[] = [];

    for (const item of data) {
      const val = item[fieldName];
      if (typeof val === 'number') numList.push(val);
      else if (typeof val === 'string' && !Number.isNaN(Number(val))) numList.push(Number(val));
    }

    if (op === 'count') {
      grandTotal.push(data.length);
    } else if (op === 'sum') {
      grandTotal.push(numList.reduce((a, b) => a + b, 0));
    } else if (op === 'average') {
      grandTotal.push(numList.length > 0 ? numList.reduce((a, b) => a + b, 0) / numList.length : 0);
    } else if (op === 'min') {
      grandTotal.push(numList.length > 0 ? Math.min(...numList) : 0);
    } else if (op === 'max') {
      grandTotal.push(numList.length > 0 ? Math.max(...numList) : 0);
    }
  }

  const resultHeaders = [
    ...pivot.rowFields,
    ...pivot.valueFields.map((v) => v.displayName || `${v.summarizeBy.toUpperCase()} of ${v.field}`),
  ];

  return {
    headers: resultHeaders,
    rows: resultRows,
    grandTotal,
  };
}
