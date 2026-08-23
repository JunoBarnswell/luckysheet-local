import {
  createFormulaError,
  isFormulaError,
  type FormulaError,
  type FormulaValue,
  type ScalarValue,
} from '../values';

type Matrix = readonly (readonly ScalarValue[])[];
type Aggregate = 'SUM' | 'COUNT' | 'AVERAGE' | 'MIN' | 'MAX';

interface Group {
  readonly key: string;
  readonly fields: readonly ScalarValue[];
  readonly rows: ScalarValue[][];
}

/**
 * Supported GROUPBY form:
 *   GROUPBY(row_fields, values, aggregation)
 *
 * Supported PIVOTBY form:
 *   PIVOTBY(row_fields, column_fields, values, aggregation)
 *
 * Field and value arrays must have matching row counts. Multiple field and
 * value columns are supported. Optional Excel layout/total/sort parameters
 * are deliberately rejected rather than silently ignored.
 */
function GROUPBY(args: FormulaValue[]): FormulaValue {
  if (args.length !== 3) {
    return createFormulaError('#VALUE!', 'GROUPBY supports row_fields, values, and aggregation');
  }
  const rowFields = matrixFrom(args[0], 'GROUPBY row_fields');
  if (isFormulaError(rowFields)) return rowFields;
  const values = matrixFrom(args[1], 'GROUPBY values');
  if (isFormulaError(values)) return values;
  const aggregation = aggregateFrom(args[2]);
  if (isFormulaError(aggregation)) return aggregation;
  const compatible = assertMatchingRows(rowFields, values, 'GROUPBY');
  if (isFormulaError(compatible)) return compatible;

  const groups = groupRows(rowFields, values);
  if (isFormulaError(groups)) return groups;
  return groups.map((group) => [
    ...group.fields,
    ...aggregateColumns(group.rows, aggregation),
  ]);
}

function PIVOTBY(args: FormulaValue[]): FormulaValue {
  if (args.length !== 4) {
    return createFormulaError('#VALUE!', 'PIVOTBY supports row_fields, column_fields, values, and aggregation');
  }
  const rowFields = matrixFrom(args[0], 'PIVOTBY row_fields');
  if (isFormulaError(rowFields)) return rowFields;
  const columnFields = matrixFrom(args[1], 'PIVOTBY column_fields');
  if (isFormulaError(columnFields)) return columnFields;
  const values = matrixFrom(args[2], 'PIVOTBY values');
  if (isFormulaError(values)) return values;
  const aggregation = aggregateFrom(args[3]);
  if (isFormulaError(aggregation)) return aggregation;

  const rowCompatibility = assertMatchingRows(rowFields, values, 'PIVOTBY row_fields');
  if (isFormulaError(rowCompatibility)) return rowCompatibility;
  const columnCompatibility = assertMatchingRows(columnFields, values, 'PIVOTBY column_fields');
  if (isFormulaError(columnCompatibility)) return columnCompatibility;

  const rowGroups = groupRows(rowFields, values);
  if (isFormulaError(rowGroups)) return rowGroups;
  const columnGroups = groupRows(columnFields, values);
  if (isFormulaError(columnGroups)) return columnGroups;
  const cells = pivotCells(rowFields, columnFields, values);
  if (isFormulaError(cells)) return cells;

  const valueWidth = values[0]?.length ?? 0;
  const headers = pivotHeaders(
    rowFields[0]?.length ?? 0,
    columnFields[0]?.length ?? 0,
    valueWidth,
    columnGroups,
    aggregation,
  );
  const body: FormulaValue[][] = rowGroups.map((rowGroup) => {
    const row: FormulaValue[] = [...rowGroup.fields];
    for (const columnGroup of columnGroups) {
      const sourceRows = cells.get(pivotCellKey(rowGroup.key, columnGroup.key)) ?? [];
      for (let column = 0; column < valueWidth; column += 1) {
        row.push(aggregate(sourceRows.map((source) => source[column] ?? null), aggregation));
      }
    }
    return row;
  });
  return [...headers, ...body];
}

function matrixFrom(value: FormulaValue | undefined, label: string): Matrix | FormulaError {
  if (value === undefined || isFormulaError(value)) {
    return isFormulaError(value)
      ? value
      : createFormulaError('#VALUE!', `${label} is required`);
  }
  if (!Array.isArray(value)) {
    const scalar = scalarFrom(value, label);
    return isFormulaError(scalar) ? scalar : [[scalar]];
  }
  if (value.length === 0) return createFormulaError('#VALUE!', `${label} cannot be empty`);

  const rows: ScalarValue[][] = [];
  let width: number | undefined;
  for (const candidate of value) {
    if (!Array.isArray(candidate)) return createFormulaError('#VALUE!', `${label} must be rectangular`);
    if (width === undefined) {
      width = candidate.length;
      if (width === 0) return createFormulaError('#VALUE!', `${label} cannot have empty rows`);
    } else if (candidate.length !== width) {
      return createFormulaError('#VALUE!', `${label} must be rectangular`);
    }
    const row: ScalarValue[] = [];
    for (const cell of candidate) {
      const scalar = scalarFrom(cell, label);
      if (isFormulaError(scalar)) return scalar;
      row.push(scalar);
    }
    rows.push(row);
  }
  return rows;
}

function scalarFrom(value: FormulaValue, label: string): ScalarValue | FormulaError {
  if (isFormulaError(value)) return value;
  if (Array.isArray(value)) return createFormulaError('#VALUE!', `${label} contains a nested array`);
  if (typeof value === 'number' && !Number.isFinite(value)) {
    return createFormulaError('#VALUE!', `${label} contains a non-finite number`);
  }
  return value;
}

function assertMatchingRows(left: Matrix, right: Matrix, label: string): true | FormulaError {
  return left.length === right.length
    ? true
    : createFormulaError('#VALUE!', `${label} arrays must have the same row count`);
}

function aggregateFrom(value: FormulaValue | undefined): Aggregate | FormulaError {
  if (isFormulaError(value)) return value;
  if (typeof value === 'number') {
    switch (value) {
      case 1: return 'SUM';
      case 2: return 'COUNT';
      case 3: return 'AVERAGE';
      case 4: return 'MIN';
      case 5: return 'MAX';
      default: return createFormulaError('#VALUE!', 'Unsupported aggregate code');
    }
  }
  if (typeof value === 'string') {
    const normalized = value.trim().toUpperCase();
    if (normalized === 'SUM' || normalized === 'COUNT' || normalized === 'AVERAGE' || normalized === 'MIN' || normalized === 'MAX') {
      return normalized;
    }
  }
  return createFormulaError('#VALUE!', 'Aggregation must be SUM, COUNT, AVERAGE, MIN, MAX, or a supported aggregate code');
}

function groupRows(fields: Matrix, values: Matrix): Group[] | FormulaError {
  const groups = new Map<string, Group>();
  for (let row = 0; row < fields.length; row += 1) {
    const fieldRow = fields[row];
    const valueRow = values[row];
    if (!fieldRow || !valueRow) return createFormulaError('#VALUE!', 'Field and value arrays must have the same row count');
    const key = tupleKey(fieldRow);
    const group = groups.get(key);
    if (group) {
      group.rows.push([...valueRow]);
      continue;
    }
    groups.set(key, {
      key,
      fields: [...fieldRow],
      rows: [[...valueRow]],
    });
  }
  return [...groups.values()];
}

function pivotCells(
  rowFields: Matrix,
  columnFields: Matrix,
  values: Matrix,
): Map<string, ScalarValue[][]> | FormulaError {
  const cells = new Map<string, ScalarValue[][]>();
  for (let row = 0; row < values.length; row += 1) {
    const rowField = rowFields[row];
    const columnField = columnFields[row];
    const valueRow = values[row];
    if (!rowField || !columnField || !valueRow) {
      return createFormulaError('#VALUE!', 'PIVOTBY arrays must have the same row count');
    }
    const key = pivotCellKey(tupleKey(rowField), tupleKey(columnField));
    const entries = cells.get(key) ?? [];
    entries.push([...valueRow]);
    cells.set(key, entries);
  }
  return cells;
}

function aggregateColumns(rows: readonly (readonly ScalarValue[])[], aggregation: Aggregate): FormulaValue[] {
  const width = rows[0]?.length ?? 0;
  const values: FormulaValue[] = [];
  for (let column = 0; column < width; column += 1) {
    values.push(aggregate(rows.map((row) => row[column] ?? null), aggregation));
  }
  return values;
}

function aggregate(values: readonly ScalarValue[], aggregation: Aggregate): FormulaValue {
  const numbers = values.filter((value): value is number => typeof value === 'number' && Number.isFinite(value));
  switch (aggregation) {
    case 'SUM': return numbers.reduce((sum, value) => sum + value, 0);
    case 'COUNT': return numbers.length;
    case 'AVERAGE': {
      if (numbers.length === 0) return createFormulaError('#DIV/0!', 'AVERAGE has no numeric values');
      return numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
    }
    case 'MIN': return numbers.length === 0 ? 0 : Math.min(...numbers);
    case 'MAX': return numbers.length === 0 ? 0 : Math.max(...numbers);
  }
}

function pivotHeaders(
  rowFieldWidth: number,
  columnFieldWidth: number,
  valueWidth: number,
  columnGroups: readonly Group[],
  aggregation: Aggregate,
): FormulaValue[][] {
  const headers: FormulaValue[][] = [];
  for (let field = 0; field < columnFieldWidth; field += 1) {
    const header: FormulaValue[] = Array.from({ length: rowFieldWidth }, () => null);
    for (const columnGroup of columnGroups) {
      for (let value = 0; value < valueWidth; value += 1) {
        header.push(columnGroup.fields[field] ?? null);
      }
    }
    headers.push(header);
  }
  if (valueWidth > 1) {
    const valueHeader: FormulaValue[] = Array.from({ length: rowFieldWidth }, () => null);
    for (const _columnGroup of columnGroups) {
      for (let value = 0; value < valueWidth; value += 1) {
        valueHeader.push(`${aggregation} ${value + 1}`);
      }
    }
    headers.push(valueHeader);
  }
  return headers;
}

function tupleKey(values: readonly ScalarValue[]): string {
  return values.map(keyPart).join('\u001F');
}

function keyPart(value: ScalarValue): string {
  if (value === null || value === '') return 'blank';
  if (typeof value === 'number') return `number:${Object.is(value, -0) ? '-0' : String(value)}`;
  if (typeof value === 'boolean') return value ? 'boolean:true' : 'boolean:false';
  return `string:${JSON.stringify(value)}`;
}

function pivotCellKey(rowKey: string, columnKey: string): string {
  return `${rowKey}\u001E${columnKey}`;
}

export const extendedMatrixFunctions: Record<string, (args: FormulaValue[]) => FormulaValue> = {
  GROUPBY,
  PIVOTBY,
};
