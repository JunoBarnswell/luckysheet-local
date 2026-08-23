import type { CellData, RangeRef, TableScalar, WorkbookModel, WorkbookTableModel } from '@react-sheets/core-model';
import { serializeQueryDefinition, type ConnectorRegistry, type QueryResult } from './index';
import {
  QueryStepPipeline,
  validateQuerySteps,
  type LoadTarget,
  type QueryDefinition,
} from './query-steps';

export interface QueryResultSnapshot {
  queryId: string;
  queryName: string;
  columns: string[];
  rowCount: number;
  loadedAt: string;
  target: LoadTarget;
  sourceRevision: number;
  persistedDefinition?: import('./index').QueryDefinitionPersistence;
}

export interface QuerySessionEntry {
  definition: QueryDefinition;
  lastResult?: QueryResultSnapshot;
}

export interface QueryLoadCommandPayload {
  query: QueryDefinition;
  target: LoadTarget;
  result: QueryResult;
}

export interface QueryCellLoadPayload {
  kind: 'cells';
  queryId: string;
  target: LoadTarget;
  clearRange: RangeRef;
  values: CellData[][];
  /** Pivot source loads refresh this persisted pivot metadata atomically. */
  pivot?: { sheetId: string; pivotId: string; nextRefreshRevision: number; nextRefreshedAt: string };
}

export interface QueryWorkbookTableLoadPayload {
  kind: 'workbook-table';
  queryId: string;
  tableId: string;
  table: WorkbookTableModel;
  result: QueryResult;
  sourceRevision: number;
}

export type QueryLoadMutationPayload = QueryCellLoadPayload | QueryWorkbookTableLoadPayload;

export interface WorkbookTableQueryRecord {
  result: QueryResult;
  sourceRevision: number;
}

export interface WorkbookTableQueryStore {
  get(tableId: string): WorkbookTableQueryRecord | undefined;
  set(tableId: string, record: WorkbookTableQueryRecord): void;
  delete(tableId: string): void;
}

/** Default local store for the columnar WorkbookTableModel data plane. */
export class InMemoryWorkbookTableQueryStore implements WorkbookTableQueryStore {
  private readonly records = new Map<string, WorkbookTableQueryRecord>();

  get(tableId: string): WorkbookTableQueryRecord | undefined {
    const record = this.records.get(tableId);
    return record ? structuredClone(record) : undefined;
  }

  set(tableId: string, record: WorkbookTableQueryRecord): void {
    this.records.set(tableId, structuredClone(record));
  }

  delete(tableId: string): void {
    this.records.delete(tableId);
  }
}

export function queryResultToRangeValues(result: QueryResult): CellData[][] {
  return [
    result.columns.map((name) => ({ value: name })),
    ...result.rows.map((row) => row.map((value) => ({ value: value ?? '' }))),
  ];
}

export async function executeQueryDefinition(
  connectors: ConnectorRegistry,
  query: QueryDefinition,
): Promise<QueryResult> {
  validateQueryDefinition(query);
  const connector = connectors.get(query.connectorId);
  await connector.connect(query.connectorConfig);
  try {
    const raw = await connector.executeQuery(query.id);
    validateQueryResult(raw);
    const pipeline = new QueryStepPipeline(query.steps);
    const transformed = pipeline.applySteps({
      columns: raw.columns,
      rows: raw.rows,
    });
    return {
      columns: transformed.columns,
      rows: transformed.rows as QueryResult['rows'],
      rowCount: transformed.rows.length,
    };
  } finally {
    await connector.disconnect();
  }
}

export function resolveLoadTarget(activeSheetId: string, selectionRange: RangeRef): LoadTarget {
  return {
    kind: 'range',
    sheetId: activeSheetId,
    range: {
      startRow: selectionRange.startRow,
      startColumn: selectionRange.startColumn,
      endRow: selectionRange.endRow,
      endColumn: selectionRange.endColumn,
    },
  };
}

export function buildQueryResultSnapshot(
  query: QueryDefinition,
  result: QueryResult,
  target: LoadTarget,
): QueryResultSnapshot {
  return {
    queryId: query.id,
    queryName: query.name,
    columns: [...result.columns],
    rowCount: result.rowCount,
    loadedAt: new Date().toISOString(),
    target,
    sourceRevision: query.sourceRevision ?? 0,
    persistedDefinition: serializeQueryDefinition(query),
  };
}

export function summarizeQueryResult(snapshot: QueryResultSnapshot): string {
  return `Loaded ${snapshot.rowCount} rows × ${snapshot.columns.length} columns into the sheet`;
}

export function createInlineJsonQuery(
  id: string,
  name: string,
  data: Record<string, unknown>[],
  steps: QueryDefinition['steps'] = [],
): QueryDefinition {
  return {
    id,
    name,
    connectorId: 'json',
    connectorConfig: { data },
    steps,
  };
}

export function validateQueryDefinition(query: QueryDefinition): void {
  if (!query || typeof query !== 'object') throw new Error('Query definition is required');
  if (typeof query.id !== 'string' || !query.id.trim()) throw new Error('Query id is required');
  if (typeof query.name !== 'string' || !query.name.trim()) throw new Error('Query name is required');
  if (typeof query.connectorId !== 'string' || !query.connectorId.trim()) throw new Error('Query connectorId is required');
  if (!query.connectorConfig || typeof query.connectorConfig !== 'object' || Array.isArray(query.connectorConfig)) {
    throw new Error(`Query ${query.id} has invalid connector configuration`);
  }
  if (!Array.isArray(query.steps)) throw new Error(`Query ${query.id} steps must be an array`);
  validateQuerySteps(query.steps);
  if (query.sourceRevision !== undefined && (!Number.isInteger(query.sourceRevision) || query.sourceRevision < 0)) {
    throw new Error(`Query ${query.id} sourceRevision must be a non-negative integer`);
  }
  if (query.refreshPolicy) {
    const intervalMs = query.refreshPolicy.intervalMs;
    const validInterval = typeof intervalMs === 'number' && Number.isInteger(intervalMs) && intervalMs > 0;
    if (query.refreshPolicy.mode === 'interval' && !validInterval) {
      throw new Error(`Query ${query.id} interval refresh requires a positive intervalMs`);
    }
    if (query.refreshPolicy.mode !== 'manual' && query.refreshPolicy.mode !== 'on-open' && query.refreshPolicy.mode !== 'interval') {
      throw new Error(`Query ${query.id} has an unknown refresh policy`);
    }
  }
}

export function queryResultToScalarMatrix(result: QueryResult): TableScalar[][] {
  return result.rows.map((row) => row.map((value) => value ?? null));
}

export function buildQueryCellValues(result: QueryResult, includeHeaders = true): CellData[][] {
  const rows: CellData[][] = [];
  if (includeHeaders) rows.push(result.columns.map((name) => ({ value: name })));
  rows.push(...result.rows.map((row) => row.map((value) => ({ value: value ?? null }))));
  return rows;
}

function assertRangeBounds(workbook: WorkbookModel, range: RangeRef): void {
  const sheet = workbook.getSheet(range.sheetId);
  if (range.startRow < 0 || range.startColumn < 0 || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) {
    throw new Error(`Query load range is outside worksheet bounds: ${range.sheetId}`);
  }
}

function findPivot(workbook: WorkbookModel, pivotId: string): { sheetId: string; pivot: import('@react-sheets/core-model').PivotModel } | undefined {
  for (const sheet of workbook.getSheets()) {
    const pivot = sheet.pivots.find((candidate) => candidate.id === pivotId);
    if (pivot) return { sheetId: sheet.id, pivot };
  }
  return undefined;
}

function sourceRangeForPivot(pivot: import('@react-sheets/core-model').PivotModel): RangeRef {
  if (pivot.dataSource?.kind === 'worksheet-ranges') {
    throw new Error(`Pivot ${pivot.id} has multiple source ranges; query load requires an explicit range target`);
  }
  return structuredClone(pivot.dataSource?.range ?? pivot.sourceRange);
}

function previousCells(workbook: WorkbookModel, range: RangeRef): Array<{ row: number; column: number; value?: CellData }> {
  const sheet = workbook.getSheet(range.sheetId);
  const cells: Array<{ row: number; column: number; value?: CellData }> = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const value = sheet.cells.get(row, column);
      cells.push({ row, column, value: value ? structuredClone(value) : undefined });
    }
  }
  return cells;
}

export interface QueryLoadPlan {
  mutationId: 'query.load.range' | 'query.load.sheet-table' | 'query.load.workbook-table' | 'query.load.pivot-source';
  payload: QueryLoadMutationPayload;
  inverse: QueryLoadMutationPayload;
  affectedRanges: RangeRef[];
}

export function buildQueryLoadPlan(
  workbook: WorkbookModel,
  params: QueryLoadCommandPayload,
  tableStore: WorkbookTableQueryStore,
): QueryLoadPlan {
  validateQueryDefinition(params.query);
  if (!params.result || !Array.isArray(params.result.columns) || !Array.isArray(params.result.rows)) {
    throw new Error('Query result is invalid');
  }
  validateQueryResult(params.result);
  const target = params.target;
  if (target.kind === 'range') {
    if (!target.sheetId || !target.range) throw new Error('Range query target requires sheetId and range');
    const startRow = target.range.startRow;
    const startColumn = target.range.startColumn;
    const values = buildQueryCellValues(params.result);
    const endRow = Math.max(target.range.endRow ?? startRow, startRow + values.length - 1);
    const endColumn = Math.max(target.range.endColumn ?? startColumn, startColumn + Math.max(params.result.columns.length, 1) - 1);
    const clearRange: RangeRef = { sheetId: target.sheetId, startRow, endRow, startColumn, endColumn };
    assertRangeBounds(workbook, clearRange);
    const payload: QueryCellLoadPayload = { kind: 'cells', queryId: params.query.id, target: structuredClone(target), clearRange, values };
    const inverse: QueryCellLoadPayload = {
      kind: 'cells',
      queryId: params.query.id,
      target: structuredClone(target),
      clearRange,
      values: [],
    };
    Object.assign(inverse, { previousCells: previousCells(workbook, clearRange) });
    return { mutationId: 'query.load.range', payload, inverse, affectedRanges: [clearRange] };
  }

  if (target.kind === 'sheet-table') {
    if (!target.sheetId || !target.tableId) throw new Error('Sheet-table query target requires sheetId and tableId');
    const sheet = workbook.getSheet(target.sheetId);
    const table = sheet.sheetTables.find((entry) => entry.id === target.tableId);
    if (!table) throw new Error(`Unknown sheet table: ${target.tableId}`);
    if (params.result.columns.length > table.columns.length) throw new Error(`Query result has too many columns for table ${table.name}`);
    const tableRange = structuredClone(table.range);
    const headerRows = table.hasHeaderRow ? 1 : 0;
    const totalRows = table.hasTotalRow ? 1 : 0;
    const capacity = tableRange.endRow - tableRange.startRow + 1 - headerRows - totalRows;
    if (params.result.rows.length > capacity) throw new Error(`Query result has too many rows for table ${table.name}`);
    const values = buildQueryCellValues(params.result, table.hasHeaderRow);
    const loadRange: RangeRef = { ...tableRange, endRow: tableRange.endRow - totalRows };
    const payload: QueryCellLoadPayload = {
      kind: 'cells', queryId: params.query.id, target: structuredClone(target), clearRange: loadRange, values,
    };
    const inverse: QueryCellLoadPayload = {
      kind: 'cells', queryId: params.query.id, target: structuredClone(target), clearRange: loadRange, values: [],
    };
    Object.assign(inverse, { previousCells: previousCells(workbook, loadRange) });
    return { mutationId: 'query.load.sheet-table', payload, inverse, affectedRanges: [tableRange] };
  }

  if (target.kind === 'workbook-table') {
    if (!target.tableId) throw new Error('Workbook-table query target requires tableId');
    const table = workbook.tables.get(target.tableId);
    if (!table) throw new Error(`Unknown workbook table: ${target.tableId}`);
    const previous = tableStore.get(target.tableId);
    const nextTable: WorkbookTableModel = {
      ...structuredClone(table),
      rowCount: params.result.rowCount,
      fields: params.result.columns.map((name, ordinal) => ({
        id: table.fields[ordinal]?.id ?? `${table.id}:field:${ordinal}`,
        name,
        ordinal,
        type: inferTableFieldType(params.result.rows.map((row) => row[ordinal] ?? null)),
      })),
      blocks: [{
        id: `${table.id}:query:${params.query.id}`,
        tableId: table.id,
        startRow: 0,
        rowCount: params.result.rowCount,
        storageKey: `query:${params.query.id}:${params.query.sourceRevision ?? 0}`,
        encoding: 'typed-column-v1',
      }],
      revision: table.revision + 1,
    };
    const payload: QueryWorkbookTableLoadPayload = {
      kind: 'workbook-table', queryId: params.query.id, tableId: table.id, table: nextTable,
      result: structuredClone(params.result), sourceRevision: params.query.sourceRevision ?? 0,
    };
    const inverse: QueryWorkbookTableLoadPayload = {
      kind: 'workbook-table', queryId: params.query.id, tableId: table.id, table: structuredClone(table),
      result: previous?.result ?? { columns: [], rows: [], rowCount: 0 }, sourceRevision: previous?.sourceRevision ?? 0,
    };
    Object.assign(inverse, { previousTable: structuredClone(table), previousRecord: previous });
    return { mutationId: 'query.load.workbook-table', payload, inverse, affectedRanges: [] };
  }

  if (target.kind === 'pivot-source') {
    if (!target.pivotId) throw new Error('Pivot-source query target requires pivotId');
    const found = findPivot(workbook, target.pivotId);
    if (!found) throw new Error(`Unknown pivot: ${target.pivotId}`);
    const sourceRange = target.range
      ? { sheetId: target.sheetId ?? sourceRangeForPivot(found.pivot).sheetId, startRow: target.range.startRow, endRow: target.range.endRow ?? target.range.startRow + params.result.rows.length, startColumn: target.range.startColumn, endColumn: target.range.endColumn ?? target.range.startColumn + Math.max(params.result.columns.length, 1) - 1 }
      : sourceRangeForPivot(found.pivot);
    const values = buildQueryCellValues(params.result);
    if (values.length > sourceRange.endRow - sourceRange.startRow + 1 || Math.max(params.result.columns.length, 1) > sourceRange.endColumn - sourceRange.startColumn + 1) {
      throw new Error(`Query result does not fit pivot ${target.pivotId} source range`);
    }
    assertRangeBounds(workbook, sourceRange);
    const payload: QueryCellLoadPayload = {
      kind: 'cells', queryId: params.query.id, target: structuredClone(target), clearRange: sourceRange, values,
      pivot: { sheetId: found.sheetId, pivotId: found.pivot.id, nextRefreshRevision: (found.pivot.refreshRevision ?? 0) + 1, nextRefreshedAt: new Date().toISOString() },
    };
    const inverse: QueryCellLoadPayload = {
      kind: 'cells', queryId: params.query.id, target: structuredClone(target), clearRange: sourceRange, values: [],
      pivot: { sheetId: found.sheetId, pivotId: found.pivot.id, nextRefreshRevision: found.pivot.refreshRevision ?? 0, nextRefreshedAt: found.pivot.lastRefreshedAt ?? '' },
    };
    Object.assign(inverse, { previousCells: previousCells(workbook, sourceRange) });
    return { mutationId: 'query.load.pivot-source', payload, inverse, affectedRanges: [sourceRange] };
  }

  throw new Error(`Unsupported query load target: ${String((target as { kind?: unknown }).kind)}`);
}

function inferTableFieldType(values: TableScalar[]): WorkbookTableModel['fields'][number]['type'] {
  if (values.every((value) => value == null || typeof value === 'number')) return 'number';
  if (values.every((value) => value == null || typeof value === 'boolean')) return 'boolean';
  if (values.every((value) => value == null || typeof value === 'string')) return 'text';
  return 'mixed';
}

function validateQueryResult(result: QueryResult): void {
  if (result.rowCount !== result.rows.length) throw new Error('Query result rowCount does not match rows');
  if (result.columns.some((column) => typeof column !== 'string' || !column.trim())) throw new Error('Query result columns must be non-empty strings');
  const names = new Set<string>();
  for (const column of result.columns) {
    if (names.has(column)) throw new Error(`Query result contains duplicate column "${column}"`);
    names.add(column);
  }
  for (const row of result.rows) {
    if (row.length !== result.columns.length) throw new Error('Query result row width does not match columns');
    for (const value of row) {
      if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
        throw new Error('Query result contains a non-scalar value');
      }
    }
  }
}
