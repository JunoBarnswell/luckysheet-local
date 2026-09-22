import type {
  DataBlockRef,
  DataSourceField,
  DataSourceFieldType,
  DataSourceManifest,
  PivotSource,
  RangeRef,
  SheetDataRegion,
  TableScalar,
  WorkbookModel,
  WorkbookTableModel,
} from '@react-sheets/core-model';
import { DEFAULT_DATA_BLOCK_ROW_COUNT, MAX_SHEET_COLUMN_COUNT, MAX_SHEET_ROW_COUNT, PIVOT_DAY_MS, pivotTimelineInstant } from '@react-sheets/core-model';
import {
  COLUMNAR_BLOCK_ENCODING,
  computeColumnarBlockChecksum,
  encodeColumnarBlock,
  type ColumnarBlockField,
} from '../data-source/codec';
import { serializeQueryDefinition, type ConnectorRegistry, type QueryDefinitionPersistence, type QueryResult } from './index';
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
  persistedDefinition?: QueryDefinitionPersistence;
}

/** Read-only result projection used before a query is loaded into a worksheet. */
export interface QueryPreview {
  queryId: string;
  queryName: string;
  columns: string[];
  rowCount: number;
  sampleRows: TableScalar[][];
}

export interface QuerySessionEntry {
  definition: QueryDefinition;
  lastResult?: QueryResultSnapshot;
}

export interface QuerySheetBinding {
  kind: 'sheet-region';
  region: SheetDataRegion;
  /** Header labels are small metadata; result rows stay in immutable blocks. */
  header: TableScalar[];
}

export interface QueryWorkbookTableBinding {
  kind: 'workbook-table';
  tableId: string;
  table: WorkbookTableModel;
}

export type QueryLoadBinding = QuerySheetBinding | QueryWorkbookTableBinding;

export interface QueryLoadExtent {
  sheetId: string;
  rowCount: number;
  columnCount: number;
}

/** The query-load mutation contains metadata only; block bytes are separate. */
export interface QueryLoadCommandPayload {
  kind: 'data-source-load';
  queryId: string;
  queryDefinition: QueryDefinitionPersistence | null;
  target: LoadTarget;
  sourceId: string;
  source: DataSourceManifest;
  binding: QueryLoadBinding;
  extent?: QueryLoadExtent;
  /** Pivot-source loads switch the Pivot to the same block-backed source. */
  pivotSource?: PivotSource;
}

export interface QueryLoadRestorePayload extends Omit<QueryLoadCommandPayload, 'source' | 'binding'> {
  source: DataSourceManifest | null;
  binding: QueryLoadBinding | null;
}

export type QueryLoadMutationPayload = QueryLoadCommandPayload | QueryLoadRestorePayload;

export interface PreparedQueryLoad {
  payload: QueryLoadCommandPayload;
  blocks: Array<{ ref: DataBlockRef; payload: ArrayBuffer }>;
}

export async function executeQueryDefinition(
  connectors: ConnectorRegistry,
  query: QueryDefinition,
): Promise<QueryResult> {
  validateQueryDefinition(query);
  const serverOnlyConnectors = new Set(['rest', 'sqlite', 'jdbc']);
  if (serverOnlyConnectors.has(query.connectorId)) throw new Error(`Connector ${query.connectorId} is server-only and cannot execute in the local workbook`);
  const definition = structuredClone(query);
  return connectors.withConnector(definition.connectorId, async (connector) => {
    if (connector.execution !== 'local') throw new Error(`Connector ${connector.id} is server-only and cannot execute in the local workbook`);
    const failures: unknown[] = [];
    try {
      await connector.connect(definition.connectorConfig);
      const raw = await connector.executeQuery(definition.id);
      validateQueryResult(raw);
      const transformed = new QueryStepPipeline(definition.steps).applySteps({ columns: raw.columns, rows: raw.rows });
      return { columns: transformed.columns, rows: transformed.rows as QueryResult['rows'], rowCount: transformed.rows.length };
    } catch (error) {
      failures.push(error);
      throw error;
    } finally {
      try { await connector.disconnect(); }
      catch (error) {
        if (failures.length > 0) throw new AggregateError([...failures, error], `Query ${definition.id} execution and disconnect both failed`);
        throw error;
      }
    }
  });
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

export function buildQueryResultSnapshot(query: QueryDefinition, result: QueryResult, target: LoadTarget): QueryResultSnapshot {
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

export function buildQueryPreview(query: QueryDefinition, result: QueryResult, sampleSize = 5): QueryPreview {
  if (!Number.isInteger(sampleSize) || sampleSize < 0) throw new Error('Query preview sample size must be a non-negative integer');
  return {
    queryId: query.id,
    queryName: query.name,
    columns: [...result.columns],
    rowCount: result.rowCount,
    sampleRows: result.rows.slice(0, sampleSize).map((row) => [...row]),
  };
}

export function summarizeQueryResult(snapshot: QueryResultSnapshot): string {
  return `Loaded ${snapshot.rowCount} rows × ${snapshot.columns.length} columns into the sheet`;
}

export function createInlineJsonQuery(id: string, name: string, data: Record<string, unknown>[], steps: QueryDefinition['steps'] = []): QueryDefinition {
  return { id, name, connectorId: 'json', connectorConfig: { data }, steps };
}

export function validateQueryDefinition(query: QueryDefinition): void {
  if (!query || typeof query !== 'object') throw new Error('Query definition is required');
  if (typeof query.id !== 'string' || !query.id.trim()) throw new Error('Query id is required');
  if (typeof query.name !== 'string' || !query.name.trim()) throw new Error('Query name is required');
  if (query.name.length > 200) throw new Error(`Query ${query.id} name is too long`);
  if (typeof query.connectorId !== 'string' || !query.connectorId.trim()) throw new Error('Query connectorId is required');
  if (!query.connectorConfig || typeof query.connectorConfig !== 'object' || Array.isArray(query.connectorConfig)) throw new Error(`Query ${query.id} has invalid connector configuration`);
  if (!Array.isArray(query.steps)) throw new Error(`Query ${query.id} steps must be an array`);
  validateQuerySteps(query.steps);
  if (query.sourceRevision !== undefined && (!Number.isInteger(query.sourceRevision) || query.sourceRevision < 0)) throw new Error(`Query ${query.id} sourceRevision must be a non-negative integer`);
  if (query.refreshPolicy) {
    const intervalMs = query.refreshPolicy.intervalMs;
    if (query.refreshPolicy.mode === 'interval' && !(typeof intervalMs === 'number' && Number.isInteger(intervalMs) && intervalMs > 0)) throw new Error(`Query ${query.id} interval refresh requires a positive intervalMs`);
    if (query.refreshPolicy.mode !== 'manual' && query.refreshPolicy.mode !== 'on-open' && query.refreshPolicy.mode !== 'interval') throw new Error(`Query ${query.id} has an unknown refresh policy`);
  }
}

function validateQueryResult(result: QueryResult): void {
  if (!result || !Array.isArray(result.columns) || !Array.isArray(result.rows) || result.rowCount !== result.rows.length) throw new Error('Query result is invalid');
  if (result.columns.length < 1) throw new Error('Query result must contain at least one column');
  if (result.columns.some((column) => typeof column !== 'string' || !column.trim())) throw new Error('Query result columns must be non-empty strings');
  const names = new Set<string>();
  for (const column of result.columns) {
    if (names.has(column)) throw new Error(`Query result contains duplicate column "${column}"`);
    names.add(column);
  }
  for (const row of result.rows) {
    if (!Array.isArray(row) || row.length !== result.columns.length) throw new Error('Query result row width does not match columns');
    for (const value of row) if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') throw new Error('Query result contains a non-scalar value');
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
  if (pivot.source.kind !== 'worksheet-range') throw new Error(`Pivot ${pivot.id} requires an explicit query target for this source kind`);
  return structuredClone(pivot.source.range);
}

function inferQueryFields(sourceId: string, columns: readonly string[], rows: readonly (readonly TableScalar[])[]): DataSourceField[] {
  const states = columns.map(() => ({ number: true, boolean: true, string: true, date: true, present: false }));
  for (const row of rows) {
    for (let ordinal = 0; ordinal < columns.length; ordinal += 1) {
      const value = row[ordinal] ?? null;
      const state = states[ordinal]!;
      if (value === null) continue;
      state.number &&= typeof value === 'number';
      state.boolean &&= typeof value === 'boolean';
      state.string &&= typeof value === 'string';
      if (value !== '') {
        state.present = true;
        state.date &&= typeof value === 'string' && /^\d{4}-\d{2}-\d{2}(?:[T ].*)?$/.test(value) && pivotTimelineInstant(value) !== undefined;
      }
    }
  }
  return columns.map((name, ordinal) => {
    const state = states[ordinal]!;
    const type: DataSourceFieldType = state.number ? 'number' : state.boolean ? 'boolean'
      : state.present && state.date ? 'date' : state.string ? 'text' : 'mixed';
    return { id: `${sourceId}:field:${ordinal}`, name, ordinal, type };
  });
}

function normalizeQueryRowsForDataSource(
  rows: readonly (readonly TableScalar[])[],
  fields: readonly DataSourceField[],
  rowOffset = 0,
): readonly (readonly TableScalar[])[] {
  if (!fields.some((field) => field.type === 'date')) return rows;
  const excelEpoch = Date.UTC(1899, 11, 30);
  return rows.map((row, rowIndex) => row.map((value, ordinal) => {
    if (fields[ordinal]?.type !== 'date') return value;
    if (value === null || value === '') return null;
    if (typeof value === 'number') return value;
    if (typeof value !== 'string') throw new Error(`Query date field ${String(ordinal)} row ${String(rowOffset + rowIndex)} cannot be encoded`);
    const instant = pivotTimelineInstant(value);
    if (instant === undefined) throw new Error(`Query date field ${String(ordinal)} row ${String(rowOffset + rowIndex)} is invalid`);
    return (instant - excelEpoch) / PIVOT_DAY_MS;
  }));
}

export function querySourceId(queryId: string): string {
  const id = queryId.trim();
  if (!id) throw new Error('Query id is required for block-backed loading');
  if (!/^[A-Za-z0-9._:-]{1,180}$/.test(id)) throw new Error(`Query id cannot be used as a data source identity: ${id}`);
  return `query:${id}`;
}

function columnarField(sourceId: string, name: string, ordinal: number, type: DataSourceFieldType): ColumnarBlockField {
  return { id: `${sourceId}:field:${ordinal}`, name, ordinal, type };
}

function targetRangeForQuery(workbook: WorkbookModel, target: LoadTarget, columns: readonly string[], rowCount: number): RangeRef {
  const checkedRange = (sheetId: string, startRow: number, startColumn: number): RangeRef => {
    workbook.getSheet(sheetId);
    const endRow = startRow + rowCount;
    const endColumn = startColumn + columns.length - 1;
    if (!Number.isSafeInteger(startRow) || !Number.isSafeInteger(startColumn) || startRow < 0 || startColumn < 0
      || !Number.isSafeInteger(endRow) || !Number.isSafeInteger(endColumn)
      || endRow >= MAX_SHEET_ROW_COUNT || endColumn >= MAX_SHEET_COLUMN_COUNT) {
      throw new Error('Query result exceeds worksheet bounds');
    }
    return { sheetId, startRow, endRow, startColumn, endColumn };
  };
  if (target.kind === 'range') {
    if (!target.sheetId || !target.range) throw new Error('Range query target requires sheetId and range');
    if (![target.range.startRow, target.range.startColumn].every((value) => Number.isSafeInteger(value) && value >= 0)) throw new Error('Range query target start is invalid');
    return checkedRange(target.sheetId, target.range.startRow, target.range.startColumn);
  }
  if (target.kind === 'sheet-table') {
    if (!target.sheetId || !target.tableId) throw new Error('Sheet-table query target requires sheetId and tableId');
    const table = workbook.getSheet(target.sheetId).sheetTables.find((entry) => entry.id === target.tableId);
    if (!table) throw new Error(`Unknown sheet table: ${target.tableId}`);
    if (!table.hasHeaderRow) throw new Error(`Sheet table ${table.name} has no header row for a canonical data region`);
    if (columns.length !== table.columns.length) throw new Error(`Query result columns must exactly match table ${table.name}`);
    const capacity = table.range.endRow - table.range.startRow - (table.hasTotalRow ? 1 : 0);
    if (rowCount > capacity) throw new Error(`Query result has too many rows for table ${table.name}`);
    return checkedRange(target.sheetId, table.range.startRow, table.range.startColumn);
  }
  if (target.kind === 'pivot-source') {
    if (!target.pivotId) throw new Error('Pivot-source query target requires pivotId');
    const found = findPivot(workbook, target.pivotId);
    if (!found) throw new Error(`Unknown pivot: ${target.pivotId}`);
    const fallback = target.range === undefined ? sourceRangeForPivot(found.pivot) : undefined;
    const sheetId = target.sheetId ?? fallback?.sheetId;
    if (!sheetId) throw new Error(`Pivot-source query target ${target.pivotId} requires sheetId`);
    const startRow = target.range?.startRow ?? fallback!.startRow;
    const startColumn = target.range?.startColumn ?? fallback!.startColumn;
    const capacityRows = target.range?.endRow === undefined ? fallback ? fallback.endRow - fallback.startRow : undefined : target.range.endRow - startRow;
    const capacityColumns = target.range?.endColumn === undefined ? fallback ? fallback.endColumn - fallback.startColumn + 1 : undefined : target.range.endColumn - startColumn + 1;
    if (capacityRows !== undefined && (capacityRows < 0 || rowCount > capacityRows)) throw new Error(`Query result does not fit pivot ${target.pivotId} source range`);
    if (capacityColumns !== undefined && (capacityColumns < 1 || columns.length > capacityColumns)) throw new Error(`Query result does not fit pivot ${target.pivotId} source range`);
    return checkedRange(sheetId, startRow, startColumn);
  }
  throw new Error(`Query target ${target.kind} does not project to a worksheet region`);
}

export interface QueryBlockLoadMetadata {
  columns: readonly string[];
  columnTypes: readonly DataSourceFieldType[];
  rowCount: number;
  blockRowCount: number;
}

export interface PreparedQueryLoadBlock {
  ref: DataBlockRef;
  payload: ArrayBuffer;
}

export function validateQueryBlockLoadMetadata(metadata: QueryBlockLoadMetadata): void {
  if (!metadata || !Array.isArray(metadata.columns) || metadata.columns.length === 0) throw new Error('Query block metadata must contain columns');
  if (metadata.columns.some((column) => typeof column !== 'string' || !column.trim())) throw new Error('Query block columns must be non-empty strings');
  if (metadata.columns.length > MAX_SHEET_COLUMN_COUNT || metadata.columns.some((column) => column.length > 200)) throw new Error('Query block columns exceed data-source limits');
  if (new Set(metadata.columns).size !== metadata.columns.length) throw new Error('Query block columns must be unique');
  if (!Array.isArray(metadata.columnTypes) || metadata.columnTypes.length !== metadata.columns.length) throw new Error('Query block column types do not match columns');
  const types = new Set<DataSourceFieldType>(['text', 'number', 'boolean', 'date', 'mixed']);
  if (metadata.columnTypes.some((type) => !types.has(type))) throw new Error('Query block contains an unsupported column type');
  if (!Number.isSafeInteger(metadata.rowCount) || metadata.rowCount < 0) throw new Error('Query block rowCount must be a non-negative integer');
  if (!Number.isSafeInteger(metadata.blockRowCount) || metadata.blockRowCount < 1 || metadata.blockRowCount > DEFAULT_DATA_BLOCK_ROW_COUNT) throw new Error('Query block row size is invalid');
}

function queryFields(sourceId: string, metadata: QueryBlockLoadMetadata): DataSourceField[] {
  return metadata.columns.map((name, ordinal) => ({
    id: `${sourceId}:field:${ordinal}`,
    name,
    ordinal,
    type: metadata.columnTypes[ordinal]!,
  }));
}

export async function encodeQueryLoadBlock(
  sourceId: string,
  revision: number,
  metadata: QueryBlockLoadMetadata,
  startRow: number,
  rows: readonly (readonly TableScalar[])[],
): Promise<PreparedQueryLoadBlock> {
  validateQueryBlockLoadMetadata(metadata);
  if (!Number.isSafeInteger(revision) || revision < 0) throw new Error('Query block revision must be a non-negative integer');
  if (!Number.isSafeInteger(startRow) || startRow < 0 || startRow % metadata.blockRowCount !== 0) throw new Error('Query block start row is invalid');
  if (rows.length < 1 || rows.length > metadata.blockRowCount || startRow + rows.length > metadata.rowCount) throw new Error('Query block row coverage is invalid');
  if (rows.some((row) => row.length !== metadata.columns.length)) throw new Error('Query block row width does not match columns');
  const fields = queryFields(sourceId, metadata);
  const normalizedRows = normalizeQueryRowsForDataSource(rows, fields);
  const payload = await encodeColumnarBlock({ fields: fields.map((field) => columnarField(sourceId, field.name, field.ordinal, field.type)), rows: normalizedRows });
  const id = `query-block:${crypto.randomUUID()}`;
  return {
    ref: {
      id,
      dataSourceId: sourceId,
      startRow,
      rowCount: rows.length,
      storageKey: `data-source/${sourceId}/revision-${revision}/${id}`,
      checksum: await computeColumnarBlockChecksum(payload),
      byteLength: payload.byteLength,
      encoding: COLUMNAR_BLOCK_ENCODING,
      revision,
    },
    payload,
  };
}

export function buildQueryLoadPayloadFromBlocks(
  workbook: WorkbookModel,
  query: QueryDefinition,
  target: LoadTarget,
  metadata: QueryBlockLoadMetadata,
  blocks: readonly DataBlockRef[],
): QueryLoadCommandPayload {
  validateQueryDefinition(query);
  validateQueryBlockLoadMetadata(metadata);
  const sourceId = querySourceId(query.id);
  const revision = (workbook.dataModel.sources.get(sourceId)?.revision ?? -1) + 1;
  const ordered = [...blocks].sort((left, right) => left.startRow - right.startRow);
  const blockIds = new Set<string>();
  let covered = 0;
  for (const block of ordered) {
    if (!blockIds.add(block.id) || block.dataSourceId !== sourceId || block.revision !== revision || block.startRow !== covered
      || block.rowCount < 1 || block.rowCount > metadata.blockRowCount || block.encoding !== COLUMNAR_BLOCK_ENCODING
      || !/^[A-Fa-f0-9]{64}$/.test(block.checksum) || !Number.isSafeInteger(block.byteLength) || block.byteLength < 1
      || !block.storageKey.trim() || block.startRow + block.rowCount > metadata.rowCount) {
      throw new Error(`Query block coverage is invalid: ${block.id}`);
    }
    covered += block.rowCount;
  }
  if (covered !== metadata.rowCount) throw new Error('Query blocks do not cover the complete result');
  const fields = queryFields(sourceId, metadata);
  const sourceRange = target.kind === 'workbook-table' ? undefined : targetRangeForQuery(workbook, target, metadata.columns, metadata.rowCount);
  const source: DataSourceManifest = {
    schema: 'DataSourceManifest',
    version: 1,
    id: sourceId,
    name: query.name,
    kind: 'chunked-table',
    ...(sourceRange ? { sourceSheetId: sourceRange.sheetId, sourceRange: structuredClone(sourceRange) } : {}),
    rowCount: metadata.rowCount,
    fields,
    blockRowCount: DEFAULT_DATA_BLOCK_ROW_COUNT,
    blocks: ordered.map((block) => structuredClone(block)),
    revision,
  };
  const definition = serializeQueryDefinition({ ...query, sourceRevision: revision });
  if (target.kind === 'workbook-table') {
    if (!target.tableId) throw new Error('Workbook-table query target requires tableId');
    const table = workbook.dataModel.tables.get(target.tableId);
    if (!table) throw new Error(`Unknown workbook table: ${target.tableId}`);
    const nextTable: WorkbookTableModel = {
      ...structuredClone(table),
      sourceId,
      sourceSheetId: undefined,
      sourceRange: undefined,
      rowCount: metadata.rowCount,
      fields,
      blocks: [],
      revision: table.revision + 1,
    };
    return { kind: 'data-source-load', queryId: query.id, queryDefinition: definition, target: structuredClone(target), sourceId, source, binding: { kind: 'workbook-table', tableId: table.id, table: nextTable } };
  }
  if (!sourceRange) throw new Error(`Query target ${target.kind} requires a worksheet range`);
  const sheet = workbook.getSheet(sourceRange.sheetId);
  return {
    kind: 'data-source-load',
    queryId: query.id,
    queryDefinition: definition,
    target: structuredClone(target),
    sourceId,
    source,
    binding: { kind: 'sheet-region', region: { id: `${sourceId}:region`, sourceId, range: structuredClone(sourceRange), headerRow: sourceRange.startRow, revision }, header: [...metadata.columns] },
    extent: { sheetId: sourceRange.sheetId, rowCount: Math.max(sheet.rowCount, sourceRange.endRow + 1), columnCount: Math.max(sheet.columnCount, sourceRange.endColumn + 1) },
    ...(target.kind === 'pivot-source' && target.pivotId ? { pivotSource: { kind: 'data-source', dataSourceId: sourceId } as PivotSource } : {}),
  };
}

export async function prepareQueryLoadPayload(workbook: WorkbookModel, query: QueryDefinition, target: LoadTarget, result: QueryResult): Promise<PreparedQueryLoad> {
  validateQueryDefinition(query);
  validateQueryResult(result);
  const sourceId = querySourceId(query.id);
  const previousSource = workbook.dataModel.sources.get(sourceId);
  const revision = (previousSource?.revision ?? -1) + 1;
  const fields = inferQueryFields(sourceId, result.columns, result.rows);
  const blocks: Array<{ ref: DataBlockRef; payload: ArrayBuffer }> = [];
  for (let startRow = 0; startRow < result.rows.length; startRow += DEFAULT_DATA_BLOCK_ROW_COUNT) {
    const sourceRows = result.rows.slice(startRow, startRow + DEFAULT_DATA_BLOCK_ROW_COUNT);
    const rows = normalizeQueryRowsForDataSource(sourceRows, fields, startRow);
    const blockPayload = await encodeColumnarBlock({ fields: fields.map((field) => columnarField(sourceId, field.name, field.ordinal, field.type)), rows });
    const blockId = `query-block:${crypto.randomUUID()}`;
    blocks.push({ ref: { id: blockId, dataSourceId: sourceId, startRow, rowCount: rows.length, storageKey: `data-source/${sourceId}/revision-${revision}/${blockId}`, checksum: await computeColumnarBlockChecksum(blockPayload), byteLength: blockPayload.byteLength, encoding: COLUMNAR_BLOCK_ENCODING, revision }, payload: blockPayload });
  }
  const sourceRange = target.kind === 'workbook-table' ? undefined : targetRangeForQuery(workbook, target, result.columns, result.rowCount);
  const source: DataSourceManifest = { schema: 'DataSourceManifest', version: 1, id: sourceId, name: query.name, kind: 'chunked-table', ...(sourceRange ? { sourceSheetId: sourceRange.sheetId, sourceRange: structuredClone(sourceRange) } : {}), rowCount: result.rows.length, fields, blockRowCount: DEFAULT_DATA_BLOCK_ROW_COUNT, blocks: blocks.map((entry) => entry.ref), revision };
  const definition = serializeQueryDefinition({ ...query, sourceRevision: revision });
  if (target.kind === 'workbook-table') {
    if (!target.tableId) throw new Error('Workbook-table query target requires tableId');
    const table = workbook.dataModel.tables.get(target.tableId);
    if (!table) throw new Error(`Unknown workbook table: ${target.tableId}`);
    const nextTable: WorkbookTableModel = { ...structuredClone(table), sourceId, sourceSheetId: undefined, sourceRange: undefined, rowCount: result.rowCount, fields: fields.map((field) => ({ id: field.id, name: field.name, ordinal: field.ordinal, type: field.type })), blocks: [], revision: table.revision + 1 };
    return { payload: { kind: 'data-source-load', queryId: query.id, queryDefinition: definition, target: structuredClone(target), sourceId, source, binding: { kind: 'workbook-table', tableId: table.id, table: nextTable } }, blocks };
  }
  if (!sourceRange) throw new Error(`Query target ${target.kind} requires a worksheet range`);
  const sheet = workbook.getSheet(sourceRange.sheetId);
  return {
    payload: {
      kind: 'data-source-load', queryId: query.id, queryDefinition: definition, target: structuredClone(target), sourceId, source,
      binding: { kind: 'sheet-region', region: { id: `${sourceId}:region`, sourceId, range: structuredClone(sourceRange), headerRow: sourceRange.startRow, revision }, header: [...result.columns] },
      extent: { sheetId: sourceRange.sheetId, rowCount: Math.max(sheet.rowCount, sourceRange.endRow + 1), columnCount: Math.max(sheet.columnCount, sourceRange.endColumn + 1) },
      ...(target.kind === 'pivot-source' && target.pivotId ? { pivotSource: { kind: 'data-source', dataSourceId: sourceId } as PivotSource } : {}),
    },
    blocks,
  };
}

function currentBinding(workbook: WorkbookModel, sourceId: string, target?: LoadTarget): QueryLoadBinding | null {
  for (const sheet of workbook.getSheets()) {
    const region = sheet.dataRegions.find((entry) => entry.sourceId === sourceId);
    if (region) {
      const header: TableScalar[] = [];
      for (let column = region.range.startColumn; column <= region.range.endColumn; column += 1) header.push(sheet.cells.get(region.headerRow, column)?.value ?? null);
      return { kind: 'sheet-region', region: structuredClone(region), header };
    }
  }
  if (target?.kind === 'workbook-table' && target.tableId) {
    const table = workbook.dataModel.tables.get(target.tableId);
    if (table) return { kind: 'workbook-table', tableId: table.id, table: structuredClone(table) };
  }
  for (const table of workbook.dataModel.tables.values()) if (table.sourceId === sourceId) return { kind: 'workbook-table', tableId: table.id, table: structuredClone(table) };
  return null;
}

function currentRanges(workbook: WorkbookModel, sourceId: string): RangeRef[] {
  return workbook.getSheets().flatMap((sheet) => sheet.dataRegions.filter((region) => region.sourceId === sourceId).map((region) => structuredClone(region.range)));
}

export interface QueryLoadPlan {
  mutationId: 'query.load.range' | 'query.load.sheet-table' | 'query.load.workbook-table' | 'query.load.pivot-source';
  payload: QueryLoadCommandPayload;
  inverse: QueryLoadRestorePayload;
  affectedRanges: RangeRef[];
}

export function buildQueryLoadPlan(workbook: WorkbookModel, params: QueryLoadCommandPayload): QueryLoadPlan {
  if (params.kind !== 'data-source-load') throw new Error('Query load payload must use the block-backed data-source contract');
  if (!params.queryId.trim() || params.sourceId !== querySourceId(params.queryId)) throw new Error('Query load source identity is invalid');
  if (!params.source || !params.binding) throw new Error('Query load source and binding are required');
  const previousBinding = currentBinding(workbook, params.sourceId, params.target);
  const previousSource = workbook.dataModel.sources.get(params.sourceId);
  const affectedRanges = [...currentRanges(workbook, params.sourceId), ...(params.binding.kind === 'sheet-region' ? [params.binding.region.range] : [])];
  const previousSheetId = previousBinding?.kind === 'sheet-region'
    ? previousBinding.region.range.sheetId
    : (params.target.sheetId ?? (params.binding.kind === 'sheet-region' ? params.binding.region.range.sheetId : undefined));
  const previousSheet = previousSheetId ? workbook.getSheet(previousSheetId) : undefined;
  const inverse: QueryLoadRestorePayload = {
    kind: 'data-source-load', queryId: params.queryId, queryDefinition: workbook.getQueryDefinition(params.queryId) ?? null,
    target: structuredClone(params.target), sourceId: params.sourceId, source: previousSource ? structuredClone(previousSource) : null,
    binding: previousBinding,
    ...(params.target.kind === 'pivot-source' && params.target.pivotId
      ? { pivotSource: workbook.getSheets().flatMap((sheet) => sheet.pivots).find((pivot) => pivot.id === params.target.pivotId)?.source }
      : {}),
    ...(previousSheet ? { extent: { sheetId: previousSheet.id, rowCount: previousSheet.rowCount, columnCount: previousSheet.columnCount } } : {}),
  };
  const mutationId = params.target.kind === 'range' ? 'query.load.range' : params.target.kind === 'sheet-table' ? 'query.load.sheet-table' : params.target.kind === 'pivot-source' ? 'query.load.pivot-source' : 'query.load.workbook-table';
  return { mutationId, payload: structuredClone(params), inverse, affectedRanges: affectedRanges.filter((range, index, all) => all.findIndex((candidate) => JSON.stringify(candidate) === JSON.stringify(range)) === index) };
}
