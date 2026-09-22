import type { RangeRef, SheetId } from './index';
import { MAX_SHEET_COLUMN_COUNT, MAX_SHEET_ROW_COUNT } from './sheet-extent';

/**
 * Canonical metadata for a block-backed workbook data source. The bytes stay
 * outside WorkbookSnapshot so collaboration and history remain JSON-only.
 */
export const DATA_SOURCE_SCHEMA = 'DataSourceManifest' as const;
export const DATA_SOURCE_VERSION = 1 as const;
export const DEFAULT_DATA_BLOCK_ROW_COUNT = 65_536;
export const LARGE_DATA_CELL_THRESHOLD = 100_000;
const MAX_DATA_SOURCE_NAME_LENGTH = 200;
const MAX_DATA_BLOCK_STORAGE_KEY_LENGTH = 500;

export type DataSourceFieldType = 'text' | 'number' | 'boolean' | 'date' | 'mixed';
export type DataBlockEncoding = 'columnar-v1';
export type DataSourceKind = 'worksheet-range' | 'sheet-table' | 'chunked-table';

export interface DataSourceField {
  id: string;
  name: string;
  ordinal: number;
  type: DataSourceFieldType;
}

/** A content-addressed block descriptor. It intentionally excludes bytes. */
export interface DataBlockRef {
  id: string;
  dataSourceId: string;
  startRow: number;
  rowCount: number;
  storageKey: string;
  checksum: string;
  byteLength: number;
  encoding: DataBlockEncoding;
  revision: number;
}

export interface DataSourceSortCriterion {
  fieldId: string;
  ascending: boolean;
}

export interface DataSourceSortState {
  criteria: DataSourceSortCriterion[];
}

export interface DataSourceManifest {
  schema: typeof DATA_SOURCE_SCHEMA;
  version: typeof DATA_SOURCE_VERSION;
  id: string;
  name: string;
  kind: DataSourceKind;
  sourceSheetId?: SheetId;
  sourceRange?: RangeRef;
  rowCount: number;
  fields: DataSourceField[];
  blockRowCount: number;
  blocks: DataBlockRef[];
  /** Logical row -> immutable physical row mapping used by virtual sorts. */
  rowOrder?: number[];
  sortState?: DataSourceSortState;
  revision: number;
}

/**
 * A rectangular sheet projection backed by a DataSourceManifest. CellMatrix
 * remains the authoritative sparse overlay for cells outside this region.
 */
export interface SheetDataRegion {
  id: string;
  sourceId: string;
  range: RangeRef;
  headerRow: number;
  revision: number;
}

export type DataBlockAvailability = 'ready' | 'loading' | 'missing' | 'error';

export interface DataBlockLoadState {
  sourceId: string;
  blockId: string;
  availability: DataBlockAvailability;
  error?: string;
}

export function isLargeDataSourceCellCount(cellCount: number): boolean {
  return Number.isSafeInteger(cellCount) && cellCount >= LARGE_DATA_CELL_THRESHOLD;
}

export function normalizeDataSourceManifest(input: DataSourceManifest): DataSourceManifest {
  if (input.schema !== DATA_SOURCE_SCHEMA || input.version !== DATA_SOURCE_VERSION) {
    throw new Error('Unsupported data source manifest');
  }
  if (typeof input.id !== 'string' || !/^[A-Za-z0-9._:-]{1,200}$/.test(input.id)) throw new Error('Data source id is invalid');
  if (typeof input.name !== 'string' || !input.name.trim() || input.name.length > MAX_DATA_SOURCE_NAME_LENGTH) throw new Error('Data source name is invalid');
  if (!['worksheet-range', 'sheet-table', 'chunked-table'].includes(input.kind)) throw new Error('Data source kind is invalid');
  if (!Number.isSafeInteger(input.rowCount) || input.rowCount < 0) throw new Error('Data source rowCount must be non-negative');
  if (input.blockRowCount !== DEFAULT_DATA_BLOCK_ROW_COUNT) throw new Error(`Data source blockRowCount must be ${String(DEFAULT_DATA_BLOCK_ROW_COUNT)}`);
  if (!Number.isSafeInteger(input.revision) || input.revision < 0) throw new Error('Data source revision must be non-negative');
  const rowOrder = input.rowOrder === undefined ? undefined : [...input.rowOrder];
  if (rowOrder !== undefined) {
    if (rowOrder.length !== input.rowCount) throw new Error('Data source rowOrder must cover every source row');
    const physicalRows = new Set<number>();
    for (const physicalRow of rowOrder) {
      if (!Number.isSafeInteger(physicalRow) || physicalRow < 0 || physicalRow >= input.rowCount || physicalRows.has(physicalRow)) {
        throw new Error('Data source rowOrder must be a permutation of source rows');
      }
      physicalRows.add(physicalRow);
    }
  }

  if (!Array.isArray(input.fields) || input.fields.length < 1 || input.fields.length > MAX_SHEET_COLUMN_COUNT) {
    throw new Error('Data source fields must contain between 1 and the worksheet column limit');
  }
  const fieldIds = new Set<string>();
  const fields = input.fields.map((field, ordinal) => {
    if (typeof field.id !== 'string' || !/^[A-Za-z0-9._:-]{1,200}$/.test(field.id) || fieldIds.has(field.id)) throw new Error(`Invalid or duplicate data source field: ${String(field.id)}`);
    if (typeof field.name !== 'string' || !field.name.trim() || field.name.length > 200) throw new Error(`Invalid data source field name: ${field.id}`);
    if (field.ordinal !== ordinal) throw new Error('Data source fields must use contiguous ordinals');
    if (!['text', 'number', 'boolean', 'date', 'mixed'].includes(field.type)) throw new Error(`Invalid data source field type: ${field.id}`);
    fieldIds.add(field.id);
    return { ...field };
  });
  const sortState = input.sortState === undefined
    ? undefined
    : { criteria: input.sortState.criteria.map((criterion) => ({ ...criterion })) };
  if (sortState !== undefined) {
    if (sortState.criteria.length === 0) throw new Error('Data source sortState requires at least one criterion');
    const sortedFieldIds = new Set<string>();
    for (const criterion of sortState.criteria) {
      if (!fieldIds.has(criterion.fieldId) || sortedFieldIds.has(criterion.fieldId) || typeof criterion.ascending !== 'boolean') {
        throw new Error('Data source sortState criteria are invalid');
      }
      sortedFieldIds.add(criterion.fieldId);
    }
  }
  const sourceSheetId = input.sourceSheetId;
  const sourceRange = input.sourceRange;
  if ((sourceSheetId === undefined) !== (sourceRange === undefined)) throw new Error('Data source worksheet identity and range must be declared together');
  if ((input.kind === 'worksheet-range' || input.kind === 'sheet-table') && sourceRange === undefined) {
    throw new Error(`${input.kind} data source requires a worksheet range`);
  }
  if (sourceRange !== undefined) {
    if (typeof sourceSheetId !== 'string' || !sourceSheetId.trim() || sourceRange.sheetId !== sourceSheetId) throw new Error('Data source range must target sourceSheetId');
    const coordinates = [sourceRange.startRow, sourceRange.endRow, sourceRange.startColumn, sourceRange.endColumn];
    if (!coordinates.every(Number.isSafeInteger)
      || sourceRange.startRow < 0 || sourceRange.endRow < sourceRange.startRow || sourceRange.endRow >= MAX_SHEET_ROW_COUNT
      || sourceRange.startColumn < 0 || sourceRange.endColumn < sourceRange.startColumn || sourceRange.endColumn >= MAX_SHEET_COLUMN_COUNT) {
      throw new Error('Data source range is invalid');
    }
    if (sourceRange.endRow - sourceRange.startRow !== input.rowCount
      || sourceRange.endColumn - sourceRange.startColumn + 1 !== fields.length) {
      throw new Error('Data source range dimensions do not match its rows and fields');
    }
  }

  if (!Array.isArray(input.blocks) || input.blocks.length > 100_000) throw new Error('Data source has too many blocks');
  const blockIds = new Set<string>();
  const storageKeys = new Set<string>();
  const blocks = input.blocks.map((block) => {
    if (block.dataSourceId !== input.id) throw new Error(`Data block ${block.id} belongs to another data source`);
    if (typeof block.id !== 'string' || !/^[A-Za-z0-9._:-]{1,200}$/.test(block.id) || blockIds.has(block.id)) throw new Error(`Invalid or duplicate data block: ${String(block.id)}`);
    if (!Number.isSafeInteger(block.startRow) || block.startRow < 0 || !Number.isSafeInteger(block.rowCount) || block.rowCount <= 0
      || block.startRow > input.rowCount || block.rowCount > input.rowCount - block.startRow) {
      throw new Error(`Invalid data block range: ${block.id}`);
    }
    if (block.rowCount > input.blockRowCount) throw new Error(`Data block exceeds blockRowCount: ${block.id}`);
    if (typeof block.storageKey !== 'string' || !block.storageKey.trim() || block.storageKey.length > MAX_DATA_BLOCK_STORAGE_KEY_LENGTH || storageKeys.has(block.storageKey)
      || !/^[A-Fa-f0-9]{64}$/.test(block.checksum) || !Number.isSafeInteger(block.byteLength) || block.byteLength < 1 || block.encoding !== 'columnar-v1') {
      throw new Error(`Invalid data block storage descriptor: ${block.id}`);
    }
    if (!Number.isSafeInteger(block.revision) || block.revision < 0) throw new Error(`Invalid data block revision: ${block.id}`);
    if (block.revision !== input.revision) throw new Error(`Data block revision does not match source revision: ${block.id}`);
    blockIds.add(block.id);
    storageKeys.add(block.storageKey);
    return { ...block };
  }).sort((left, right) => left.startRow - right.startRow);

  let coveredUntil = 0;
  for (const block of blocks) {
    if (block.startRow < coveredUntil) throw new Error('Data blocks must not overlap');
    if (block.startRow !== coveredUntil) throw new Error('Data blocks must provide contiguous source coverage');
    coveredUntil = block.startRow + block.rowCount;
  }
  if (coveredUntil !== input.rowCount) throw new Error('Data blocks must cover the complete source rowCount');

  return {
    ...input,
    fields,
    blocks,
    ...(rowOrder === undefined ? {} : { rowOrder }),
    ...(sortState === undefined ? {} : { sortState }),
    sourceRange: sourceRange ? { ...sourceRange } : undefined,
  };
}
