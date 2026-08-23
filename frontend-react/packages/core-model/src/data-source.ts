import type { RangeRef, SheetId } from './index';

/**
 * Canonical metadata for a block-backed workbook data source. The bytes stay
 * outside WorkbookSnapshot so collaboration and history remain JSON-only.
 */
export const DATA_SOURCE_SCHEMA = 'DataSourceManifest' as const;
export const DATA_SOURCE_VERSION = 1 as const;
export const DEFAULT_DATA_BLOCK_ROW_COUNT = 65_536;
export const LARGE_DATA_CELL_THRESHOLD = 100_000;

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
  if (!input.id.trim() || !input.name.trim()) throw new Error('Data source id and name are required');
  if (!Number.isSafeInteger(input.rowCount) || input.rowCount < 0) throw new Error('Data source rowCount must be non-negative');
  if (!Number.isSafeInteger(input.blockRowCount) || input.blockRowCount <= 0) throw new Error('Data source blockRowCount must be positive');
  if (!Number.isSafeInteger(input.revision) || input.revision < 0) throw new Error('Data source revision must be non-negative');

  const fieldIds = new Set<string>();
  const fields = input.fields.map((field, ordinal) => {
    if (!field.id.trim() || fieldIds.has(field.id)) throw new Error(`Duplicate data source field: ${field.id}`);
    if (field.ordinal !== ordinal) throw new Error('Data source fields must use contiguous ordinals');
    fieldIds.add(field.id);
    return { ...field };
  });
  const blockIds = new Set<string>();
  const blocks = input.blocks.map((block) => {
    if (block.dataSourceId !== input.id) throw new Error(`Data block ${block.id} belongs to another data source`);
    if (!block.id.trim() || blockIds.has(block.id)) throw new Error(`Duplicate data block: ${block.id}`);
    if (!Number.isSafeInteger(block.startRow) || block.startRow < 0 || !Number.isSafeInteger(block.rowCount) || block.rowCount <= 0) {
      throw new Error(`Invalid data block range: ${block.id}`);
    }
    if (!block.storageKey.trim() || !block.checksum.trim() || !Number.isSafeInteger(block.byteLength) || block.byteLength < 0) {
      throw new Error(`Invalid data block storage descriptor: ${block.id}`);
    }
    if (!Number.isSafeInteger(block.revision) || block.revision < 0) throw new Error(`Invalid data block revision: ${block.id}`);
    blockIds.add(block.id);
    return { ...block };
  }).sort((left, right) => left.startRow - right.startRow);

  for (let index = 1; index < blocks.length; index += 1) {
    const previous = blocks[index - 1]!;
    const current = blocks[index]!;
    if (previous.startRow + previous.rowCount > current.startRow) throw new Error('Data blocks must not overlap');
  }

  return {
    ...input,
    fields,
    blocks,
    sourceRange: input.sourceRange ? { ...input.sourceRange } : undefined,
  };
}
