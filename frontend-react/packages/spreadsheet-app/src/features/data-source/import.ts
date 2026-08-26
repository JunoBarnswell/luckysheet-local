import type {
  CellComment,
  CellData,
  CellStyle,
  DataBlockRef,
  DataSourceField,
  DataSourceFieldType,
  DataSourceManifest,
  RangeRef,
  SheetDataRegion,
  SheetId,
  SheetSnapshot,
  TableScalar,
} from '@react-sheets/core-model';
import {
  DEFAULT_DATA_BLOCK_ROW_COUNT,
  LARGE_DATA_CELL_THRESHOLD,
} from '@react-sheets/core-model';
import {
  COLUMNAR_BLOCK_ENCODING,
  computeColumnarBlockChecksum,
  encodeColumnarBlock,
  type ColumnarBlockField,
} from './codec';

/** A metadata-only cell entry kept beside the immutable columnar payload. */
export interface SparseCellOverlayMetadataCell {
  row: number;
  column: number;
  formula?: string;
  style?: CellStyle;
  comment?: CellComment;
}

/**
 * Metadata that must survive when a caller removes the ordinary cells that
 * backed a block. Coordinates are relative to the owning block, except for
 * the separate header overlay, whose coordinates are relative to the region.
 */
export interface SparseCellOverlayMetadata {
  schema: 'SparseCellOverlayMetadata';
  revision: number;
  cells: SparseCellOverlayMetadataCell[];
}

export interface SheetDataSourceImportInput {
  sheet: SheetSnapshot;
  /** Inclusive range. The first row is the field-header row. */
  range: RangeRef;
  /** Stable identity supplied by the workbook/session owner. */
  sourceId: string;
  sourceName?: string;
  regionId?: string;
  revision?: number;
}

export interface EncodedSheetDataBlock {
  ref: DataBlockRef;
  payload: ArrayBuffer;
  metadata: SparseCellOverlayMetadata;
}

export interface SheetDataSourceImport {
  manifest: DataSourceManifest;
  region: SheetDataRegion;
  /** Header values are returned separately so callers can restore the header. */
  header: TableScalar[];
  headerMetadata: SparseCellOverlayMetadata;
  blocks: EncodedSheetDataBlock[];
  nonEmptyCellCount: number;
}

/** The import gate is intentionally strict: equal to the limit stays sparse. */
export const DATA_SOURCE_IMPORT_CELL_THRESHOLD = LARGE_DATA_CELL_THRESHOLD;
export const DATA_SOURCE_IMPORT_BLOCK_ROWS = DEFAULT_DATA_BLOCK_ROW_COUNT;

function fail(message: string): never {
  throw new Error(`Sheet data source import ${message}`);
}

function isInteger(value: number): boolean {
  return Number.isSafeInteger(value);
}

function validateInput(input: SheetDataSourceImportInput): void {
  if (!input || typeof input !== 'object') fail('requires an input object');
  if (!input.sheet || typeof input.sheet !== 'object') fail('requires a sheet snapshot');
  if (!input.sourceId.trim()) fail('requires a source id');
  if (!input.range || input.range.sheetId !== input.sheet.id) {
    fail('range must target the supplied sheet');
  }
  const { startRow, endRow, startColumn, endColumn } = input.range;
  if (![startRow, endRow, startColumn, endColumn].every(isInteger)
    || startRow < 0 || endRow < startRow || startColumn < 0 || endColumn < startColumn) {
    fail('range must be a non-empty inclusive rectangle');
  }
  if (endRow >= input.sheet.rowCount || endColumn >= input.sheet.columnCount) {
    fail('range is outside the sheet bounds');
  }
  const revision = input.revision ?? 0;
  if (!isInteger(revision) || revision < 0) fail('revision must be non-negative');
  if (input.regionId !== undefined && !input.regionId.trim()) fail('region id cannot be empty');
  if (input.sourceName !== undefined && !input.sourceName.trim()) fail('source name cannot be empty');
}

function cellAt(sheet: SheetSnapshot, row: number, column: number): CellData | undefined {
  return sheet.cells[String(row)]?.[String(column)];
}

/** Values/formulas are content; formatting by itself is not a non-empty cell. */
function isNonEmptyCell(cell: CellData | undefined): boolean {
  return cell !== undefined && (cell.value !== null || cell.formula !== undefined);
}

export function countNonEmptyCells(sheet: SheetSnapshot, range: RangeRef): number {
  let count = 0;
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      if (isNonEmptyCell(cellAt(sheet, row, column))) count += 1;
    }
  }
  return count;
}

export function qualifiesForDataSourceImport(nonEmptyCellCount: number): boolean {
  return isInteger(nonEmptyCellCount) && nonEmptyCellCount > DATA_SOURCE_IMPORT_CELL_THRESHOLD;
}

function cloneStyle(style: CellStyle): CellStyle {
  return structuredClone(style);
}

function cloneComment(comment: CellComment): CellComment {
  return structuredClone(comment);
}

function reviewCommentAt(sheet: SheetSnapshot, row: number, column: number): CellComment | undefined {
  const threadId = sheet.review.threadIdsByCell[`${row}:${column}`]?.[0];
  const thread = threadId ? sheet.review.threadsById[threadId] : undefined;
  if (!thread) return undefined;
  return {
    id: thread.id,
    author: thread.author,
    text: thread.text,
    createdAt: thread.createdAt,
    mentions: thread.mentions,
    replies: thread.replies,
    resolved: thread.resolved,
    resolvedAt: thread.resolvedAt,
  };
}

function overlayCell(row: number, column: number, cell: CellData, comment?: CellComment): SparseCellOverlayMetadataCell | undefined {
  if (cell.formula === undefined && cell.style === undefined && comment === undefined) return undefined;
  return {
    row,
    column,
    ...(cell.formula === undefined ? {} : { formula: cell.formula }),
    ...(cell.style === undefined ? {} : { style: cloneStyle(cell.style) }),
    ...(comment === undefined ? {} : { comment: cloneComment(comment) }),
  };
}

function emptyOverlay(revision: number): SparseCellOverlayMetadata {
  return { schema: 'SparseCellOverlayMetadata', revision, cells: [] };
}

function headerName(value: TableScalar, ordinal: number): string {
  if (value === null) return `Column ${ordinal + 1}`;
  const text = String(value).trim();
  return text || `Column ${ordinal + 1}`;
}

function fieldId(sourceId: string, ordinal: number): string {
  return `${sourceId}:field:${ordinal}`;
}

function dateFormatOf(cell: CellData | undefined): string | undefined {
  return cell?.numberFormat ?? cell?.style?.numberFormat;
}

function hasDateTokens(format: string | undefined): boolean {
  if (!format) return false;
  const withoutLiterals = format
    .replace(/"(?:[^"]|"")*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '');
  return /[dy]/i.test(withoutLiterals);
}

function inferFieldType(cells: readonly (CellData | undefined)[]): DataSourceFieldType {
  const present = cells.filter((cell): cell is CellData => cell !== undefined && cell.value !== null);
  if (present.length === 0) return 'mixed';
  const types = new Set(present.map((cell) => typeof cell.value));
  if (types.size !== 1) return 'mixed';
  const firstType = present[0]!.value;
  if (typeof firstType === 'string') return 'text';
  if (typeof firstType === 'boolean') return 'boolean';
  if (present.every((cell) => hasDateTokens(dateFormatOf(cell)))) return 'date';
  return 'number';
}

function scalarOf(cell: CellData | undefined): TableScalar {
  return cell?.value ?? null;
}

function dataField(
  sourceId: string,
  ordinal: number,
  name: string,
  type: DataSourceFieldType,
): DataSourceField {
  return { id: fieldId(sourceId, ordinal), name, ordinal, type };
}

function columnarField(field: DataSourceField): ColumnarBlockField {
  return { id: field.id, name: field.name, ordinal: field.ordinal, type: field.type };
}

function dataBlockId(sourceId: string, startRow: number): string {
  return `${sourceId}:block:${startRow}`;
}

function dataBlockStorageKey(sourceId: string, blockId: string): string {
  return `data-source/${sourceId}/${blockId}`;
}

function dataRows(input: SheetDataSourceImportInput): number {
  return input.range.endRow - input.range.startRow;
}

function buildHeader(
  input: SheetDataSourceImportInput,
  columnCount: number,
  revision: number,
): { values: TableScalar[]; metadata: SparseCellOverlayMetadata } {
  const values: TableScalar[] = [];
  const metadata = emptyOverlay(revision);
  for (let column = 0; column < columnCount; column += 1) {
    const absoluteColumn = input.range.startColumn + column;
    const cell = cellAt(input.sheet, input.range.startRow, absoluteColumn);
    values.push(scalarOf(cell));
    const entry = cell ? overlayCell(0, column, cell, reviewCommentAt(input.sheet, input.range.startRow, absoluteColumn)) : undefined;
    if (entry) metadata.cells.push(entry);
  }
  return { values, metadata };
}

function buildManifest(
  input: SheetDataSourceImportInput,
  fields: DataSourceField[],
  blocks: DataBlockRef[],
  rowCount: number,
  revision: number,
): DataSourceManifest {
  return {
    schema: 'DataSourceManifest',
    version: 1,
    id: input.sourceId,
    name: input.sourceName ?? `${input.sheet.name} data`,
    kind: 'worksheet-range',
    sourceSheetId: input.sheet.id as SheetId,
    sourceRange: { ...input.range },
    rowCount,
    fields,
    blockRowCount: DATA_SOURCE_IMPORT_BLOCK_ROWS,
    blocks,
    revision,
  };
}

/**
 * Convert one headered, contiguous sheet rectangle into block-backed data.
 *
 * The function is deterministic for the same source id, range and snapshot.
 * It returns `undefined` at or below the large-data gate so the caller can
 * retain the ordinary sparse CellMatrix path. It never mutates the snapshot.
 */
export async function encodeSheetDataRegion(
  input: SheetDataSourceImportInput,
): Promise<SheetDataSourceImport | undefined> {
  validateInput(input);
  const nonEmptyCellCount = countNonEmptyCells(input.sheet, input.range);
  if (!qualifiesForDataSourceImport(nonEmptyCellCount)) return undefined;

  const revision = input.revision ?? 0;
  const columnCount = input.range.endColumn - input.range.startColumn + 1;
  const rowCount = dataRows(input);
  if (rowCount <= 0) fail('requires at least one data row after the header');

  const header = buildHeader(input, columnCount, revision);
  const fields: DataSourceField[] = [];
  for (let ordinal = 0; ordinal < columnCount; ordinal += 1) {
    const absoluteColumn = input.range.startColumn + ordinal;
    const cells: CellData[] = [];
    for (let row = input.range.startRow + 1; row <= input.range.endRow; row += 1) {
      const cell = cellAt(input.sheet, row, absoluteColumn);
      if (cell) cells.push(cell);
    }
    fields.push(dataField(input.sourceId, ordinal, headerName(header.values[ordinal]!, ordinal), inferFieldType(cells)));
  }

  const blockPayloads: EncodedSheetDataBlock[] = [];
  for (let blockStart = 0; blockStart < rowCount; blockStart += DATA_SOURCE_IMPORT_BLOCK_ROWS) {
    const blockRowCount = Math.min(DATA_SOURCE_IMPORT_BLOCK_ROWS, rowCount - blockStart);
    const rows: TableScalar[][] = [];
    const metadata = emptyOverlay(revision);
    for (let relativeRow = 0; relativeRow < blockRowCount; relativeRow += 1) {
      const absoluteRow = input.range.startRow + 1 + blockStart + relativeRow;
      const row: TableScalar[] = [];
      for (let relativeColumn = 0; relativeColumn < columnCount; relativeColumn += 1) {
        const absoluteColumn = input.range.startColumn + relativeColumn;
        const cell = cellAt(input.sheet, absoluteRow, absoluteColumn);
        row.push(scalarOf(cell));
        const entry = cell ? overlayCell(relativeRow, relativeColumn, cell, reviewCommentAt(input.sheet, absoluteRow, absoluteColumn)) : undefined;
        if (entry) metadata.cells.push(entry);
      }
      rows.push(row);
    }

    const payload = await encodeColumnarBlock({
      fields: fields.map(columnarField),
      rows,
    });
    const blockId = dataBlockId(input.sourceId, blockStart);
    const ref: DataBlockRef = {
      id: blockId,
      dataSourceId: input.sourceId,
      startRow: blockStart,
      rowCount: blockRowCount,
      storageKey: dataBlockStorageKey(input.sourceId, blockId),
      checksum: await computeColumnarBlockChecksum(payload),
      byteLength: payload.byteLength,
      encoding: COLUMNAR_BLOCK_ENCODING,
      revision,
    };
    blockPayloads.push({ ref, payload, metadata });
  }

  const region: SheetDataRegion = {
    id: input.regionId ?? `${input.sourceId}:region`,
    sourceId: input.sourceId,
    range: { ...input.range },
    headerRow: input.range.startRow,
    revision,
  };
  return {
    manifest: buildManifest(input, fields, blockPayloads.map((entry) => entry.ref), rowCount, revision),
    region,
    header: header.values,
    headerMetadata: header.metadata,
    blocks: blockPayloads,
    nonEmptyCellCount,
  };
}
