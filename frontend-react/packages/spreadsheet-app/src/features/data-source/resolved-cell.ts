import type {
  CellData,
  CellStyle,
  CellValue,
  DataSourceManifest,
  RangeRef,
  SheetDataRegion,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import type { DataSourceContentLoadState, DataSourceContentQuery } from './content-query';

/**
 * A field-level patch deliberately distinguishes "do not touch" from
 * "remove this field".  Block bytes are immutable, so edits to a block are
 * represented by this patch surface and never by a second materialized cell
 * value that could shadow the block value.
 */
export type CellPatchField<T> =
  | { kind: 'inherit' }
  | { kind: 'set'; value: T }
  | { kind: 'clear' };

type CellDataField = Exclude<keyof CellData, 'value'>;
type CellPatchKey = 'value' | CellDataField;

export interface CellPatch {
  schema: 'CellPatch';
  revision?: number;
  value?: CellPatchField<CellValue>;
  formula?: CellPatchField<NonNullable<CellData['formula']>>;
  displayValue?: CellPatchField<NonNullable<CellData['displayValue']>>;
  styleId?: CellPatchField<NonNullable<CellData['styleId']>>;
  style?: CellPatchField<CellStyle>;
  editor?: CellPatchField<NonNullable<CellData['editor']>>;
  presentation?: CellPatchField<NonNullable<CellData['presentation']>>;
  phonetic?: CellPatchField<NonNullable<CellData['phonetic']>>;
  numberFormat?: CellPatchField<NonNullable<CellData['numberFormat']>>;
  richText?: CellPatchField<NonNullable<CellData['richText']>>;
  formulaMetadata?: CellPatchField<NonNullable<CellData['formulaMetadata']>>;
  formulaValue?: CellPatchField<NonNullable<CellData['formulaValue']>>;
  filterMetadata?: CellPatchField<NonNullable<CellData['filterMetadata']>>;
}

/** Canonical runtime/persistence carrier beside a sparse CellMatrix entry. */
export type CellPatchCarrier = CellData & { __cellPatch: CellPatch };

export type ResolvedCellSource = 'cell-matrix' | 'data-block' | 'data-block-overlay';

export interface ResolvedCell {
  sheetId: string;
  row: number;
  column: number;
  source: ResolvedCellSource;
  /** The immutable block value before a sparse patch is applied. */
  base?: CellData;
  /** The canonical field-level patch, if the cell has one. */
  patch?: CellPatch;
  /** The cell visible to formulas/renderers. */
  cell?: CellData;
  /** Block load state. Ordinary sparse cells do not have a load state. */
  state?: DataSourceContentLoadState;
  region?: SheetDataRegion;
}

export type DataContentMap = ReadonlyMap<string, DataSourceContentQuery>;

/** The single application read contract for sparse and block-backed cells. */
export interface WorkbookCellResolver {
  resolve(sheet: WorksheetModel, row: number, column: number): ResolvedCell | undefined;
}

export function createWorkbookCellResolver(dataContent: DataContentMap = new Map()): WorkbookCellResolver {
  return { resolve: (sheet, row, column) => resolveCell(sheet, row, column, dataContent) };
}

export interface MaterializedDataRegionCell {
  row: number;
  column: number;
  cell: CellData;
}

/** Prepared transaction data; preparing never mutates WorkbookModel. */
export interface PreparedDataRegionMaterialization {
  sheetId: string;
  region: SheetDataRegion;
  regionIndex: number;
  manifest: DataSourceManifest;
  /** Whether commit is expected to remove the last manifest reference. */
  willRemoveSource: boolean;
  range: RangeRef;
  previousCells: MaterializedDataRegionCell[];
  materializedCells: MaterializedDataRegionCell[];
  materializedCellCount: number;
}

const PATCH_FIELDS: readonly CellDataField[] = [
  'formula',
  'displayValue',
  'styleId',
  'style',
  'editor',
  'numberFormat',
  'formulaValue',
  'filterMetadata',
  'phonetic',
];
const PATCH_KEYS: readonly CellPatchKey[] = ['value', ...PATCH_FIELDS];

function clone<T>(value: T): T {
  return structuredClone(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isPatchField(value: unknown): value is CellPatchField<unknown> {
  return isRecord(value) && (value.kind === 'inherit' || value.kind === 'set' || value.kind === 'clear')
    && (value.kind !== 'set' || 'value' in value);
}

function assertPatchField(value: unknown, key: string): asserts value is CellPatchField<unknown> {
  if (!isPatchField(value)) throw new Error(`Cell patch field ${key} is invalid`);
}

function assertRevision(value: unknown): void {
  if (value !== undefined && (!Number.isSafeInteger(value) || Number(value) < 0)) {
    throw new Error('Cell patch revision must be a non-negative safe integer');
  }
}

/** Validate and clone a patch before it enters the model or a command payload. */
export function normalizeCellPatch(input: CellPatch): CellPatch {
  if (!isRecord(input) || input.schema !== 'CellPatch') throw new Error('Cell patch schema is invalid');
  assertRevision(input.revision);
  const result: CellPatch = {
    schema: 'CellPatch',
    ...(input.revision === undefined ? {} : { revision: input.revision }),
  };
  for (const key of PATCH_KEYS) {
    const field = input[key];
    if (field === undefined) continue;
    assertPatchField(field, key);
    result[key] = clone(field) as never;
  }
  return result;
}

function hasPatchEffect(patch: CellPatch): boolean {
  return PATCH_KEYS.some((key) => patch[key] !== undefined && patch[key]!.kind !== 'inherit');
}

export function readCellPatch(cell: CellData | undefined): CellPatch | undefined {
  const candidate = (cell as CellPatchCarrier | undefined)?.__cellPatch;
  if (!candidate) return undefined;
  return normalizeCellPatch(candidate);
}

function fieldValue<T>(field: CellPatchField<T> | undefined, base: T | undefined): T | undefined {
  if (!field || field.kind === 'inherit') return base;
  return field.kind === 'set' ? clone(field.value) : undefined;
}

/** Apply a patch without mutating either the block base or the sparse cell. */
export function applyCellPatch(base: CellData | undefined, patch: CellPatch | undefined): CellData | undefined {
  if (!base && !patch) return undefined;
  const normalized = patch ? normalizeCellPatch(patch) : undefined;
  const result: Partial<CellData> = base ? clone(base) : { value: null };
  if (normalized?.value) result.value = fieldValue(normalized.value, result.value);
  if (result.value === undefined) result.value = null;
  for (const key of PATCH_FIELDS) {
    const field = normalized?.[key];
    if (!field || field.kind === 'inherit') continue;
    const value = fieldValue(field, result[key]);
    if (value === undefined) delete result[key];
    else result[key] = value as never;
  }
  return result as CellData;
}

function isBodyCell(region: SheetDataRegion, row: number, column: number): boolean {
  return row > region.headerRow
    && row <= region.range.endRow
    && column >= region.range.startColumn
    && column <= region.range.endColumn;
}

export function dataRegionAt(sheet: WorksheetModel, row: number, column: number): SheetDataRegion | undefined {
  return sheet.dataRegions.find((region) => isBodyCell(region, row, column));
}

function blockBase(
  row: number,
  column: number,
  region: SheetDataRegion,
  dataContent: DataContentMap,
): { base?: CellData; state: DataSourceContentLoadState } {
  const query = dataContent.get(region.sourceId);
  if (!query) {
    return {
      base: { value: '#BLOCK!', style: { textColor: '#b91c1c', bold: true } },
      state: { sourceId: region.sourceId, blockId: null, availability: 'missing', error: `Data source ${region.sourceId} is unavailable` },
    };
  }
  const result = query.peekCellValue(row - region.headerRow - 1, column - region.range.startColumn);
  if (result.value !== undefined) {
    return {
      base: { value: result.value },
      state: result.state,
    };
  }
  if (result.state.availability === 'loading') {
    return {
      base: { value: 'Loading…', style: { textColor: '#64748b', italic: true } },
      state: result.state,
    };
  }
  return {
    base: {
      value: '#BLOCK!',
      style: { textColor: '#b91c1c', bold: result.state.availability === 'error' },
    },
    state: result.state,
  };
}

/**
 * Resolve both ordinary sparse cells and block-backed cells through one read
 * surface.  `sheet.cells` is only an overlay inside a data-region body; the
 * block/query value remains the base value.
 */
export function resolveCell(
  sheet: WorksheetModel,
  row: number,
  column: number,
  dataContent: DataContentMap = new Map(),
): ResolvedCell | undefined {
  const raw = sheet.cells.get(row, column);
  const region = dataRegionAt(sheet, row, column);
  if (!region) {
    if (!raw) return undefined;
    return {
      sheetId: sheet.id,
      row,
      column,
      source: 'cell-matrix',
      base: raw,
      cell: raw,
    };
  }

  const loaded = blockBase(row, column, region, dataContent);
  const patch = readCellPatch(raw);
  if (raw && !patch) {
    throw new Error(`Data region ${region.id} contains a non-canonical cell overlay; migrate it before resolving`);
  }
  const cell = applyCellPatch(loaded.base, patch);
  return {
    sheetId: sheet.id,
    row,
    column,
    source: patch && hasPatchEffect(patch) ? 'data-block-overlay' : 'data-block',
    base: loaded.base,
    patch,
    cell,
    state: loaded.state,
    region: clone(region),
  };
}

function cellsInRange(sheet: WorksheetModel, range: RangeRef): MaterializedDataRegionCell[] {
  const cells: MaterializedDataRegionCell[] = [];
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const cell = sheet.cells.get(row, column);
      if (cell) cells.push({ row, column, cell: clone(cell) });
    }
  }
  return cells;
}

function sourceStillReferenced(workbook: WorkbookModel, sourceId: string): boolean {
  return workbook.getSheets().some((candidate) => candidate.dataRegions.some((region) => region.sourceId === sourceId));
}

/**
 * Prepare one complete data-region materialization without mutating the
 * workbook. The returned plan is the only input accepted by the commit path.
 */
export async function prepareDataRegionMaterialization(
  workbook: WorkbookModel,
  sheetId: string,
  regionId: string,
  dataContent: DataContentMap,
): Promise<PreparedDataRegionMaterialization> {
  const sheet = workbook.getSheet(sheetId);
  const regionIndex = sheet.dataRegions.findIndex((entry) => entry.id === regionId);
  if (regionIndex < 0) throw new Error(`Unknown data region: ${regionId}`);
  const region = clone(sheet.dataRegions[regionIndex]!);
  const manifest = workbook.getDataSource(region.sourceId);
  const query = dataContent.get(region.sourceId);
  if (!query) throw new Error(`Data source ${region.sourceId} is unavailable; cannot materialize data region ${region.id}`);
  if (query.manifest.id !== manifest.id || query.manifest.revision !== manifest.revision) {
    throw new Error(`Data source ${region.sourceId} manifest is stale; cannot materialize data region ${region.id}`);
  }
  const width = region.range.endColumn - region.range.startColumn + 1;
  const bodyRowCount = region.range.endRow - region.headerRow;
  if (region.headerRow !== region.range.startRow || bodyRowCount <= 0) {
    throw new Error(`Data region ${region.id} must contain a header and at least one data row`);
  }
  if (query.manifest.rowCount !== bodyRowCount || query.manifest.fields.length !== width) {
    throw new Error(`Data region ${region.id} does not match its data source shape`);
  }

  // Read every block before touching the workbook. Missing/error blocks leave
  // the region and its source reference intact for retry.
  const rows = await query.getRows(0, bodyRowCount);
  if (!rows.value || rows.state.availability !== 'ready') {
    throw new Error(rows.state.error ?? `Data region ${region.id} could not be fully loaded`);
  }

  // Hydration/import must have run the explicit migration before commands are
  // prepared. Preparation itself is read-only and never mutates the model.
  const previousCells = cellsInRange(sheet, region.range);
  const materialized: MaterializedDataRegionCell[] = [];
  for (let row = region.range.startRow; row <= region.range.endRow; row += 1) {
    for (let column = region.range.startColumn; column <= region.range.endColumn; column += 1) {
      if (row === region.headerRow) {
        const existing = sheet.cells.get(row, column);
        const field = manifest.fields[column - region.range.startColumn];
        materialized.push({
          row,
          column,
          cell: existing ? clone(existing) : { value: field?.name ?? null },
        });
        continue;
      }
      const resolved = resolveCell(sheet, row, column, dataContent);
      if (!resolved?.cell) throw new Error(`Data region ${region.id} cell ${row}:${column} could not be resolved`);
      materialized.push({ row, column, cell: clone(resolved.cell) });
    }
  }

  const range = clone(region.range);
  return {
    sheetId,
    region,
    regionIndex,
    manifest,
    willRemoveSource: !sourceStillReferenced(workbook, region.sourceId),
    range,
    previousCells,
    materializedCells: materialized,
    materializedCellCount: materialized.length,
  };
}

export function clearCellPatch<T>(value?: T): CellPatchField<T> {
  return value === undefined ? { kind: 'clear' } : { kind: 'set', value: clone(value) };
}

export function setCellPatch<T>(value: T): CellPatchField<T> {
  return { kind: 'set', value: clone(value) };
}

export function inheritCellPatch<T>(): CellPatchField<T> {
  return { kind: 'inherit' };
}
