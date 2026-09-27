import type {
  DataSourceManifest,
  RangeRef,
  SheetDataRegion,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import type { DataSourceContentQuery } from './content-query';
import { readCellPatch, type CellPatch } from './resolved-cell';

export interface CanonicalDataSourceRegion {
  manifest: DataSourceManifest;
  sheet: WorksheetModel;
  region: SheetDataRegion;
}

export interface DataSourceCellPatchIdentity {
  row: number;
  column: number;
  value?: CellPatch['value'];
  formula?: CellPatch['formula'];
  formulaValue?: CellPatch['formulaValue'];
}

interface CachedPatchIdentity {
  rangeKey: string;
  cellRevision: number;
  entries: readonly DataSourceCellPatchIdentity[];
}

const patchIdentityCache = new WeakMap<WorksheetModel, Map<string, CachedPatchIdentity>>();

function sameRange(left: RangeRef | undefined, right: RangeRef | undefined): boolean {
  return left !== undefined && right !== undefined
    && left.sheetId === right.sheetId
    && left.startRow === right.startRow
    && left.endRow === right.endRow
    && left.startColumn === right.startColumn
    && left.endColumn === right.endColumn;
}

/** Fixed-order semantic identity shared by Pivot revision and reader validation. */
export function canonicalDataSourceManifestIdentity(manifest: DataSourceManifest): unknown {
  return {
    schema: manifest.schema,
    version: manifest.version,
    id: manifest.id,
    name: manifest.name,
    kind: manifest.kind,
    sourceSheetId: manifest.sourceSheetId,
    sourceRange: manifest.sourceRange === undefined ? undefined : {
      sheetId: manifest.sourceRange.sheetId,
      startRow: manifest.sourceRange.startRow,
      endRow: manifest.sourceRange.endRow,
      startColumn: manifest.sourceRange.startColumn,
      endColumn: manifest.sourceRange.endColumn,
    },
    rowCount: manifest.rowCount,
    fields: manifest.fields.map((field) => ({
      id: field.id,
      name: field.name,
      ordinal: field.ordinal,
      type: field.type,
    })),
    blockRowCount: manifest.blockRowCount,
    blocks: manifest.blocks.map((block) => ({
      id: block.id,
      dataSourceId: block.dataSourceId,
      startRow: block.startRow,
      rowCount: block.rowCount,
      storageKey: block.storageKey,
      checksum: block.checksum,
      byteLength: block.byteLength,
      encoding: block.encoding,
      revision: block.revision,
    })),
    rowOrder: manifest.rowOrder === undefined ? undefined : [...manifest.rowOrder],
    sortState: manifest.sortState === undefined ? undefined : {
      criteria: manifest.sortState.criteria.map((criterion) => ({ ...criterion })),
    },
    revision: manifest.revision,
  };
}

export function sameCanonicalDataSourceManifest(left: DataSourceManifest, right: DataSourceManifest): boolean {
  if (left.schema !== right.schema || left.version !== right.version || left.id !== right.id
    || left.name !== right.name || left.kind !== right.kind || left.sourceSheetId !== right.sourceSheetId
    || left.rowCount !== right.rowCount || left.blockRowCount !== right.blockRowCount || left.revision !== right.revision) {
    return false;
  }
  if (left.sourceRange === undefined ? right.sourceRange !== undefined
    : right.sourceRange === undefined || !sameRange(left.sourceRange, right.sourceRange)) return false;
  if (left.fields.length !== right.fields.length || left.blocks.length !== right.blocks.length) return false;
  for (let index = 0; index < left.fields.length; index += 1) {
    const a = left.fields[index]!;
    const b = right.fields[index]!;
    if (a.id !== b.id || a.name !== b.name || a.ordinal !== b.ordinal || a.type !== b.type) return false;
  }
  for (let index = 0; index < left.blocks.length; index += 1) {
    const a = left.blocks[index]!;
    const b = right.blocks[index]!;
    if (a.id !== b.id || a.dataSourceId !== b.dataSourceId || a.startRow !== b.startRow
      || a.rowCount !== b.rowCount || a.storageKey !== b.storageKey || a.checksum !== b.checksum
      || a.byteLength !== b.byteLength || a.encoding !== b.encoding || a.revision !== b.revision) return false;
  }
  if (left.rowOrder === undefined ? right.rowOrder !== undefined
    : right.rowOrder === undefined || left.rowOrder.length !== right.rowOrder.length) return false;
  if (left.rowOrder && right.rowOrder) {
    for (let index = 0; index < left.rowOrder.length; index += 1) {
      if (left.rowOrder[index] !== right.rowOrder[index]) return false;
    }
  }
  if (left.sortState === undefined ? right.sortState !== undefined
    : right.sortState === undefined || left.sortState.criteria.length !== right.sortState.criteria.length) return false;
  if (left.sortState && right.sortState) {
    for (let index = 0; index < left.sortState.criteria.length; index += 1) {
      const a = left.sortState.criteria[index]!;
      const b = right.sortState.criteria[index]!;
      if (a.fieldId !== b.fieldId || a.ascending !== b.ascending) return false;
    }
  }
  return true;
}

/**
 * Resolve the one worksheet projection owned by a manifest sourceRange.
 * Other projections may reference the same immutable source, but Pivot
 * provenance always belongs to this canonical range and is never selected by
 * first-match order.
 */
export function resolveCanonicalDataSourceRegion(
  workbook: WorkbookModel,
  sourceId: string,
  query?: DataSourceContentQuery,
): CanonicalDataSourceRegion {
  const manifest = workbook.getDataSource(sourceId);
  const range = manifest.sourceRange;
  if (!range || manifest.sourceSheetId !== range.sheetId) {
    throw new Error(`Data source ${sourceId} has no canonical worksheet range`);
  }
  const sheet = workbook.getSheet(range.sheetId);
  const matches = sheet.dataRegions.filter((candidate) => candidate.sourceId === sourceId && sameRange(candidate.range, range));
  if (matches.length !== 1) {
    throw new Error(`Data source ${sourceId} must have exactly one canonical sheet region; found ${String(matches.length)}`);
  }
  const region = matches[0]!;
  const width = range.endColumn - range.startColumn + 1;
  const bodyRows = range.endRow - range.startRow;
  if (range.startRow < 0 || range.endRow >= sheet.rowCount || range.startColumn < 0 || range.endColumn >= sheet.columnCount
    || region.headerRow !== range.startRow || region.revision !== manifest.revision
    || bodyRows !== manifest.rowCount || width !== manifest.fields.length) {
    throw new Error(`Data source ${sourceId} has inconsistent manifest and sheet-region metadata`);
  }
  if (query !== undefined && !sameCanonicalDataSourceManifest(manifest, query.manifest)) {
    throw new Error(`Data source ${sourceId} content reader does not match its canonical manifest`);
  }
  return { manifest, sheet, region };
}

/**
 * Deterministic, range-local CellPatch identity. The sparse scan runs once per
 * CellMatrix revision, so repeated Pivot revision reads remain O(1).
 */
export function dataSourceCellPatchIdentity(source: CanonicalDataSourceRegion): readonly DataSourceCellPatchIdentity[] {
  const { manifest, region, sheet } = source;
  const range = manifest.sourceRange!;
  const rangeKey = JSON.stringify([range.startRow, range.endRow, range.startColumn, range.endColumn]);
  let sheetCache = patchIdentityCache.get(sheet);
  if (!sheetCache) {
    sheetCache = new Map<string, CachedPatchIdentity>();
    patchIdentityCache.set(sheet, sheetCache);
  }
  const activeSourceIds = new Set(sheet.dataRegions.map((entry) => entry.sourceId));
  for (const cachedSourceId of sheetCache.keys()) {
    if (!activeSourceIds.has(cachedSourceId)) sheetCache.delete(cachedSourceId);
  }
  const cached = sheetCache.get(manifest.id);
  if (cached?.rangeKey === rangeKey && cached.cellRevision === sheet.cells.revision) return cached.entries;

  const entries: DataSourceCellPatchIdentity[] = [];
  sheet.cells.forEachInRange(range.startRow + 1, range.endRow, range.startColumn, range.endColumn, (cell, row, column) => {
    const patch = readCellPatch(cell);
    if (!patch) throw new Error(`Data region ${region.id} contains a non-canonical cell overlay`);
    if (patch.value === undefined && patch.formula === undefined && patch.formulaValue === undefined) return;
    entries.push({
      row: row - region.headerRow - 1,
      column: column - range.startColumn,
      ...(patch.value === undefined ? {} : { value: patch.value }),
      ...(patch.formula === undefined ? {} : { formula: patch.formula }),
      ...(patch.formulaValue === undefined ? {} : { formulaValue: patch.formulaValue }),
    });
  });
  entries.sort((left, right) => left.row - right.row || left.column - right.column);
  sheetCache.set(manifest.id, { rangeKey, cellRevision: sheet.cells.revision, entries });
  return entries;
}
