import type {
  DefinedNameModel,
  DataSourceManifest,
  SheetId,
  SheetSnapshot,
  RangeRef,
  WorkbookModel,
  WorkbookTableModel,
  UnitId,
} from './index';
import type { PrintDocumentSnapshot, QueryDefinitionSnapshot } from './workbook-state';
import { WorkbookModel as WorkbookModelClass } from './index';

/**
 * The single persisted/transport snapshot contract. Floating objects are
 * represented only by the canonical drawing collection and payload map.
 * There is deliberately no versioned sibling or legacy union in production.
 */
export interface WorkbookSnapshot {
  schema: 'WorkbookSnapshot';
  /** Canonical persisted schema revision. Non-matching snapshots are rejected. */
  version: 4;
  unitId: UnitId;
  name: string;
  dimensionMetrics: WorkbookDimensionMetrics;
  definedNames?: Record<string, string>;
  definedNameModels?: DefinedNameModel[];
  tables?: WorkbookTableModel[];
  /** Metadata only; large source bytes live in the local data-block store. */
  dataSources: DataSourceManifest[];
  printDocuments?: PrintDocumentSnapshot[];
  queryDefinitions?: QueryDefinitionSnapshot[];
  sheets: SheetSnapshot[];
}

export interface WorkbookDimensionMetrics {
  normalFontFamily: string;
  normalFontSizePx: number;
  maximumDigitWidthPx: number;
}

export const WORKBOOK_SNAPSHOT_SCHEMA_REVISION = 4 as const;

/**
 * One-way browser-storage migration. It preserves v2 native geometry exactly
 * as CSS pixels; it deliberately does not guess whether an old XLSX value was
 * originally a point or character-width measurement.
 */
export function migrateStoredWorkbookSnapshot(value: unknown): WorkbookSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Stored workbook snapshot must be an object');
  const input = structuredClone(value) as Record<string, any>;
  if (input.schema !== 'WorkbookSnapshot') throw new Error('Unsupported workbook snapshot schema');
  if (input.version === WORKBOOK_SNAPSHOT_SCHEMA_REVISION) return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
  if (input.version === 3 && Array.isArray(input.sheets)) {
    input.version = WORKBOOK_SNAPSHOT_SCHEMA_REVISION;
    for (const sheet of input.sheets as Array<Record<string, any>>) {
      if (sheet.pane && typeof sheet.pane === 'object' && sheet.pane.kind !== 'none') {
        sheet.pane.state = sheet.pane.kind === 'split' ? 'split' : (sheet.pane.state ?? 'frozen');
      }
      if (sheet.filter) {
        const legacy = sheet.filter as Record<string, any>;
        const columns: Record<string, any> = {};
        for (const [key, condition] of Object.entries(legacy.criteria ?? {})) {
          const item = condition as Record<string, any>;
          columns[key] = {
            column: Number(key),
            showButton: true,
            hiddenButton: false,
            criterion: item.selectedValues
              ? { kind: 'values', values: item.selectedValues, includeBlank: !item.excludeBlanks }
              : item.conditionOperator
                ? { kind: 'custom', join: 'and', conditions: [{ operator: item.conditionOperator, value: item.conditionValue ?? null }] }
                : undefined,
          };
        }
        sheet.autoFilter = { sheetId: legacy.sheetId, range: legacy.range, columns };
        delete sheet.filter;
      }
    }
    return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
  }
  if (input.version !== 2 || !Array.isArray(input.sheets)) throw new Error(`Unsupported workbook snapshot version: ${String(input.version)}`);
  input.version = WORKBOOK_SNAPSHOT_SCHEMA_REVISION;
  input.dimensionMetrics = { normalFontFamily: 'Calibri', normalFontSizePx: 14.6666666667, maximumDigitWidthPx: 7 };
  for (const sheet of input.sheets as Array<Record<string, any>>) {
    sheet.defaultRowHeightPx = finiteStoredSize(sheet.defaultRowHeight, 28);
    sheet.defaultColumnWidthPx = finiteStoredSize(sheet.defaultColumnWidth, 110);
    sheet.rowHeightsPx = storedSizeMap(sheet.rowHeights);
    sheet.columnWidthsPx = storedSizeMap(sheet.columnWidths);
    const freeze = sheet.freeze as Record<string, unknown> | undefined;
    const xSplit = Number(freeze?.xSplit ?? 0);
    const ySplit = Number(freeze?.ySplit ?? 0);
    sheet.pane = xSplit > 0 || ySplit > 0
      ? { kind: 'frozen', xSplit, ySplit, startRow: Number(freeze?.startRow ?? ySplit), startColumn: Number(freeze?.startColumn ?? xSplit), state: 'frozen' }
      : { kind: 'none' };
    migrateLegacyFontSizes(sheet.cells);
    migrateLegacyFontSizes(sheet.conditionalFormats);
    delete sheet.defaultRowHeight;
    delete sheet.defaultColumnWidth;
    delete sheet.rowHeights;
    delete sheet.columnWidths;
    delete sheet.freeze;
  }
  return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
}

function finiteStoredSize(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : fallback;
}

function storedSizeMap(value: unknown): Record<number, number> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  const result: Record<number, number> = {};
  for (const [key, entry] of Object.entries(value)) {
    const index = Number(key);
    if (Number.isSafeInteger(index) && index >= 0 && typeof entry === 'number' && Number.isFinite(entry) && entry > 0) result[index] = entry;
  }
  return result;
}

function migrateLegacyFontSizes(value: unknown): void {
  if (!value || typeof value !== 'object') return;
  if (Array.isArray(value)) {
    for (const entry of value) migrateLegacyFontSizes(entry);
    return;
  }
  const record = value as Record<string, unknown>;
  if (typeof record.fontSize === 'number' && record.fontSizePx === undefined) record.fontSizePx = record.fontSize;
  delete record.fontSize;
  for (const entry of Object.values(record)) migrateLegacyFontSizes(entry);
}

/**
 * The application reads one Snapshot contract only. A non-canonical record
 * must be repaired before it enters the workbook runtime.
 */
export function assertCanonicalWorkbookSnapshot(snapshot: WorkbookSnapshot): WorkbookSnapshot {
  if (snapshot.version !== WORKBOOK_SNAPSHOT_SCHEMA_REVISION) {
    throw new Error(`Unsupported workbook snapshot version: ${String(snapshot.version)}`);
  }
  if (!Array.isArray(snapshot.dataSources)) throw new Error('Workbook snapshot dataSources must be an array');
  if (!snapshot.dimensionMetrics || !snapshot.dimensionMetrics.normalFontFamily.trim()
    || !Number.isFinite(snapshot.dimensionMetrics.normalFontSizePx) || snapshot.dimensionMetrics.normalFontSizePx <= 0
    || !Number.isFinite(snapshot.dimensionMetrics.maximumDigitWidthPx) || snapshot.dimensionMetrics.maximumDigitWidthPx <= 0) throw new Error('Workbook snapshot dimensionMetrics is invalid');
  for (const sheet of snapshot.sheets) {
    const pane = sheet.pane;
    if (pane.kind === 'frozen') {
      if (!Number.isInteger(pane.xSplit) || !Number.isInteger(pane.ySplit) || pane.xSplit < 0 || pane.ySplit < 0) {
        throw new Error('Frozen pane split counts must be non-negative integers');
      }
    } else if (pane.kind === 'split') {
      if (!Number.isFinite(pane.xSplit) || !Number.isFinite(pane.ySplit) || pane.xSplit < 0 || pane.ySplit < 0) {
        throw new Error('Split pane positions must be finite non-negative numbers');
      }
    }
    if (sheet.autoFilter) {
      if (sheet.autoFilter.range.sheetId !== sheet.id) throw new Error('AutoFilter range must target its worksheet');
      for (const [key, column] of Object.entries(sheet.autoFilter.columns)) {
        if (!Number.isSafeInteger(Number(key)) || column.column !== Number(key)) throw new Error('AutoFilter column identity is invalid');
      }
    }
    const tableFilters = (sheet.sheetTables ?? []).filter((table) => Boolean(table.autoFilter));
    if (sheet.autoFilter && tableFilters.some((table) => rangesOverlap(sheet.autoFilter!.range, table.autoFilter!.range))) {
      throw new Error('Worksheet and Table AutoFilter ranges cannot overlap');
    }
    if (tableFilters.length > 1) throw new Error('A worksheet cannot have multiple Table AutoFilter owners');
  }
  return structuredClone(snapshot);
}

function rangesOverlap(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow && right.startRow <= left.endRow
    && left.startColumn <= right.endColumn && right.startColumn <= left.endColumn;
}

export function loadWorkbookFromSnapshot(snapshot: WorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(assertCanonicalWorkbookSnapshot(snapshot));
}

export function createWorkbookSnapshot(workbook: WorkbookModel): WorkbookSnapshot {
  return workbook.snapshot();
}
