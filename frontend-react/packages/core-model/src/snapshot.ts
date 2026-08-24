import type {
  DefinedNameModel,
  DataSourceManifest,
  SheetId,
  SheetSnapshot,
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
  version: 3;
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

export const WORKBOOK_SNAPSHOT_SCHEMA_REVISION = 3 as const;

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
  return structuredClone(snapshot);
}

export function loadWorkbookFromSnapshot(snapshot: WorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(assertCanonicalWorkbookSnapshot(snapshot));
}

export function createWorkbookSnapshot(workbook: WorkbookModel): WorkbookSnapshot {
  return workbook.snapshot();
}
