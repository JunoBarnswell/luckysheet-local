import type { DefinedNameModel, SheetSnapshot, RangeRef, CellStyleTemplate, UnitId, WorkbookModel, WorkbookTheme, CellHyperlink } from './index';
import type { PrintDocumentSnapshot, QueryDefinitionSnapshot } from './workbook-state';
import { WorkbookModel as WorkbookModelClass } from './index';
import { MAX_DRAWING_SOURCE_CELLS } from './generated-workbook-limits';
import { canonicalizePivotDefinition, pivotSourceIdentity } from './pivot';
import { isAssetRef } from './asset';
import { canonicalSnapSettings, validateDrawingGraph } from './drawing-planner';
import { isCellEditorConfig } from './cell-editor';
import { DEFAULT_WORKBOOK_EDITING_OPTIONS, isWorkbookEditingOptions, type WorkbookEditingOptions } from './editing-options';
import { isChartSubtypeForType } from './domain';
import type { ChartDrawingPayload } from './domain';
import { isEmbeddedObjectDrawingPayload, isEquationDrawingPayload, isIconDrawingPayload, isModel3dDrawingPayload, isScreenshotDrawingPayload, isSignatureLineDrawingPayload, isSmartArtDrawingPayload, isWordArtDrawingPayload } from './domain';
import { isCellPhoneticMetadata } from './phonetic';
import type { ReviewStoreSnapshot } from './review-store';
import { normalizeDefinedNameModel } from './domain';
import { DEFAULT_WORKBOOK_CALCULATION_SETTINGS, isWorkbookCalculationSettings, type WorkbookCalculationSettings, type WorkbookCollationContext } from '@react-sheets/formula-engine';

/**
 * The single persisted/transport snapshot contract. Floating objects are
 * represented only by the canonical drawing collection and payload map.
 * There is deliberately no versioned sibling or legacy union in production.
 */
export interface WorkbookSnapshot {
  schema: 'WorkbookSnapshot';
  /** Canonical persisted schema revision. Non-matching snapshots are rejected. */
  version: 10;
  unitId: UnitId;
  name: string;
  dimensionMetrics: WorkbookDimensionMetrics;
  /** Workbook-owned deterministic ordering semantics for values and query keys. */
  collationContext?: WorkbookCollationContext;
  /** Workbook-owned calculation policy; formula workers consume this exact state. */
  calculationSettings: WorkbookCalculationSettings;
  editingOptions: WorkbookEditingOptions;
  /** Workbook-owned theme identity and resolved colors. */
  theme?: WorkbookTheme;
  definedNameModels: DefinedNameModel[];
  dataModel: import('./data-model').WorkbookDataModel;
  printDocuments?: PrintDocumentSnapshot[];
  queryDefinitions?: QueryDefinitionSnapshot[];
  /** Workbook-owned reusable styles and cell editor definitions. */
  cellStyleTemplates?: CellStyleTemplate[];
  sheets: SheetSnapshot[];
}

export { MAX_DRAWING_SOURCE_CELLS } from './generated-workbook-limits';

export interface WorkbookDimensionMetrics {
  normalFontFamily: string;
  normalFontSizePx: number;
  maximumDigitWidthPx: number;
}

export const WORKBOOK_SNAPSHOT_SCHEMA_REVISION = 10 as const;

/**
 * One-way browser-storage migration. It preserves v2 native geometry exactly
 * as CSS pixels; it deliberately does not guess whether an old XLSX value was
 * originally a point or character-width measurement.
 */
export function migrateStoredWorkbookSnapshot(value: unknown): WorkbookSnapshot {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Stored workbook snapshot must be an object');
  const input = structuredClone(value) as Record<string, any>;
  if (input.schema !== 'WorkbookSnapshot') throw new Error('Unsupported workbook snapshot schema');
  if (input.version === undefined && Array.isArray(input.sheets)) {
    input.version = input.dimensionMetrics && input.sheets.every((sheet: Record<string, unknown>) => sheet.pane && sheet.defaultRowHeightPx && sheet.defaultColumnWidthPx) ? 4 : 2;
  }
  if (input.version === WORKBOOK_SNAPSHOT_SCHEMA_REVISION) return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
  if (input.version === 9 && Array.isArray(input.sheets)) {
    migrateV9ToV10(input);
    input.version = WORKBOOK_SNAPSHOT_SCHEMA_REVISION;
    return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
  }
  if (input.version === 8 && Array.isArray(input.sheets)) {
    input.version = 9;
    input.editingOptions = structuredClone(DEFAULT_WORKBOOK_EDITING_OPTIONS);
    return migrateStoredWorkbookSnapshot(input);
  }
  if (input.version === 7 && Array.isArray(input.sheets)) {
    input.version = 9;
    input.editingOptions = structuredClone(DEFAULT_WORKBOOK_EDITING_OPTIONS);
    for (const sheet of input.sheets as Array<Record<string, any>>) migrateLegacyReview(sheet);
    return migrateStoredWorkbookSnapshot(input);
  }
  if (input.version === 6 && Array.isArray(input.sheets)) {
    input.version = 9;
    input.editingOptions = structuredClone(DEFAULT_WORKBOOK_EDITING_OPTIONS);
    input.calculationSettings = input.calculationSettings ?? structuredClone(DEFAULT_WORKBOOK_CALCULATION_SETTINGS);
    if (containsLegacyImageDataUrl(input)) throw new Error('ASSET_MIGRATION_REQUIRED: legacy image data must be assetized before runtime load');
    for (const sheet of input.sheets as Array<Record<string, any>>) migrateLegacyReview(sheet);
    return migrateStoredWorkbookSnapshot(input);
  }
  if (input.version === 5 && Array.isArray(input.sheets)) {
    input.version = 9;
    input.editingOptions = structuredClone(DEFAULT_WORKBOOK_EDITING_OPTIONS);
    input.calculationSettings = input.calculationSettings ?? structuredClone(DEFAULT_WORKBOOK_CALCULATION_SETTINGS);
    for (const sheet of input.sheets as Array<Record<string, any>>) migrateLegacyReview(sheet);
    return migrateStoredWorkbookSnapshot(input);
  }
  if (input.version === 4 && Array.isArray(input.sheets)) {
    input.version = 5;
    input.dataModel = {
      sources: Array.isArray(input.dataSources) ? input.dataSources : [],
      tables: Array.isArray(input.tables) ? input.tables : [],
      relationships: [],
      views: [],
    };
    delete input.dataSources;
    delete input.tables;
    for (const sheet of input.sheets as Array<Record<string, any>>) sheet.kind = sheet.kind ?? 'worksheet';
    return migrateStoredWorkbookSnapshot(input);
  }
  if (input.version === 3 && Array.isArray(input.sheets)) {
    input.version = 4;
    for (const sheet of input.sheets as Array<Record<string, any>>) {
      if (sheet.pane && typeof sheet.pane === 'object' && sheet.pane.kind !== 'none') {
        sheet.pane.state = sheet.pane.kind === 'split' ? 'split' : (sheet.pane.state ?? 'frozen');
      }
      migrateLegacyFilter(sheet);
    }
    return migrateStoredWorkbookSnapshot(input);
  }
  if (input.version !== 2 || !Array.isArray(input.sheets)) throw new Error(`Unsupported workbook snapshot version: ${String(input.version)}`);
  input.version = 4;
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
    migrateLegacyFilter(sheet);
    migrateLegacyFontSizes(sheet.cells);
    migrateLegacyFontSizes(sheet.conditionalFormats);
    delete sheet.defaultRowHeight;
    delete sheet.defaultColumnWidth;
    delete sheet.rowHeights;
    delete sheet.columnWidths;
    delete sheet.freeze;
  }
  return migrateStoredWorkbookSnapshot(input);
}

function migrateV9ToV10(snapshot: Record<string, any>): void {
  if (!Array.isArray(snapshot.sheets)) throw new Error('SNAPSHOT_MIGRATION_CONFLICT: v9 sheets must be an array');

  const models: DefinedNameModel[] = [];
  const identities = new Map<string, DefinedNameModel>();
  const addDefinedName = (value: unknown, source: string): void => {
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid defined name from ${source}`);
    }
    const raw = value as Record<string, unknown>;
    const allowed = new Set(['name', 'formula', 'scope', 'sheetId', 'anchor', 'hidden', 'comment']);
    if (Object.keys(raw).some((key) => !allowed.has(key))) {
      throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: unsupported defined name field from ${source}`);
    }
    let normalized: DefinedNameModel;
    try {
      normalized = normalizeDefinedNameModel(raw as unknown as DefinedNameModel);
    } catch (error) {
      throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid defined name from ${source}`, { cause: error });
    }
    const identity = `${normalized.scope}:${normalized.sheetId ?? ''}:${normalized.name.toUpperCase()}`;
    const previous = identities.get(identity);
    if (previous) {
      if (source.startsWith('definedNames.') && previous.formula === normalized.formula) return;
      if (!sameReviewValue(previous, normalized)) {
        throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: defined name identity ${identity} has conflicting values`);
      }
      return;
    }
    identities.set(identity, normalized);
    models.push(normalized);
  };

  if (snapshot.definedNameModels !== undefined) {
    if (!Array.isArray(snapshot.definedNameModels)) {
      throw new Error('SNAPSHOT_MIGRATION_CONFLICT: definedNameModels must be an array');
    }
    snapshot.definedNameModels.forEach((value: unknown, index: number) => addDefinedName(value, `definedNameModels[${index}]`));
  }
  if (snapshot.definedNames !== undefined) {
    if (!snapshot.definedNames || typeof snapshot.definedNames !== 'object' || Array.isArray(snapshot.definedNames)) {
      throw new Error('SNAPSHOT_MIGRATION_CONFLICT: definedNames must be an object');
    }
    for (const [name, formula] of Object.entries(snapshot.definedNames as Record<string, unknown>)) {
      if (typeof formula !== 'string') throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: definedNames.${name} formula must be a string`);
      addDefinedName({ name, formula, scope: 'workbook' }, `definedNames.${name}`);
    }
  }
  snapshot.definedNameModels = models;
  delete snapshot.definedNames;

  for (const rawSheet of snapshot.sheets as Array<Record<string, any>>) {
    if (!rawSheet || typeof rawSheet !== 'object' || Array.isArray(rawSheet)) {
      throw new Error('SNAPSHOT_MIGRATION_CONFLICT: worksheet must be an object');
    }
    const entries = new Map<string, { row: number; column: number; hyperlink: CellHyperlink }>();
    const addHyperlink = (row: unknown, column: unknown, value: unknown, source: string): void => {
      if (!Number.isSafeInteger(row) || Number(row) < 0 || Number(row) > 1_048_575
        || !Number.isSafeInteger(column) || Number(column) < 0 || Number(column) > 16_383) {
        throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid hyperlink coordinate from ${source}`);
      }
      const hyperlink = normalizeStoredHyperlink(value, source);
      const key = `${Number(row)}:${Number(column)}`;
      const previous = entries.get(key);
      if (previous) {
        if (!sameReviewValue(previous.hyperlink, hyperlink)) {
          throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: hyperlink at ${rawSheet.id}!${key} has conflicting values`);
        }
        return;
      }
      entries.set(key, { row: Number(row), column: Number(column), hyperlink });
    };

    if (rawSheet.hyperlinks !== undefined) {
      if (!Array.isArray(rawSheet.hyperlinks)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: hyperlinks on ${rawSheet.id} must be an array`);
      rawSheet.hyperlinks.forEach((entry: any, index: number) => {
        if (!entry || typeof entry !== 'object' || Array.isArray(entry)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid hyperlink entry on ${rawSheet.id}`);
        addHyperlink(entry.row, entry.column, entry.hyperlink, `sheet.hyperlinks[${index}]`);
      });
    }
    if (!rawSheet.cells || typeof rawSheet.cells !== 'object' || Array.isArray(rawSheet.cells)) {
      throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: cells on ${rawSheet.id} must be an object`);
    }
    for (const [rowKey, rowValue] of Object.entries(rawSheet.cells as Record<string, any>)) {
      if (!rowValue || typeof rowValue !== 'object' || Array.isArray(rowValue)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid row on ${rawSheet.id}`);
      for (const [columnKey, cell] of Object.entries(rowValue as Record<string, any>)) {
        if (!cell || typeof cell !== 'object' || Array.isArray(cell)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid cell on ${rawSheet.id}!${rowKey}:${columnKey}`);
        const hasLegacyString = Object.prototype.hasOwnProperty.call(cell, 'hyperlink');
        const hasLegacyDetail = Object.prototype.hasOwnProperty.call(cell, 'hyperlinkDetail');
        if (!hasLegacyString && !hasLegacyDetail) continue;
        const row = Number(rowKey);
        const column = Number(columnKey);
        let hyperlink: CellHyperlink;
        if (hasLegacyDetail) hyperlink = normalizeStoredHyperlink(cell.hyperlinkDetail, `cell ${rawSheet.id}!${rowKey}:${columnKey}.hyperlinkDetail`);
        else hyperlink = normalizeStoredHyperlink({
          id: `legacy-hyperlink-${rawSheet.id}-${row}-${column}`,
          target: { kind: 'url', url: cell.hyperlink },
        }, `cell ${rawSheet.id}!${rowKey}:${columnKey}.hyperlink`);
        if (hasLegacyString) {
          if (typeof cell.hyperlink !== 'string' || cell.hyperlink.trim() === ''
            || hyperlink.target.kind !== 'url' || hyperlink.target.url !== cell.hyperlink) {
            throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: hyperlink fields disagree on ${rawSheet.id}!${rowKey}:${columnKey}`);
          }
        }
        addHyperlink(row, column, hyperlink, `cell ${rawSheet.id}!${rowKey}:${columnKey}`);
        delete cell.hyperlink;
        delete cell.hyperlinkDetail;
      }
    }
    rawSheet.hyperlinks = [...entries.values()].map((entry) => ({ row: entry.row, column: entry.column, hyperlink: entry.hyperlink }));
    migrateDataRegionOverlays(rawSheet);
  }
}

const CELL_PATCH_FIELDS = [
  'formula', 'displayValue', 'styleId', 'style', 'editor', 'presentation', 'numberFormat',
  'richText', 'phonetic', 'formulaMetadata', 'formulaValue', 'filterMetadata',
] as const;
const CELL_DATA_FIELDS = new Set(['value', ...CELL_PATCH_FIELDS]);
const CELL_PATCH_KEYS = new Set(['schema', 'revision', 'value', ...CELL_PATCH_FIELDS]);

type SnapshotCellPatchField =
  | { kind: 'inherit' }
  | { kind: 'set'; value: unknown }
  | { kind: 'clear' };

type SnapshotCellPatch = {
  schema: 'CellPatch';
  revision?: number;
  value?: SnapshotCellPatchField;
  [key: string]: unknown;
};

function migrateDataRegionOverlays(sheet: Record<string, any>): void {
  if (sheet.dataRegions === undefined) return;
  if (!Array.isArray(sheet.dataRegions)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: dataRegions on ${sheet.id} must be an array`);
  if (!sheet.cells || typeof sheet.cells !== 'object' || Array.isArray(sheet.cells)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: cells on ${sheet.id} must be an object`);
  for (const region of sheet.dataRegions as Array<Record<string, any>>) {
    if (!region || typeof region !== 'object' || Array.isArray(region) || !region.range || typeof region.range !== 'object') {
      throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: data region on ${sheet.id} is invalid`);
    }
    const range = region.range as Record<string, any>;
    const startRow = Number(range.startRow);
    const headerRow = Number(region.headerRow);
    const endRow = Number(range.endRow);
    const startColumn = Number(range.startColumn);
    const endColumn = Number(range.endColumn);
    if (!Number.isSafeInteger(startRow) || !Number.isSafeInteger(headerRow) || !Number.isSafeInteger(endRow) || !Number.isSafeInteger(startColumn) || !Number.isSafeInteger(endColumn)
      || startRow < 0 || headerRow < startRow || headerRow > endRow || startColumn < 0 || startColumn > endColumn) {
      throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: data region bounds on ${sheet.id} are invalid`);
    }
    for (let row = headerRow + 1; row <= endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        const cell = sheet.cells[String(row)]?.[String(column)] as Record<string, any> | undefined;
        if (!cell) continue;
        const label = `${sheet.id}!${row}:${column}`;
        if (Object.prototype.hasOwnProperty.call(cell, '__cellPatch')) {
          validateCellPatchCarrier(cell, label);
          continue;
        }
        if (!Object.prototype.hasOwnProperty.call(cell, 'value')) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: data region cell ${label} has no value`);
        const unknown = Object.keys(cell).filter((key) => !CELL_DATA_FIELDS.has(key));
        if (unknown.length > 0) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: data region cell ${label} contains unsupported fields: ${unknown.join(', ')}`);
        const hasMetadata = CELL_PATCH_FIELDS.some((key) => cell[key] !== undefined);
        const patch: SnapshotCellPatch = {
          schema: 'CellPatch',
          value: hasMetadata ? { kind: 'inherit' } : { kind: 'set', value: structuredClone(cell.value) },
        };
        for (const key of CELL_PATCH_FIELDS) {
          if (cell[key] !== undefined) patch[key] = { kind: 'set', value: structuredClone(cell[key]) } satisfies SnapshotCellPatchField;
        }
        sheet.cells[String(row)][String(column)] = {
          value: hasMetadata ? null : structuredClone(cell.value),
          __cellPatch: patch,
        };
      }
    }
  }
}

export function assertCanonicalDataRegionOverlays(sheet: SheetSnapshot): void {
  const regions = sheet.dataRegions ?? [];
  if (!Array.isArray(regions)) throw new Error(`Worksheet ${sheet.id} dataRegions must be an array`);
  for (const [rowKey, row] of Object.entries(sheet.cells)) {
    const rowNumber = Number(rowKey);
    for (const [columnKey, cell] of Object.entries(row)) {
      const columnNumber = Number(columnKey);
      const inBody = regions.find((region) => rowNumber > region.headerRow
        && rowNumber <= region.range.endRow
        && columnNumber >= region.range.startColumn
        && columnNumber <= region.range.endColumn);
      const hasCarrier = Object.prototype.hasOwnProperty.call(cell, '__cellPatch');
      if (inBody && !hasCarrier) throw new Error(`Data region ${inBody.id} contains a non-canonical cell overlay at ${rowKey}:${columnKey}`);
      if (!inBody && hasCarrier) throw new Error(`Cell patch carrier ${sheet.id}!${rowKey}:${columnKey} is outside a data region`);
      if (hasCarrier) validateCellPatchCarrier(cell as unknown as Record<string, unknown>, `${sheet.id}!${rowKey}:${columnKey}`);
    }
  }
}

function validateCellPatchCarrier(cell: Record<string, unknown>, label: string): void {
  const keys = Object.keys(cell);
  if (keys.some((key) => key !== 'value' && key !== '__cellPatch')) throw new Error(`Data region cell ${label} mixes raw fields with its canonical CellPatch carrier`);
  if (!Object.prototype.hasOwnProperty.call(cell, 'value')) throw new Error(`Data region cell ${label} carrier value is missing`);
  const patch = cell.__cellPatch;
  if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new Error(`Data region cell ${label} CellPatch is invalid`);
  const value = patch as SnapshotCellPatch;
  if (value.schema !== 'CellPatch' || (value.revision !== undefined && (!Number.isSafeInteger(value.revision) || value.revision < 0))) {
    throw new Error(`Data region cell ${label} CellPatch identity is invalid`);
  }
  if (Object.keys(value).some((key) => !CELL_PATCH_KEYS.has(key))) throw new Error(`Data region cell ${label} CellPatch contains unsupported fields`);
  if (value.value !== undefined) validateCellPatchField(value.value, `${label}.value`);
  const expectedValue = value.value?.kind === 'set' ? value.value.value : null;
  if (!sameReviewValue(cell.value, expectedValue)) throw new Error(`Data region cell ${label} carrier value disagrees with its CellPatch`);
  for (const key of CELL_PATCH_FIELDS) {
    if (value[key] !== undefined) validateCellPatchField(value[key], `${label}.${key}`);
  }
}

function validateCellPatchField(value: unknown, label: string): asserts value is SnapshotCellPatchField {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`Data region CellPatch field ${label} is invalid`);
  const field = value as Record<string, unknown>;
  if (!['inherit', 'set', 'clear'].includes(String(field.kind)) || (field.kind === 'set' && !Object.prototype.hasOwnProperty.call(field, 'value'))) {
    throw new Error(`Data region CellPatch field ${label} is invalid`);
  }
}

function normalizeStoredHyperlink(value: unknown, source: string): CellHyperlink {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid hyperlink from ${source}`);
  const record = value as Record<string, any>;
  const allowed = new Set(['id', 'target', 'tooltip']);
  if (Object.keys(record).some((key) => !allowed.has(key)) || typeof record.id !== 'string' || record.id.trim() === ''
    || !record.target || typeof record.target !== 'object' || Array.isArray(record.target)) {
    throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid hyperlink from ${source}`);
  }
  const target = record.target as Record<string, any>;
  if (!['url', 'email', 'sheet', 'name'].includes(target.kind)) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid hyperlink target from ${source}`);
  if (target.kind === 'url' && (typeof target.url !== 'string' || target.url.trim() === '')) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid URL hyperlink from ${source}`);
  if (target.kind === 'email' && (typeof target.address !== 'string' || target.address.trim() === '')) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid email hyperlink from ${source}`);
  if (target.kind === 'sheet' && (typeof target.sheetId !== 'string' || target.sheetId.trim() === '')) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid sheet hyperlink from ${source}`);
  if (target.kind === 'name' && (typeof target.name !== 'string' || target.name.trim() === '')) throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid defined-name hyperlink from ${source}`);
  if (record.tooltip !== undefined && typeof record.tooltip !== 'string') throw new Error(`SNAPSHOT_MIGRATION_CONFLICT: invalid hyperlink tooltip from ${source}`);
  return structuredClone(record) as CellHyperlink;
}

function emptyReviewSnapshot(): ReviewStoreSnapshot {
  return { notesByCell: {}, notesById: {}, threadIdsByCell: {}, threadsById: {} };
}

function sameReviewValue(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function migrateLegacyReview(sheet: Record<string, any>): void {
  const review = emptyReviewSnapshot();
  const putNote = (row: number, column: number, note: Record<string, any>): void => {
    if (!Number.isSafeInteger(row) || row < 0 || row > 1_048_575 || !Number.isSafeInteger(column) || column < 0 || column > 16_383
      || !note || typeof note !== 'object' || Array.isArray(note) || typeof note.id !== 'string' || !note.id.trim()) {
      throw new Error(`REVIEW_MIGRATION_CONFLICT: invalid legacy note on ${sheet.id}`);
    }
    const key = `${row}:${column}`;
    const currentId = review.notesByCell[key];
    if (currentId !== undefined) {
      if (currentId !== note.id || !sameReviewValue(review.notesById[currentId], note)) throw new Error(`REVIEW_MIGRATION_CONFLICT: notes disagree at ${sheet.id}!${key}`);
      return;
    }
    const identityOwner = Object.entries(review.notesByCell).find(([, id]) => id === note.id)?.[0];
    if (identityOwner !== undefined && identityOwner !== key) throw new Error(`REVIEW_MIGRATION_CONFLICT: note ${note.id} belongs to ${identityOwner}`);
    if (review.notesById[note.id] && !sameReviewValue(review.notesById[note.id], note)) throw new Error(`REVIEW_MIGRATION_CONFLICT: note identity ${note.id} has different content`);
    review.notesByCell[key] = note.id;
    review.notesById[note.id] = structuredClone(note) as ReviewStoreSnapshot['notesById'][string];
  };
  const putThread = (thread: Record<string, any>): void => {
    if (!thread || typeof thread !== 'object' || Array.isArray(thread) || typeof thread.id !== 'string' || !thread.id.trim()
      || thread.sheetId !== sheet.id || !Number.isSafeInteger(thread.row) || thread.row < 0 || thread.row > 1_048_575
      || !Number.isSafeInteger(thread.column) || thread.column < 0 || thread.column > 16_383) {
      throw new Error(`REVIEW_MIGRATION_CONFLICT: invalid legacy comment ${thread?.id ?? '<unknown>'} on ${sheet.id}`);
    }
    const current = review.threadsById[thread.id];
    if (current) {
      if (!sameReviewValue(current, thread)) throw new Error(`REVIEW_MIGRATION_CONFLICT: comment identity ${thread.id} has different content`);
      return;
    }
    review.threadsById[thread.id] = structuredClone(thread) as ReviewStoreSnapshot['threadsById'][string];
    const key = `${thread.row}:${thread.column}`;
    review.threadIdsByCell[key] ??= [];
    if (!review.threadIdsByCell[key]!.includes(thread.id)) review.threadIdsByCell[key]!.push(thread.id);
  };
  if (sheet.notes !== undefined && !Array.isArray(sheet.notes)) throw new Error(`REVIEW_MIGRATION_CONFLICT: notes on ${sheet.id} must be an array`);
  if (sheet.commentThreads !== undefined && !Array.isArray(sheet.commentThreads)) throw new Error(`REVIEW_MIGRATION_CONFLICT: comments on ${sheet.id} must be an array`);
  for (const entry of Array.isArray(sheet.notes) ? sheet.notes : []) {
    const row = Number(entry.row);
    const column = Number(entry.column);
    if (!Number.isSafeInteger(row) || row < 0 || !Number.isSafeInteger(column) || column < 0 || !entry.note?.id) throw new Error(`REVIEW_MIGRATION_CONFLICT: invalid legacy note on ${sheet.id}`);
    putNote(row, column, entry.note);
  }
  for (const thread of Array.isArray(sheet.commentThreads) ? sheet.commentThreads : []) putThread(thread);
  for (const [rowKey, rowValue] of Object.entries(sheet.cells ?? {})) {
    const row = Number(rowKey);
    if (!Number.isSafeInteger(row) || row < 0 || !rowValue || typeof rowValue !== 'object' || Array.isArray(rowValue)) {
      throw new Error(`REVIEW_MIGRATION_CONFLICT: invalid legacy row on ${sheet.id}`);
    }
    for (const [columnKey, cell] of Object.entries(rowValue as Record<string, any>)) {
      const column = Number(columnKey);
      if (!Number.isSafeInteger(column) || column < 0 || !cell || typeof cell !== 'object' || Array.isArray(cell)) {
        throw new Error(`REVIEW_MIGRATION_CONFLICT: invalid legacy cell on ${sheet.id}!${rowKey}:${columnKey}`);
      }
      if (Object.prototype.hasOwnProperty.call(cell, 'note')) putNote(row, column, cell.note);
      if (Object.prototype.hasOwnProperty.call(cell, 'comment')) {
        if (cell.comment?.sheetId !== undefined && cell.comment.sheetId !== sheet.id) {
          throw new Error(`REVIEW_MIGRATION_CONFLICT: comment ${cell.comment?.id ?? '<unknown>'} targets ${cell.comment.sheetId}, expected ${sheet.id}`);
        }
        putThread({ ...cell.comment, sheetId: sheet.id, row, column, replies: Array.isArray(cell.comment?.replies) ? cell.comment.replies : [] });
      }
      delete cell.note;
      delete cell.comment;
    }
  }
  sheet.review = review;
  delete sheet.notes;
  delete sheet.commentThreads;
}

function migrateLegacyFilter(sheet: Record<string, any>): void {
  if (!sheet.filter || typeof sheet.filter !== 'object') return;
  const legacy = sheet.filter as Record<string, any>;
  const columns: Record<string, any> = {};
  for (const [key, condition] of Object.entries(legacy.criteria ?? {})) {
    const item = condition as Record<string, any>;
    const selectedValues = Array.isArray(item.selectedValues) ? item.selectedValues : undefined;
    columns[key] = {
      column: Number(key),
      showButton: true,
      hiddenButton: false,
      criterion: selectedValues
        ? { kind: 'values', values: selectedValues, includeBlank: selectedValues.some((value: unknown) => value === '' || value === null) }
        : item.conditionOperator
          ? { kind: 'custom', join: 'and', conditions: [{ operator: item.conditionOperator, value: item.conditionValue ?? null }] }
          : undefined,
    };
  }
  sheet.autoFilter = { sheetId: legacy.sheetId ?? sheet.id, range: legacy.range, columns };
  delete sheet.filter;
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
  if (Object.prototype.hasOwnProperty.call(snapshot as object, 'definedNames')) {
    throw new Error('Workbook snapshot contains the removed definedNames field');
  }
  if (!Array.isArray(snapshot.definedNameModels)) {
    throw new Error('Workbook snapshot definedNameModels must be an array');
  }
  validateCanonicalDefinedNames(snapshot.definedNameModels);
  if (!snapshot.dataModel || !Array.isArray(snapshot.dataModel.sources) || !Array.isArray(snapshot.dataModel.tables)
    || !Array.isArray(snapshot.dataModel.relationships) || !Array.isArray(snapshot.dataModel.views)) {
    throw new Error('Workbook snapshot dataModel is invalid');
  }
  if (!isWorkbookEditingOptions(snapshot.editingOptions)) throw new Error('Workbook snapshot editingOptions are invalid');
  if (!snapshot.dimensionMetrics || !snapshot.dimensionMetrics.normalFontFamily.trim()
    || !Number.isFinite(snapshot.dimensionMetrics.normalFontSizePx) || snapshot.dimensionMetrics.normalFontSizePx <= 0
    || !Number.isFinite(snapshot.dimensionMetrics.maximumDigitWidthPx) || snapshot.dimensionMetrics.maximumDigitWidthPx <= 0) throw new Error('Workbook snapshot dimensionMetrics is invalid');
  if (snapshot.collationContext) {
    const context = snapshot.collationContext;
    if (!context.cultureId.trim() || !Array.isArray(context.typeOrder) || context.typeOrder.length !== 5 || new Set(context.typeOrder).size !== 5
      || !Array.isArray(context.customLists) || context.customLists.some((list) => !Array.isArray(list) || list.some((entry) => typeof entry !== 'string'))) {
      throw new Error('Workbook snapshot collationContext is invalid');
    }
  }
  if (!isWorkbookCalculationSettings(snapshot.calculationSettings)) {
    throw new Error('Workbook calculation settings are invalid');
  }
  if (snapshot.theme) {
    if (!snapshot.theme.id.trim() || !snapshot.theme.colors || typeof snapshot.theme.colors !== 'object' || Array.isArray(snapshot.theme.colors)
      || Object.entries(snapshot.theme.colors).some(([key, color]) => !key.trim() || typeof color !== 'string' || !/^#[0-9a-f]{6}$/i.test(color))) {
      throw new Error('Workbook theme is invalid');
    }
  }
  const pivotIds = new Set<string>();
  for (const sheet of snapshot.sheets) {
    if (!['worksheet', 'table-sheet', 'gantt-sheet', 'report-sheet'].includes(sheet.kind)) throw new Error('Worksheet kind is invalid');
    if (sheet.kind === 'table-sheet' && !sheet.tableSheet) throw new Error('TableSheet definition is required');
    if (sheet.kind === 'gantt-sheet' && !sheet.ganttSheet) throw new Error('GanttSheet definition is required');
    if (sheet.kind === 'report-sheet' && !sheet.reportSheet) throw new Error('ReportSheet definition is required');
    validateReviewSnapshot(sheet.review, sheet.id);
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
    for (const table of tableFilters) {
      const filter = table.autoFilter!;
      if (filter.sheetId !== sheet.id || filter.range.sheetId !== sheet.id || !sameRange(filter.range, table.range)) {
        throw new Error('Table AutoFilter must equal its Table range');
      }
      for (const [key, column] of Object.entries(filter.columns)) {
        if (!Number.isSafeInteger(Number(key)) || column.column !== Number(key)
          || column.column < filter.range.startColumn || column.column > filter.range.endColumn) {
          throw new Error('Table AutoFilter column identity is invalid');
        }
      }
    }
    if (sheet.autoFilter && tableFilters.some((table) => rangesOverlap(sheet.autoFilter!.range, table.autoFilter!.range))) {
      throw new Error('Worksheet and Table AutoFilter ranges cannot overlap');
    }
    validateCanonicalHyperlinks(sheet);
    for (const pivot of sheet.pivots) {
      if (!pivot.id.trim() || pivotIds.has(pivot.id)) throw new Error(`Pivot identity is duplicated or empty: ${pivot.id}`);
      pivotIds.add(pivot.id);
      canonicalizePivotDefinition(pivot);
    }
    for (const payload of Object.values(sheet.drawingPayloads)) {
      const runtimeKind = (payload as { kind?: string }).kind;
      if (runtimeKind === 'data-chart') throw new Error('MIGRATION_REQUIRED: legacy data-chart payloads must be converted by the offline snapshot migrator');
      if (runtimeKind === 'chart' && !['worksheet-ranges', 'pivot', 'table', 'report-range'].includes(String((payload as { source?: { kind?: string } }).source?.kind))) {
        throw new Error('MIGRATION_REQUIRED: chart payload is missing the canonical source discriminator');
      }
      if (payload.kind === 'chart') validateChartSnapshotPayload(payload, snapshot, sheet.id);
      if (payload.kind === 'camera') validateDrawingSourceRange(payload.sourceRange, snapshot, 'Camera');
      if (payload.kind === 'image' && !isAssetRef(payload.asset)) throw new Error(`Drawing image asset is invalid: ${payload.asset}`);
      if (payload.kind === 'screenshot') {
        if (!isScreenshotDrawingPayload(payload)) throw new Error('Drawing screenshot payload is invalid');
        validateDrawingSourceRange(payload.sourceRange, snapshot, 'Screenshot');
      }
      if (payload.kind === 'icon' && !isIconDrawingPayload(payload)) throw new Error('Drawing icon payload is invalid');
      if (payload.kind === 'model3d' && !isModel3dDrawingPayload(payload)) throw new Error('Drawing 3D model payload is invalid');
      if (payload.kind === 'smartart' && !isSmartArtDrawingPayload(payload)) throw new Error('Drawing SmartArt payload is invalid');
      if (payload.kind === 'wordart' && !isWordArtDrawingPayload(payload)) throw new Error('Drawing WordArt payload is invalid');
      if (payload.kind === 'signature-line' && !isSignatureLineDrawingPayload(payload)) throw new Error('Drawing signature-line payload is invalid');
      if (payload.kind === 'embedded-object' && !isEmbeddedObjectDrawingPayload(payload)) throw new Error('Drawing embedded-object payload is invalid');
      if (payload.kind === 'equation' && !isEquationDrawingPayload(payload)) throw new Error('Drawing equation payload is invalid');
    }
    for (const row of Object.values(sheet.cells)) {
      for (const cell of Object.values(row)) {
        if (Object.prototype.hasOwnProperty.call(cell, 'hyperlink') || Object.prototype.hasOwnProperty.call(cell, 'hyperlinkDetail')) {
          throw new Error(`Cell ${sheet.id} contains legacy hyperlink metadata`);
        }
        if ('note' in cell || 'comment' in cell) throw new Error(`Cell ${sheet.id} contains legacy review metadata`);
        if (cell.phonetic && !isCellPhoneticMetadata(cell.phonetic)) throw new Error(`Cell ${sheet.id} contains invalid phonetic metadata`);
        if (cell.presentation?.kind === 'image' && !isAssetRef(cell.presentation.asset)) throw new Error('Cell image asset is invalid');
      }
    }
    assertCanonicalDataRegionOverlays(sheet);
  }
  for (const sheet of snapshot.sheets) {
    for (const drawing of sheet.drawings) {
      const payload = sheet.drawingPayloads[drawing.payloadId];
      if (!payload) throw new Error(`Drawing payload is missing: ${drawing.payloadId}`);
      if (payload.kind !== 'chart' && payload.kind !== 'slicer' && payload.kind !== 'timeline') continue;
      const pivotId = payload.kind === 'chart' ? (payload.source.kind === 'pivot' ? payload.source.pivotId : undefined) : payload.pivotId;
      if (payload.kind === 'chart' && pivotId === undefined) continue;
      if (typeof pivotId !== 'string' || !pivotId.trim() || !pivotIds.has(pivotId)) {
        throw new Error(`Drawing ${drawing.id} references missing Pivot: ${pivotId ?? ''}`);
      }
      if ((payload.kind === 'slicer' || payload.kind === 'timeline') && payload.connections) {
        const primary = snapshot.sheets.flatMap((candidate) => candidate.pivots).find((pivot) => pivot.id === payload.pivotId);
        if (!primary) throw new Error(`Drawing ${drawing.id} references missing primary Pivot: ${payload.pivotId}`);
        const primaryField = primary.fieldCatalog.fields.find((field) => field.fieldId === payload.fieldId);
        if (!primaryField) throw new Error(`Drawing ${drawing.id} references missing primary field: ${payload.fieldId}`);
        for (const connection of payload.connections) {
          if (!pivotIds.has(connection.pivotId)) throw new Error(`Drawing ${drawing.id} references missing connected Pivot: ${connection.pivotId}`);
          if (connection.pivotId === payload.pivotId) throw new Error(`Drawing ${drawing.id} repeats its primary Pivot connection`);
          const target = snapshot.sheets.flatMap((candidate) => candidate.pivots).find((pivot) => pivot.id === connection.pivotId);
          if (!target || pivotSourceIdentity(target.source) !== connection.sourceKey) throw new Error(`Drawing ${drawing.id} has an incompatible Pivot source connection: ${connection.pivotId}`);
          if (connection.sourceKey !== pivotSourceIdentity(primary.source)) throw new Error(`Drawing ${drawing.id} connects incompatible Pivot caches: ${connection.pivotId}`);
          const targetField = target.fieldCatalog.fields.find((field) => field.fieldId === connection.fieldId);
          if (!targetField || targetField.ordinal !== primaryField.ordinal || targetField.dataType !== primaryField.dataType || targetField.name !== primaryField.name) {
            throw new Error(`Drawing ${drawing.id} references an incompatible connected field: ${connection.fieldId}`);
          }
          if (payload.kind === 'timeline' && (primaryField.dataType !== 'date' || targetField.dataType !== 'date')) throw new Error(`Drawing ${drawing.id} Timeline field is not date-semantic`);
        }
      }
    }
  }
  const chartTableIds = new Set(snapshot.sheets.flatMap((sheet) => Object.values(sheet.drawingPayloads).flatMap((payload) => {
    if (payload.kind !== 'chart' || payload.source.kind !== 'table') return [];
    return [payload.source.tableId];
  })));
  for (const sheet of snapshot.sheets) {
    for (const payload of Object.values(sheet.drawingPayloads)) {
      if (payload.kind === 'chart' && payload.source.kind === 'report-range') validateDrawingSourceRange(payload.source.range, snapshot, 'Chart');
    }
  }
  for (const table of snapshot.dataModel.tables) {
    if (table.sourceRange && chartTableIds.has(table.id)) validateDrawingSourceRange(table.sourceRange, snapshot, 'Chart table');
  }
  const canonical = structuredClone(snapshot);
  for (const sheet of canonical.sheets) {
    sheet.drawingGroups ??= [];
    sheet.snapSettings = canonicalSnapSettings(sheet.snapSettings);
    validateDrawingGraph(sheet);
  }
  canonical.cellStyleTemplates ??= [];
  const templateIds = new Set<string>();
  for (const template of canonical.cellStyleTemplates) {
    if (!template.id.trim() || !template.name.trim() || templateIds.has(template.id)) throw new Error('Cell style template identity is invalid');
    if (template.style.indent !== undefined && (!Number.isInteger(template.style.indent) || template.style.indent < 0 || template.style.indent > 250)) {
      throw new Error('Cell style template indent is invalid');
    }
    if (template.editor && !isCellEditorConfig(template.editor)) {
      throw new Error('Cell style template editor is invalid');
    }
    templateIds.add(template.id);
  }
  return canonical;
}

function validateReviewSnapshot(review: ReviewStoreSnapshot, sheetId: string): void {
  if (!review || typeof review !== 'object' || Array.isArray(review)
    || !review.notesByCell || !review.notesById || !review.threadIdsByCell || !review.threadsById) {
    throw new Error(`ReviewStore snapshot is invalid for sheet ${sheetId}`);
  }
  const noteIds = new Set<string>();
  for (const [id, note] of Object.entries(review.notesById)) {
    if (!id.trim() || !note || typeof note !== 'object' || Array.isArray(note) || note.id !== id) throw new Error(`Review note identity is invalid for ${sheetId}: ${id}`);
    noteIds.add(id);
  }
  const indexedNotes = new Set<string>();
  for (const [key, id] of Object.entries(review.notesByCell)) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (!/^\d+:\d+$/.test(key) || !Number.isSafeInteger(row) || row < 0 || row > 1_048_575 || !Number.isSafeInteger(column) || column < 0 || column > 16_383
      || !noteIds.has(id) || indexedNotes.has(id)) throw new Error(`Review note index is invalid for ${sheetId}!${key}`);
    indexedNotes.add(id);
  }
  if (indexedNotes.size !== noteIds.size) throw new Error(`Review note store contains an unindexed note on ${sheetId}`);
  const threadIds = new Set(Object.keys(review.threadsById));
  const indexedThreads = new Set<string>();
  for (const [key, ids] of Object.entries(review.threadIdsByCell)) {
    const [rowText, columnText] = key.split(':');
    const row = Number(rowText);
    const column = Number(columnText);
    if (!/^\d+:\d+$/.test(key) || !Number.isSafeInteger(row) || row < 0 || row > 1_048_575 || !Number.isSafeInteger(column) || column < 0 || column > 16_383
      || !Array.isArray(ids) || new Set(ids).size !== ids.length) throw new Error(`Review thread index is invalid for ${sheetId}!${key}`);
    for (const id of ids) {
      const thread = review.threadsById[id];
      if (!thread || thread.id !== id || thread.sheetId !== sheetId || thread.row !== row || thread.column !== column || indexedThreads.has(id)) {
        throw new Error(`Review thread index is invalid for ${sheetId}!${key}`);
      }
      indexedThreads.add(id);
    }
  }
  for (const [id, thread] of Object.entries(review.threadsById)) {
    if (!id.trim() || !thread || typeof thread !== 'object' || Array.isArray(thread) || thread.id !== id || thread.sheetId !== sheetId
      || !Number.isSafeInteger(thread.row) || thread.row < 0 || thread.row > 1_048_575
      || !Number.isSafeInteger(thread.column) || thread.column < 0 || thread.column > 16_383) {
      throw new Error(`Review thread identity is invalid for ${sheetId}: ${id}`);
    }
  }
  if (indexedThreads.size !== threadIds.size) throw new Error(`Review store contains an unindexed thread on ${sheetId}`);
}

function validateCanonicalDefinedNames(models: readonly DefinedNameModel[]): void {
  const identities = new Set<string>();
  for (const [index, value] of models.entries()) {
    if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error(`Defined name ${index} is invalid`);
    const allowed = new Set(['name', 'formula', 'scope', 'sheetId', 'anchor', 'hidden', 'comment']);
    if (Object.keys(value as object).some((key) => !allowed.has(key))) throw new Error(`Defined name ${index} contains unsupported fields`);
    let normalized: DefinedNameModel;
    try {
      normalized = normalizeDefinedNameModel(value);
    } catch (error) {
      throw new Error(`Defined name ${index} is invalid`, { cause: error });
    }
    const identity = `${normalized.scope}:${normalized.sheetId ?? ''}:${normalized.name.toUpperCase()}`;
    if (identities.has(identity)) throw new Error(`Defined name identity is duplicated: ${identity}`);
    identities.add(identity);
  }
}

function validateCanonicalHyperlinks(sheet: SheetSnapshot): void {
  if (sheet.hyperlinks === undefined) return;
  if (!Array.isArray(sheet.hyperlinks)) throw new Error(`Worksheet ${sheet.id} hyperlinks must be an array`);
  const coordinates = new Set<string>();
  for (const [index, entry] of sheet.hyperlinks.entries()) {
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)
      || !Number.isSafeInteger(entry.row) || entry.row < 0 || entry.row > 1_048_575
      || !Number.isSafeInteger(entry.column) || entry.column < 0 || entry.column > 16_383) {
      throw new Error(`Worksheet ${sheet.id} hyperlink ${index} is invalid`);
    }
    const key = `${entry.row}:${entry.column}`;
    if (coordinates.has(key)) throw new Error(`Worksheet ${sheet.id} contains duplicate hyperlink at ${key}`);
    coordinates.add(key);
    try {
      normalizeStoredHyperlink(entry.hyperlink, `worksheet ${sheet.id}!${key}`);
    } catch (error) {
      throw new Error(`Worksheet ${sheet.id} hyperlink ${index} is invalid`, { cause: error });
    }
  }
}

function containsLegacyImageDataUrl(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  if (Array.isArray(value)) return value.some(containsLegacyImageDataUrl);
  const record = value as Record<string, unknown>;
  if (typeof record.src === 'string' && record.src.startsWith('data:image/')) return true;
  return Object.values(record).some(containsLegacyImageDataUrl);
}

function validateChartSnapshotPayload(payload: ChartDrawingPayload, snapshot: WorkbookSnapshot, ownerSheetId: string): void {
  if (!isChartSubtypeForType(payload.chartType, payload.subtype)) throw new Error(`Chart subtype ${payload.subtype} does not belong to ${payload.chartType}`);
  if (payload.source.kind === 'worksheet-ranges') {
    if (payload.source.ranges.length === 0) throw new Error(`Chart ${payload.chartId} requires at least one source range`);
    for (const range of payload.source.ranges) validateDrawingSourceRange(range, snapshot, `Chart ${payload.chartId}`);
  } else if (payload.source.kind === 'report-range') {
    validateDrawingSourceRange(payload.source.range, snapshot, `Chart ${payload.chartId}`);
  } else if (payload.source.kind === 'table') {
    const tableSource = payload.source;
    const table = snapshot.dataModel.tables.find((candidate) => candidate.id === tableSource.tableId);
    if (!table) throw new Error(`Chart ${payload.chartId} references missing table ${tableSource.tableId}`);
    if (table.sourceRange) validateDrawingSourceRange(table.sourceRange, snapshot, `Chart ${payload.chartId} table`);
    for (const area of ['values', 'category', 'details', 'color', 'size', 'tooltip', 'filter'] as const) {
      if (!Array.isArray(tableSource.bindings[area])) throw new Error(`Chart ${payload.chartId} binding area is invalid: ${area}`);
      for (const binding of tableSource.bindings[area]) {
        if (!binding.fieldId.trim() || binding.area !== area) throw new Error(`Chart ${payload.chartId} binding area is inconsistent: ${area}`);
        if (!table.fields.some((field) => field.id === binding.fieldId)) throw new Error(`Chart ${payload.chartId} references missing table field ${binding.fieldId}`);
      }
    }
  } else if (!payload.source.pivotId.trim()) {
    throw new Error(`Chart ${payload.chartId} Pivot source is invalid`);
  }
  if (payload.dataOrientation === 'rows' && payload.series?.some((series) => series.range.startRow === series.range.endRow)) {
    throw new Error(`Chart ${payload.chartId} row-oriented series must contain more than one column`);
  }
  if (payload.categoryRange) validateDrawingSourceRange(payload.categoryRange, snapshot, `Chart ${payload.chartId} category`);
  for (const series of payload.series ?? []) {
    validateDrawingSourceRange(series.range, snapshot, `Chart ${payload.chartId} series`);
    for (const range of [series.xRange, series.yRange, series.sizeRange, series.categoryRange, series.errorBars?.plusRange, series.errorBars?.minusRange, series.stockRoles?.open, series.stockRoles?.high, series.stockRoles?.low, series.stockRoles?.close, series.stockRoles?.volume, series.dataLabels?.valuesFromCells]) {
      if (range) validateDrawingSourceRange(range, snapshot, `Chart ${payload.chartId} series binding`);
    }
    if (series.chartType && series.subtype && !isChartSubtypeForType(series.chartType, series.subtype)) throw new Error(`Chart ${payload.chartId} series subtype ${series.subtype} does not belong to ${series.chartType}`);
    const seriesType = series.chartType ?? payload.chartType;
    if (payload.chartType === 'combo' && !['column', 'bar', 'line', 'area'].includes(seriesType)) {
      throw new Error(`Chart ${payload.chartId} combo series type ${seriesType} is not supported by the canonical combo layout`);
    }
    if (payload.chartType !== 'combo' && series.chartType && series.chartType !== payload.chartType) {
      throw new Error(`Chart ${payload.chartId} cannot mix ${payload.chartType} with ${series.chartType} series`);
    }
    if (seriesType === 'scatter' || seriesType === 'bubble') {
      if (!series.xRange || !series.yRange || (seriesType === 'bubble' && !series.sizeRange)) {
        throw new Error(`Chart ${payload.chartId} ${seriesType} series requires explicit X/Y${seriesType === 'bubble' ? '/Size' : ''} ranges`);
      }
    }
    if (seriesType === 'stock' && (!series.stockRoles?.high || !series.stockRoles.low || !series.stockRoles.close)) {
      throw new Error(`Chart ${payload.chartId} stock series requires explicit High/Low/Close ranges`);
    }
    if (series.errorBars?.type === 'custom' && (!series.errorBars.plusRange || !series.errorBars.minusRange)) {
      throw new Error(`Chart ${payload.chartId} custom error bars require explicit plus and minus ranges`);
    }
  }
  if (payload.chartType === 'combo' && (!payload.series?.length || payload.series.some((series) => !series.chartType))) throw new Error(`Chart ${payload.chartId} Combo requires an explicit type for every series`);
  if (ownerSheetId.trim() === '') throw new Error('Chart owner sheet is required');
}

function validateDrawingSourceRange(range: RangeRef, snapshot: WorkbookSnapshot, label: string): void {
  const sheet = snapshot.sheets.find((candidate) => candidate.id === range.sheetId);
  const validCoordinates = [range.startRow, range.endRow, range.startColumn, range.endColumn]
    .every((coordinate) => Number.isSafeInteger(coordinate) && coordinate >= 0);
  if (!sheet || !validCoordinates || range.startRow > range.endRow || range.startColumn > range.endColumn
    || range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) {
    throw new Error(`${label} source range is outside its worksheet bounds`);
  }
  const rows = range.endRow - range.startRow + 1;
  const columns = range.endColumn - range.startColumn + 1;
  if (rows > MAX_DRAWING_SOURCE_CELLS || columns > MAX_DRAWING_SOURCE_CELLS || rows * columns > MAX_DRAWING_SOURCE_CELLS) {
    throw new Error(`${label} source range exceeds the ${MAX_DRAWING_SOURCE_CELLS}-cell rendering limit`);
  }
}

function rangesOverlap(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow && right.startRow <= left.endRow
    && left.startColumn <= right.endColumn && right.startColumn <= left.endColumn;
}

function sameRange(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId && left.startRow === right.startRow && left.endRow === right.endRow
    && left.startColumn === right.startColumn && left.endColumn === right.endColumn;
}

export function loadWorkbookFromSnapshot(snapshot: WorkbookSnapshot): WorkbookModelClass {
  return WorkbookModelClass.fromSnapshot(assertCanonicalWorkbookSnapshot(snapshot));
}

export function createWorkbookSnapshot(workbook: WorkbookModel): WorkbookSnapshot {
  return workbook.snapshot();
}
