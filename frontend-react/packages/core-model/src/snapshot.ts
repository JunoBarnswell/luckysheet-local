import type { DefinedNameModel, SheetSnapshot, RangeRef, CellStyleTemplate, UnitId, WorkbookModel, WorkbookTheme } from './index';
import type { PrintDocumentSnapshot, QueryDefinitionSnapshot } from './workbook-state';
import { WorkbookModel as WorkbookModelClass } from './index';
import { MAX_DRAWING_SOURCE_CELLS } from './generated-workbook-limits';
import { canonicalizePivotDefinition, pivotSourceIdentity } from './pivot';
import { isAssetRef } from './asset';
import { canonicalSnapSettings, validateDrawingGraph } from './drawing-planner';
import type { ReviewStoreSnapshot } from './review-store';
import { DEFAULT_WORKBOOK_CALCULATION_SETTINGS, isWorkbookCalculationSettings, type WorkbookCalculationSettings, type WorkbookCollationContext } from '@react-sheets/formula-engine';

/**
 * The single persisted/transport snapshot contract. Floating objects are
 * represented only by the canonical drawing collection and payload map.
 * There is deliberately no versioned sibling or legacy union in production.
 */
export interface WorkbookSnapshot {
  schema: 'WorkbookSnapshot';
  /** Canonical persisted schema revision. Non-matching snapshots are rejected. */
  version: 8;
  unitId: UnitId;
  name: string;
  dimensionMetrics: WorkbookDimensionMetrics;
  /** Workbook-owned deterministic ordering semantics for values and query keys. */
  collationContext?: WorkbookCollationContext;
  /** Workbook-owned calculation policy; formula workers consume this exact state. */
  calculationSettings: WorkbookCalculationSettings;
  /** Workbook-owned theme identity and resolved colors. */
  theme?: WorkbookTheme;
  definedNames?: Record<string, string>;
  definedNameModels?: DefinedNameModel[];
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

export const WORKBOOK_SNAPSHOT_SCHEMA_REVISION = 8 as const;

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
  if (input.version === 7 && Array.isArray(input.sheets)) {
    input.version = WORKBOOK_SNAPSHOT_SCHEMA_REVISION;
    for (const sheet of input.sheets as Array<Record<string, any>>) migrateLegacyReview(sheet);
    return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
  }
  if (input.version === 6 && Array.isArray(input.sheets)) {
    input.version = WORKBOOK_SNAPSHOT_SCHEMA_REVISION;
    input.calculationSettings = input.calculationSettings ?? structuredClone(DEFAULT_WORKBOOK_CALCULATION_SETTINGS);
    if (containsLegacyImageDataUrl(input)) throw new Error('ASSET_MIGRATION_REQUIRED: legacy image data must be assetized before runtime load');
    for (const sheet of input.sheets as Array<Record<string, any>>) migrateLegacyReview(sheet);
    return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
  }
  if (input.version === 5 && Array.isArray(input.sheets)) {
    input.version = WORKBOOK_SNAPSHOT_SCHEMA_REVISION;
    input.calculationSettings = input.calculationSettings ?? structuredClone(DEFAULT_WORKBOOK_CALCULATION_SETTINGS);
    for (const sheet of input.sheets as Array<Record<string, any>>) migrateLegacyReview(sheet);
    return assertCanonicalWorkbookSnapshot(input as WorkbookSnapshot);
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
    review.notesById[note.id] = structuredClone(note);
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
    review.threadsById[thread.id] = structuredClone(thread);
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
  if (!snapshot.dataModel || !Array.isArray(snapshot.dataModel.sources) || !Array.isArray(snapshot.dataModel.tables)
    || !Array.isArray(snapshot.dataModel.relationships) || !Array.isArray(snapshot.dataModel.views)) {
    throw new Error('Workbook snapshot dataModel is invalid');
  }
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
    for (const pivot of sheet.pivots) {
      if (!pivot.id.trim() || pivotIds.has(pivot.id)) throw new Error(`Pivot identity is duplicated or empty: ${pivot.id}`);
      pivotIds.add(pivot.id);
      canonicalizePivotDefinition(pivot);
    }
    for (const payload of Object.values(sheet.drawingPayloads)) {
      if (payload.kind === 'camera') validateDrawingSourceRange(payload.sourceRange, snapshot, 'Camera');
      if (payload.kind === 'image' && !isAssetRef(payload.asset)) throw new Error(`Drawing image asset is invalid: ${payload.asset}`);
    }
    for (const row of Object.values(sheet.cells)) {
      for (const cell of Object.values(row)) {
        if ('note' in cell || 'comment' in cell) throw new Error(`Cell ${sheet.id} contains legacy review metadata`);
        if (cell.presentation?.kind === 'image' && !isAssetRef(cell.presentation.asset)) throw new Error('Cell image asset is invalid');
      }
    }
  }
  for (const sheet of snapshot.sheets) {
    for (const drawing of sheet.drawings) {
      const payload = sheet.drawingPayloads[drawing.payloadId];
      if (!payload) throw new Error(`Drawing payload is missing: ${drawing.payloadId}`);
      if (payload.kind !== 'chart' && payload.kind !== 'slicer' && payload.kind !== 'timeline') continue;
      if (payload.kind === 'chart' && payload.pivotId === undefined) continue;
      if (typeof payload.pivotId !== 'string' || !payload.pivotId.trim() || !pivotIds.has(payload.pivotId)) {
        throw new Error(`Drawing ${drawing.id} references missing Pivot: ${payload.pivotId ?? ''}`);
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
  const dataChartTableIds = new Set(snapshot.sheets.flatMap((sheet) => Object.values(sheet.drawingPayloads)
    .flatMap((payload) => payload.kind === 'data-chart' && payload.source.kind === 'table' ? [payload.source.tableId] : [])));
  for (const sheet of snapshot.sheets) {
    for (const payload of Object.values(sheet.drawingPayloads)) {
      if (payload.kind === 'data-chart' && payload.source.kind === 'report-sheet') validateDrawingSourceRange(payload.source.range, snapshot, 'Data chart');
    }
  }
  for (const table of snapshot.dataModel.tables) {
    if (table.sourceRange && dataChartTableIds.has(table.id)) validateDrawingSourceRange(table.sourceRange, snapshot, 'Data chart table');
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
    if (template.editor && !['text', 'number', 'date', 'list', 'checkbox'].includes(template.editor.kind)) {
      throw new Error('Cell style template editor is invalid');
    }
    if (template.editor?.kind === 'list' && (!Array.isArray(template.editor.values) || template.editor.values.some((value) => !value.trim()))) {
      throw new Error('Cell style template list editor values are invalid');
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
    const [row, column] = key.split(':').map(Number);
    if (!/^\d+:\d+$/.test(key) || !Number.isSafeInteger(row) || row < 0 || row > 1_048_575 || !Number.isSafeInteger(column) || column < 0 || column > 16_383
      || !noteIds.has(id) || indexedNotes.has(id)) throw new Error(`Review note index is invalid for ${sheetId}!${key}`);
    indexedNotes.add(id);
  }
  if (indexedNotes.size !== noteIds.size) throw new Error(`Review note store contains an unindexed note on ${sheetId}`);
  const threadIds = new Set(Object.keys(review.threadsById));
  const indexedThreads = new Set<string>();
  for (const [key, ids] of Object.entries(review.threadIdsByCell)) {
    const [row, column] = key.split(':').map(Number);
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

function containsLegacyImageDataUrl(value: unknown): boolean {
  if (!value || typeof value !== 'object') return false;
  if (Array.isArray(value)) return value.some(containsLegacyImageDataUrl);
  const record = value as Record<string, unknown>;
  if (typeof record.src === 'string' && record.src.startsWith('data:image/')) return true;
  return Object.values(record).some(containsLegacyImageDataUrl);
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
