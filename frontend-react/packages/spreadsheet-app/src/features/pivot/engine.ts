import type {
  PivotAggregateFunction,
  PivotDefinition,
  PivotFieldCatalog,
  PivotFieldDataType,
  PivotFieldDefinition,
  PivotErrorValue,
  PivotFieldPlacement,
  PivotFilter,
  PivotGroup,
  PivotDateGroupUnit,
  PivotGridProjection,
  PivotHitTest,
  PivotLayout,
  PivotCalculatedItem,
  PivotMemberKey,
  PivotModel,
  PivotProjectionCell,
  PivotRefreshState,
  PivotReportFilterSummary,
  PivotReportFilterSummaryEntry,
  PivotResultCell,
  PivotResultNode,
  PivotResultTree,
  PivotResultValueField,
  PivotScalar,
  PivotSort,
  PivotShowAsBaseItem,
  PivotTopBottomMode,
  PivotSource,
  PivotSourceRowPath,
  PivotTarget,
  PivotSlicerDrawingPayload,
  PivotSlicerItemProjection,
  PivotTimelineDrawingPayload,
  PivotValueField,
  ContextHit,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  PIVOT_GRID_PROJECTION_SCHEMA,
  PIVOT_RESULT_TREE_SCHEMA,
  DEFAULT_PIVOT_DISPLAY_OPTIONS,
  DEFAULT_PIVOT_STYLE_OPTIONS,
  DEFAULT_SHEET_COLUMN_COUNT,
  DEFAULT_SHEET_ROW_COUNT,
  MAX_SHEET_COLUMN_COUNT,
  MAX_SHEET_ROW_COUNT,
  createPivotCollator,
  createPivotMemberKey,
  formatPivotMember,
  isPivotError,
  normalizePivotTimelinePeriod,
  pivotMemberKey,
  normalizePivotRefreshPolicy,
  normalizePivotDisplayOptions,
  normalizePivotNumberFormat,
  pivotNumericValue,
  PIVOT_MAX_MEMBER_COUNT,
  pivotTimelineInstant,
  pivotMemberKeyEquals,
  pivotScalarFromMemberKey,
  parsePivotCalculatedItemFormula,
} from '@react-sheets/core-model';
import type { PivotCalculatedItemFormulaToken } from '@react-sheets/core-model';
import type { PivotTimelinePeriodBounds } from '@react-sheets/core-model';
import { collectNameReferences, FormulaEngine, isFormulaError, parseFormula, type FormulaValue } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import { configureWorkbookSpillEnvironments, syncWorkbookSheetTables } from '../../formula-spill-sync';
import {
  assertPivotSourceIndex,
  createPivotSourceIndex,
  inferPivotSourceFieldType,
  pivotSourceColumnValues,
  pivotSourceRowPaths,
  pivotSourceValueAt,
  type PivotSourceFieldInput,
  type PivotSourceIndex,
} from './source-index';

export type PivotSourceTableInput = PivotSourceIndex;
export type { PivotSourceFieldInput, PivotSourceIndex } from './source-index';

interface SourceTable {
  index: PivotSourceIndex;
  fields: PivotSourceFieldInput[];
  fieldOrdinals: ReadonlyMap<string, number>;
  rows: SourceRow[];
}

interface SourceRow {
  source: SourceTable;
  row: number;
  overrides?: ReadonlyMap<string, PivotScalar>;
  pathsOverride?: readonly PivotSourceRowPath[];
}

type SourceField = PivotSourceFieldInput;

function openSourceTable(index: PivotSourceIndex): SourceTable {
  assertPivotSourceIndex(index);
  const fields = index.fields.map((field) => ({ ...field }));
  const fieldOrdinals = new Map(fields.map((field, ordinal) => [field.fieldId, ordinal] as const));
  const table = { index, fields, fieldOrdinals, rows: [] as SourceRow[] };
  table.rows = Array.from({ length: index.rowCount }, (_, row) => ({ source: table, row }));
  return table;
}

function sourceRowValue(row: SourceRow, fieldId: string): PivotScalar {
  if (row.overrides?.has(fieldId)) return row.overrides.get(fieldId) ?? null;
  const ordinal = row.source.fieldOrdinals.get(fieldId);
  return ordinal === undefined ? null : pivotSourceValueAt(row.source.index, ordinal, row.row);
}

function sourceRowPaths(row: SourceRow): readonly PivotSourceRowPath[] {
  return row.pathsOverride ?? pivotSourceRowPaths(row.source.index, row.row);
}

function sourceColumnValues(table: SourceTable, fieldId: string): PivotScalar[] {
  const ordinal = table.fieldOrdinals.get(fieldId);
  return ordinal === undefined ? [] : pivotSourceColumnValues(table.index, ordinal);
}

interface AxisGroup {
  values: PivotScalar[];
  rows: SourceRow[];
  rowSet: Set<SourceRow>;
}

const PIVOT_MAX_RESULT_CELL_COUNT = 250_000;
const PIVOT_MAX_PROVENANCE_REFERENCE_COUNT = 2_000_000;

export interface PivotResultTable {
  headers: string[];
  rows: Array<{ keys: string[]; values: PivotScalar[] }>;
  grandTotal: PivotScalar[];
  tree: PivotResultTree;
}

export interface PivotRevisionKey {
  pivotId: string;
  sourceRevision: string;
  layoutRevision: string;
  filterRevision: string;
}

export interface PivotProjectionSourceState {
  availability: 'loading' | 'ready' | 'missing' | 'error';
  error?: string;
  sourceRevision?: string | number;
}

export interface PivotProjectionOptions {
  sourceState?: PivotProjectionSourceState;
  /** The session's canonical FormulaEngine; required for live spill values. */
  formula?: FormulaEngine;
  /** Explicit refresh failure retained alongside the last-valid projection. */
  refreshError?: string;
  /** Already-normalized command preflight definition; never persisted. */
  canonicalDefinition?: PivotDefinition;
}

interface LastValidPivotProjection {
  projection: PivotGridProjection;
  result: PivotResultTree;
}

interface BlockPivotResultCacheEntry {
  sourceRevision: string;
  layoutRevision: string;
  filterRevision: string;
  result: PivotResultTree;
}

/**
 * Render state is ephemeral and belongs to a workbook session. It is not part
 * of PivotDefinition, WorkbookSnapshot, or collaborative operations. A
 * collision/load failure must never destroy the last successful projection.
 */
const lastValidPivotProjections = new WeakMap<WorkbookModel, Map<string, LastValidPivotProjection>>();
const blockPivotResultCaches = new WeakMap<WorkbookModel, Map<string, BlockPivotResultCacheEntry>>();

const same = (left: PivotScalar, right: PivotScalar): boolean => {
  if ((left == null || left === '') && (right == null || right === '')) return true;
  if (isPivotError(left) || isPivotError(right)) {
    return isPivotError(left) && isPivotError(right) && left.code === right.code;
  }
  return left === right;
};

const display = (value: PivotScalar): string => formatPivotMember(value);

function stableSerialize(value: unknown): string {
  if (value === undefined) return 'undefined';
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableSerialize).join(',')}]`;
  const record = value as Record<string, unknown>;
  return `{${Object.keys(record).filter((key) => record[key] !== undefined).sort().map((key) => `${JSON.stringify(key)}:${stableSerialize(record[key])}`).join(',')}}`;
}

function fingerprint(value: unknown): string {
  const input = stableSerialize(value);
  let hash = 0xcbf29ce484222325n;
  for (let index = 0; index < input.length; index += 1) {
    hash ^= BigInt(input.charCodeAt(index));
    hash = BigInt.asUintN(64, hash * 0x100000001b3n);
  }
  return hash.toString(16).padStart(16, '0');
}

function sourceRevision(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): string {
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') {
    const manifest = workbook.getDataSource(source.dataSourceId);
    return fingerprint({
      source,
      revision: manifest.revision,
      blocks: manifest.blocks.map((block) => ({ id: block.id, checksum: block.checksum, revision: block.revision })),
    });
  }
  const ranges = sourceRanges(workbook, pivot, formula);
  const revisions = ranges.map((range, index) => {
    const sheet = workbook.getSheet(range.sheetId);
    // CellMatrix revision is supplied by the block/data-source implementation
    // when available. Do not scan a whole range merely to build a cache key.
    const revision = (sheet.cells as unknown as { revision?: number }).revision;
    const sourceId = source.kind === 'worksheet-ranges' ? source.ranges[index]?.sourceId : undefined;
    return `${sourceId ?? index}:${range.sheetId}:${revision ?? 'live'}:${sheet.cells.count()}`;
  }).sort();
  const spills = formula ? ranges.map((range) => formula.getSpillsForSheet(range.sheetId)
    .filter((spill) => spill.range.startRow <= range.endRow && range.startRow <= spill.range.endRow
      && spill.range.startColumn <= range.endColumn && range.startColumn <= spill.range.endColumn)
    .map((spill) => ({ anchor: spill.anchor, range: spill.range, values: spill.values, state: spill.state }))) : [];
  return fingerprint({
    source: canonicalPivotSource(source),
    revisions,
    ...(formula && spills.some((entries) => entries.length > 0) ? { spills } : {}),
  });
}

function canonicalPivotSource(source: PivotSource): PivotSource {
  if (source.kind !== 'worksheet-ranges') return structuredClone(source);
  return {
    ...structuredClone(source),
    ranges: [...source.ranges].sort((left, right) => left.sourceId.localeCompare(right.sourceId)),
    relationships: [...source.relationships].sort((left, right) => left.id.localeCompare(right.id)),
  };
}

function linkedFilterDefinitions(workbook: WorkbookModel, pivot: PivotModel): unknown[] {
  return workbook.getSheets().flatMap((sheet) => sheet.drawings.map((drawing) => {
    const payload = sheet.drawingPayloads.get(drawing.payloadId);
    if (!payload || (payload.kind !== 'slicer' && payload.kind !== 'timeline')) return undefined;
    const linked = [payload.pivotId, ...(payload.connections ?? []).map((connection) => connection.pivotId)];
    if (!linked.includes(pivot.id)) return undefined;
    // A newly-created control with its default "all"/empty period has no
    // semantic effect on the aggregate.  Its drawing identity and styling
    // must not invalidate a completed Pivot result; only an active filter or
    // period, together with its report connections, belongs in filterRevision.
    const active = payload.kind === 'slicer'
      ? payload.filter.mode !== 'all' && payload.filter.memberKeys.length > 0
      : payload.period.start !== undefined || payload.period.end !== undefined;
    if (!active) return undefined;
    return {
      kind: payload.kind,
      pivotId: payload.pivotId,
      fieldId: payload.fieldId,
      ...(payload.kind === 'slicer' ? { filter: payload.filter } : { period: payload.period, level: payload.level }),
      connections: payload.connections ?? [],
    };
  })).filter((entry): entry is NonNullable<typeof entry> => Boolean(entry));
}

export function getPivotRevisionKey(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotRevisionKey {
  const { expansion: _expansion, reportLayout: _reportLayout, filters: _filters, ...calculationLayout } = pivot.layout;
  return {
    pivotId: pivot.id,
    sourceRevision: sourceRevision(workbook, pivot, formula),
    // Live member values belong exclusively to sourceRevision. Including them
    // here made an ordinary source edit look like a layout mutation and caused
    // manual-refresh PivotTables to discard their last refreshed result.
    layoutRevision: fingerprint({
      source: canonicalPivotSource(pivot.source),
      fieldCatalog: pivot.fieldCatalog.fields.map(({ fieldId, name, dataType, ordinal }) => ({ fieldId, name, dataType, ordinal })),
      layout: calculationLayout,
    }),
    filterRevision: fingerprint({ filters: pivot.layout.filters, linked: linkedFilterDefinitions(workbook, pivot) }),
  };
}

/** A derived result is reusable only when every canonical Pivot revision matches. */
export function pivotResultMatchesRevision(workbook: WorkbookModel, pivot: PivotModel, result: PivotResultTree | undefined, formula?: FormulaEngine): result is PivotResultTree {
  if (!result || result.pivotId !== pivot.id) return false;
  const revision = getPivotRevisionKey(workbook, pivot, formula);
  return result.sourceRevision === revision.sourceRevision
    && result.layoutRevision === revision.layoutRevision
    && result.filterRevision === revision.filterRevision;
}

/** Manual-refresh PivotTables may reuse source-stale data only when their layout and filters still match. */
export function pivotResultMatchesLayoutAndFilter(workbook: WorkbookModel, pivot: PivotModel, result: PivotResultTree | undefined, formula?: FormulaEngine): result is PivotResultTree {
  if (!result || result.pivotId !== pivot.id) return false;
  const revision = getPivotRevisionKey(workbook, pivot, formula);
  return result.layoutRevision === revision.layoutRevision && result.filterRevision === revision.filterRevision;
}

export function getLastValidPivotResult(workbook: WorkbookModel, pivotId: string): PivotResultTree | undefined {
  const entry = lastValidPivotProjections.get(workbook)?.get(pivotId);
  return entry ? structuredClone(entry.result) : undefined;
}

export function getLastValidPivotProjection(workbook: WorkbookModel, pivotId: string): PivotGridProjection | undefined {
  const entry = lastValidPivotProjections.get(workbook)?.get(pivotId);
  return entry ? structuredClone(entry.projection) : undefined;
}

/** Drop the ephemeral last-valid projection for one pivot or a workbook. */
export function clearPivotResultCache(workbook: WorkbookModel, pivotId?: string): void {
  const cache = lastValidPivotProjections.get(workbook);
  if (!cache) return;
  if (!pivotId) {
    cache.clear();
    return;
  }
  cache.delete(pivotId);
}

function getPivotSource(pivot: PivotModel): PivotSource {
  return pivot.source;
}

function getPivotTarget(pivot: PivotModel): PivotTarget {
  return pivot.target;
}

function sourceIdentity(source: PivotSource, range: RangeRef, ordinal: number, rangeIndex = 0): string {
  if (source.kind === 'table') return `table:${source.tableId}:column:${ordinal}`;
  if (source.kind === 'named-range') return `name:${source.sheetId ?? '*'}:${source.name}:column:${ordinal}`;
  if (source.kind === 'data-source') return `data-source:${source.dataSourceId}:column:${ordinal}`;
  if (source.kind === 'worksheet-ranges') {
    const sourceId = source.ranges[rangeIndex]?.sourceId;
    if (!sourceId) throw new Error(`Worksheet source range ${String(rangeIndex)} has no stable sourceId`);
    return `source:${sourceId}:column:${ordinal}`;
  }
  return `sheet:${range.sheetId}:column:${range.startColumn + ordinal}:range:${rangeIndex}`;
}

/** Stable field identity used by the catalog and all layout references. */
export function getStablePivotFieldId(source: PivotSource, range: RangeRef, ordinal: number, rangeIndex = 0): string {
  return sourceIdentity(source, range, ordinal, rangeIndex);
}

function createPivotFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.primarySheetId, recalculationMode: 'manual' });
  engine.setDefinedNameModels(workbook.definedNameModels);
  configureWorkbookSpillEnvironments(engine, workbook);
  syncWorkbookSheetTables(engine, workbook);
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) engine.setFormula(address, cell.formula);
      else if (cell.value != null) engine.setValue(address, cell.value as never);
    });
  }
  engine.recalculate();
  return engine;
}

function sourceRanges(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): RangeRef[] {
  const source = getPivotSource(pivot);
  if (source.kind === 'worksheet-range') return [source.range];
  if (source.kind === 'worksheet-ranges') return source.ranges.map((sourceRange) => sourceRange.range);
  if (source.kind === 'table') {
    return [resolvePivotTable(workbook, source.tableId).range];
  }
  if (source.kind === 'data-source') {
    const manifest = workbook.getDataSource(source.dataSourceId);
    // A block-backed source is not required to have a worksheet materializing
    // range.  The async block acquisition path owns its rows and returns an
    // explicit source index; callers that need structural overlap receive no
    // fabricated worksheet range.
    return manifest.sourceRange ? [manifest.sourceRange] : [];
  }
  return [resolveNamedRange(workbook, source.name, source.sheetId, formula ?? createPivotFormulaEngine(workbook))];
}

function resolvePivotTable(workbook: WorkbookModel, tableId: string): {
  range: RangeRef;
  fields: Array<{ id: string; name: string }>;
} {
  const workbookTable = workbook.dataModel.tables.get(tableId);
  if (workbookTable?.sourceRange) {
    return {
      range: workbookTable.sourceRange,
      fields: workbookTable.fields.map((field) => ({ id: field.id, name: field.name })),
    };
  }
  const sheetTable = workbook.getSheets()
    .flatMap((sheet) => sheet.sheetTables)
    .find((table) => table.id === tableId || table.name === tableId);
  if (!sheetTable) throw new Error(`Unknown Pivot table source: ${tableId}`);
  return {
    range: sheetTable.range,
    fields: sheetTable.columns.map((column) => ({ id: column.id, name: column.name })),
  };
}

function cellScalar(value: unknown): PivotScalar {
  if (Array.isArray(value) && value.length === 1 && Array.isArray(value[0]) && value[0].length === 1) {
    return cellScalar(value[0][0]);
  }
  if (Array.isArray(value)) throw new Error('Pivot source array value must be resolved through its spill range');
  if (isFormulaError(value)) return { kind: 'error', code: value.code, ...(value.message ? { message: value.message } : {}) };
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || value === null) return value;
  throw new Error(`Unsupported Pivot source value type: ${typeof value}`);
}

function formulaCellValue(formula: FormulaEngine, address: { sheetId: string; row: number; column: number }, fallback: unknown): FormulaValue | unknown {
  // Spill children are derived values and therefore have no authored cell
  // result. Authored cells absent from a session engine still use the model
  // value until the canonical engine receives that input.
  if (formula.getSpillValueAt(address.sheetId, address.row, address.column) !== undefined || formula.getCellResult(address) !== undefined) {
    return formula.getCellValue(address);
  }
  return fallback;
}

function parseColumnLabel(value: string): number {
  let column = 0;
  for (const character of value.toUpperCase()) {
    if (character < 'A' || character > 'Z') throw new Error(`Invalid named range column: ${value}`);
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column - 1;
}

function parseA1Range(formula: string, workbook: WorkbookModel, fallbackSheetId: string, calculator?: FormulaEngine): RangeRef {
  const cleaned = formula.trim().replace(/^=/, '').replace(/^\+/, '');
  const spillReference = cleaned.endsWith('#');
  const reference = spillReference ? cleaned.slice(0, -1) : cleaned;
  const match = reference.match(/^(?:'((?:[^']|'')+)'|([A-Za-z0-9_-]+))?!?\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/);
  if (!match) throw new Error(`Named range is not a worksheet range: ${formula}`);
  const sheetName = (match[1] ?? match[2])?.replace(/''/g, "'");
  const sheet = sheetName ? workbook.getSheetByName(sheetName) : workbook.getSheet(fallbackSheetId);
  if (!sheet) throw new Error(`Named range references unknown worksheet: ${sheetName ?? fallbackSheetId}`);
  const startColumn = parseColumnLabel(match[3]!);
  const startRow = Number(match[4]) - 1;
  const endColumn = match[5] ? parseColumnLabel(match[5]) : startColumn;
  const endRow = match[6] ? Number(match[6]) - 1 : startRow;
  if (startRow < 0 || endRow < startRow || startColumn < 0 || endColumn < startColumn) throw new Error(`Invalid named range: ${formula}`);
  if (spillReference) {
    if (match[5] || !calculator) throw new Error(`Named range spill reference is not resolved: ${formula}`);
    const spill = calculator.getSpillsForSheet(sheet.id).find((candidate) => candidate.anchor.row === startRow && candidate.anchor.column === startColumn);
    if (!spill) throw new Error(`Named range spill anchor has no resolved spill: ${formula}`);
    if (spill.state !== 'ok') throw new Error(`Named range spill is blocked: ${formula}`);
    return structuredClone(spill.range);
  }
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

function resolveNamedRange(workbook: WorkbookModel, name: string, sheetId?: string, calculator?: FormulaEngine): RangeRef {
  const definedName = sheetId === undefined
    ? workbook.getDefinedNameExact(name, 'workbook')
    : workbook.getDefinedNameExact(name, 'sheet', sheetId);
  const formula = definedName?.formula ?? '';
  if (!formula) throw new Error(`Unknown named range: ${name}`);
  return parseA1Range(formula, workbook, sheetId ?? workbook.primarySheetId, calculator);
}

function readRange(sheet: WorksheetModel, range: RangeRef, source: PivotSource, rangeIndex: number, persisted?: PivotFieldCatalog, formula?: FormulaEngine): SourceTable {
  if (formula) {
    for (const spill of formula.getSpillsForSheet(sheet.id)) {
      const intersects = spill.range.startRow <= range.endRow && range.startRow <= spill.range.endRow
        && spill.range.startColumn <= range.endColumn && range.startColumn <= spill.range.endColumn;
      if (intersects && spill.state !== 'ok') throw new Error(`Pivot source intersects blocked spill at ${sheet.id}!${spill.anchor.row}:${spill.anchor.column}`);
    }
  }
  const fields: SourceField[] = [];
  for (let ordinal = 0; ordinal <= range.endColumn - range.startColumn; ordinal += 1) {
    const column = range.startColumn + ordinal;
    const headerCell = sheet.cells.get(range.startRow, column);
    const raw = formula
      ? formulaCellValue(formula, { sheetId: sheet.id, row: range.startRow, column }, headerCell?.formulaValue ?? headerCell?.value ?? null)
      : headerCell?.formulaValue ?? headerCell?.value ?? null;
    const name = raw == null || raw === '' ? `Column ${ordinal + 1}` : String(raw);
    // Ordinal/source-column identity survives a header rename. A changed
    // physical column is a new field, while a changed caption is not.
    const fieldId = sourceIdentity(source, range, ordinal, rangeIndex);
    const persistedField = persisted?.fields.find((field) => field.fieldId === fieldId);
    fields.push({ fieldId: persistedField?.fieldId ?? fieldId, name, ordinal });
  }
  const columnValues = fields.map(() => [] as PivotScalar[]);
  const rowPaths: PivotSourceRowPath[][] = [];
  const sourceId = source.kind === 'worksheet-ranges' ? source.ranges[rangeIndex]?.sourceId : undefined;
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    fields.forEach((field, ordinal) => {
      const cell = sheet.cells.get(row, range.startColumn + ordinal);
      const raw = formula
        ? formulaCellValue(formula, { sheetId: sheet.id, row, column: range.startColumn + ordinal }, cell?.formulaValue ?? cell?.value ?? null)
        : cell?.formulaValue ?? cell?.value ?? null;
      columnValues[ordinal]!.push(cellScalar(raw));
    });
    rowPaths.push([{ ...(sourceId ? { sourceId } : {}), recordId: `${sourceId ?? range.sheetId}:${row}`, sheetId: range.sheetId, row }]);
  }
  return openSourceTable(createPivotSourceIndex({
    columns: fields.map((field, ordinal) => ({ field, values: columnValues[ordinal]! })),
    rowPaths,
  }));
}

export interface PivotSourceAcquireOptions {
  signal?: AbortSignal;
  yieldEveryCells?: number;
  onChunk?: (metrics: { cells: number; durationMs: number }) => void;
}

async function readRangeAsync(
  sheet: WorksheetModel,
  range: RangeRef,
  source: PivotSource,
  rangeIndex: number,
  persisted: PivotFieldCatalog | undefined,
  formula: FormulaEngine | undefined,
  options: PivotSourceAcquireOptions,
): Promise<SourceTable> {
  assertPivotAcquireActive(options.signal);
  if (formula) {
    for (const spill of formula.getSpillsForSheet(sheet.id)) {
      const intersects = spill.range.startRow <= range.endRow && range.startRow <= spill.range.endRow
        && spill.range.startColumn <= range.endColumn && range.startColumn <= spill.range.endColumn;
      if (intersects && spill.state !== 'ok') throw new Error(`Pivot source intersects blocked spill at ${sheet.id}!${spill.anchor.row}:${spill.anchor.column}`);
    }
  }
  const fields: SourceField[] = [];
  for (let ordinal = 0; ordinal <= range.endColumn - range.startColumn; ordinal += 1) {
    const column = range.startColumn + ordinal;
    const headerCell = sheet.cells.get(range.startRow, column);
    const raw = formula
      ? formulaCellValue(formula, { sheetId: sheet.id, row: range.startRow, column }, headerCell?.formulaValue ?? headerCell?.value ?? null)
      : headerCell?.formulaValue ?? headerCell?.value ?? null;
    const name = raw == null || raw === '' ? `Column ${ordinal + 1}` : String(raw);
    const fieldId = sourceIdentity(source, range, ordinal, rangeIndex);
    const persistedField = persisted?.fields.find((field) => field.fieldId === fieldId);
    fields.push({ fieldId: persistedField?.fieldId ?? fieldId, name, ordinal });
  }
  const columnValues = fields.map(() => [] as PivotScalar[]);
  const rowPaths: PivotSourceRowPath[][] = [];
  const sourceId = source.kind === 'worksheet-ranges' ? source.ranges[rangeIndex]?.sourceId : undefined;
  const yieldEvery = Math.max(128, options.yieldEveryCells ?? 2_048);
  let visitedCells = 0;
  let chunkCells = 0;
  let chunkStartedAt = performance.now();
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    for (let ordinal = 0; ordinal < fields.length; ordinal += 1) {
      const cell = sheet.cells.get(row, range.startColumn + ordinal);
      const raw = formula
        ? formulaCellValue(formula, { sheetId: sheet.id, row, column: range.startColumn + ordinal }, cell?.formulaValue ?? cell?.value ?? null)
        : cell?.formulaValue ?? cell?.value ?? null;
      columnValues[ordinal]!.push(cellScalar(raw));
      visitedCells += 1;
      chunkCells += 1;
      if (visitedCells % yieldEvery === 0) {
        options.onChunk?.({ cells: chunkCells, durationMs: performance.now() - chunkStartedAt });
        assertPivotAcquireActive(options.signal);
        await yieldPivotAcquire();
        chunkCells = 0;
        chunkStartedAt = performance.now();
      }
    }
    rowPaths.push([{ ...(sourceId ? { sourceId } : {}), recordId: `${sourceId ?? range.sheetId}:${row}`, sheetId: range.sheetId, row }]);
  }
  if (chunkCells > 0) options.onChunk?.({ cells: chunkCells, durationMs: performance.now() - chunkStartedAt });
  assertPivotAcquireActive(options.signal);
  return openSourceTable(createPivotSourceIndex({
    columns: fields.map((field, ordinal) => ({ field, values: columnValues[ordinal]! })),
    rowPaths,
  }));
}

function assertPivotAcquireActive(signal: AbortSignal | undefined): void {
  if (signal?.aborted) throw new DOMException('Pivot source acquisition cancelled', 'AbortError');
}

function yieldPivotAcquire(): Promise<void> {
  if (typeof MessageChannel === 'undefined') return new Promise((resolve) => setTimeout(resolve, 0));
  return new Promise((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => {
      channel.port1.close();
      channel.port2.close();
      resolve();
    };
    channel.port2.postMessage(undefined);
  });
}

interface LocalSourceNode {
  sourceId: string;
  range: RangeRef;
  table: SourceTable;
}

interface LocalRelationship {
  id: string;
  left: { sourceId: string; fieldId: string };
  right: { sourceId: string; fieldId: string };
  join: 'inner' | 'left';
}

function sourceField(table: SourceTable, fieldId: string, sourceId: string): SourceField {
  const field = table.fields.find((candidate) => candidate.fieldId === fieldId);
  if (!field) throw new Error(`Pivot relationship references unknown field ${sourceId}:${fieldId}`);
  return field;
}

function sourceFieldType(table: SourceTable, fieldId: string): PivotFieldDataType {
  return inferPivotSourceFieldType(sourceColumnValues(table, fieldId));
}

function joinKey(value: PivotScalar): string {
  return pivotMemberKey(createPivotMemberKey(value));
}

function assertUniqueLookupKeys(table: SourceTable, fieldId: string, sourceId: string): void {
  const keys = new Set<string>();
  for (const row of table.rows) {
    const key = joinKey(sourceRowValue(row, fieldId));
    if (keys.has(key)) throw new Error(`Pivot relationship lookup key is not unique: ${sourceId}:${fieldId}`);
    keys.add(key);
  }
}

function validateRelationshipGraph(nodes: LocalSourceNode[], relationships: readonly LocalRelationship[]): { edges: LocalRelationship[]; rootId: string } {
  const nodeIds = new Set(nodes.map((node) => node.sourceId));
  const nodeById = new Map(nodes.map((node) => [node.sourceId, node]));
  const relationshipIds = new Set<string>();
  const edges = relationships.map((relationship) => {
    if (!relationship.id || relationshipIds.has(relationship.id)) throw new Error(`Pivot relationship id is duplicated: ${relationship.id}`);
    relationshipIds.add(relationship.id);
    if (!nodeIds.has(relationship.left.sourceId) || !nodeIds.has(relationship.right.sourceId) || relationship.left.sourceId === relationship.right.sourceId) {
      throw new Error(`Pivot relationship references an unknown or self source node: ${relationship.id}`);
    }
    const leftNode = nodeById.get(relationship.left.sourceId)!;
    const rightNode = nodeById.get(relationship.right.sourceId)!;
    const leftField = sourceField(leftNode.table, relationship.left.fieldId, relationship.left.sourceId);
    const rightField = sourceField(rightNode.table, relationship.right.fieldId, relationship.right.sourceId);
    const leftType = sourceFieldType(leftNode.table, leftField.fieldId);
    const rightType = sourceFieldType(rightNode.table, rightField.fieldId);
    if (leftType === 'mixed' || rightType === 'mixed' || leftType !== rightType) {
      throw new Error(`Pivot relationship key types are incompatible: ${relationship.id}`);
    }
    assertUniqueLookupKeys(rightNode.table, rightField.fieldId, relationship.right.sourceId);
    if (relationship.join === 'inner') assertUniqueLookupKeys(leftNode.table, leftField.fieldId, relationship.left.sourceId);
    return structuredClone(relationship);
  });
  if (nodes.length > 1 && edges.length === 0) throw new Error('Pivot relationship graph is disconnected');
  const parent = new Map<string, string>(nodes.map((node) => [node.sourceId, node.sourceId]));
  const find = (sourceId: string): string => {
    const current = parent.get(sourceId);
    if (!current || current === sourceId) return sourceId;
    const root = find(current);
    parent.set(sourceId, root);
    return root;
  };
  for (const edge of edges) {
    const left = find(edge.left.sourceId);
    const right = find(edge.right.sourceId);
    if (left === right) throw new Error(`Pivot relationship graph contains a cycle: ${edge.id}`);
    parent.set(left, right);
  }
  const rootCandidates = edges.some((edge) => edge.join === 'left')
    ? nodes.filter((node) => !edges.some((edge) => edge.join === 'left' && edge.right.sourceId === node.sourceId))
    : [[...nodes].sort((left, right) => left.sourceId.localeCompare(right.sourceId))[0]!];
  if (rootCandidates.length !== 1) throw new Error('Pivot relationship graph has an ambiguous root');
  const rootId = rootCandidates[0]!.sourceId;
  const reachable = new Set<string>([rootId]);
  while (true) {
    const next = edges.flatMap((edge) => {
      if (reachable.has(edge.left.sourceId) && !reachable.has(edge.right.sourceId)) return [edge.right.sourceId];
      if (reachable.has(edge.right.sourceId) && !reachable.has(edge.left.sourceId)) return [edge.left.sourceId];
      return [];
    });
    if (!next.length) break;
    next.forEach((sourceId) => reachable.add(sourceId));
  }
  if (reachable.size !== nodes.length) throw new Error('Pivot relationship graph is disconnected');
  return { edges: edges.sort((left, right) => left.id.localeCompare(right.id)), rootId };
}

function joinSourceTables(current: SourceTable, attached: SourceTable, currentFieldId: string, attachedFieldId: string, join: 'inner' | 'left'): SourceTable {
  const lookup = new Map<string, SourceRow>();
  for (const row of attached.rows) lookup.set(joinKey(sourceRowValue(row, attachedFieldId)), row);
  const fields = [...current.fields, ...attached.fields].map((field, ordinal) => ({ ...field, ordinal }));
  const columnValues = fields.map(() => [] as PivotScalar[]);
  const rowPaths: PivotSourceRowPath[][] = [];
  for (const left of current.rows) {
    const match = lookup.get(joinKey(sourceRowValue(left, currentFieldId)));
    if (!match) {
      if (join === 'left') {
        current.fields.forEach((field, ordinal) => columnValues[ordinal]!.push(sourceRowValue(left, field.fieldId)));
        attached.fields.forEach((_field, ordinal) => columnValues[current.fields.length + ordinal]!.push(null));
        rowPaths.push([...sourceRowPaths(left)]);
      }
      continue;
    }
    current.fields.forEach((field, ordinal) => columnValues[ordinal]!.push(sourceRowValue(left, field.fieldId)));
    attached.fields.forEach((field, ordinal) => columnValues[current.fields.length + ordinal]!.push(sourceRowValue(match, field.fieldId)));
    const recordId = sourceRowPaths(left)[0]?.recordId ?? sourceRowPaths(match)[0]?.recordId;
    rowPaths.push([...sourceRowPaths(left), ...sourceRowPaths(match)].map((path) => ({ ...path, ...(recordId ? { recordId } : {}) })));
  }
  return openSourceTable(createPivotSourceIndex({
    columns: fields.map((field, ordinal) => ({ field, values: columnValues[ordinal]! })),
    rowPaths,
  }));
}

function remapTableSourceFields(table: SourceTable, stored: readonly { id: string; name: string }[]): SourceTable {
  const fields = table.fields.map((field, index) => {
    const declared = stored[index];
    return {
      ...field,
      ...(declared?.id ? { fieldId: declared.id } : {}),
      ...(declared?.name ? { name: declared.name } : {}),
    };
  });
  return openSourceTable({ ...table.index, fields });
}

function buildSourceTable(workbook: WorkbookModel, pivot: PivotModel, catalog?: PivotFieldCatalog, formula?: FormulaEngine): SourceTable {
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') {
    throw new Error(`Block-backed data source ${source.dataSourceId} requires asynchronous Pivot computation`);
  }
  const ranges = sourceRanges(workbook, pivot, formula);
  if (source.kind === 'worksheet-ranges') {
    const nodes = source.ranges.map((sourceRange, index) => ({
      sourceId: sourceRange.sourceId,
      range: sourceRange.range,
      table: readRange(workbook.getSheet(sourceRange.range.sheetId), sourceRange.range, source, index, catalog, formula),
    }));
    if (new Set(nodes.map((node) => node.sourceId)).size !== nodes.length || nodes.some((node) => !node.sourceId.trim())) {
      throw new Error('Every local worksheet range must have a unique stable sourceId');
    }
    const plan = validateRelationshipGraph(nodes, source.relationships);
    let current = nodes.find((node) => node.sourceId === plan.rootId)!.table;
    const visited = new Set<string>([plan.rootId]);
    while (visited.size < nodes.length) {
      const candidate = plan.edges.find((edge) => (visited.has(edge.left.sourceId) && !visited.has(edge.right.sourceId)) || (visited.has(edge.right.sourceId) && !visited.has(edge.left.sourceId)));
      if (!candidate) throw new Error('Pivot relationship graph cannot be planned from its root');
      if (visited.has(candidate.left.sourceId)) {
        const attached = nodes.find((node) => node.sourceId === candidate.right.sourceId)!;
        current = joinSourceTables(current, attached.table, candidate.left.fieldId, candidate.right.fieldId, candidate.join);
        visited.add(candidate.right.sourceId);
      } else {
        if (candidate.join === 'left') throw new Error(`Left relationship ${candidate.id} cannot be traversed from its lookup side`);
        const attached = nodes.find((node) => node.sourceId === candidate.left.sourceId)!;
        current = joinSourceTables(current, attached.table, candidate.right.fieldId, candidate.left.fieldId, 'inner');
        visited.add(candidate.left.sourceId);
      }
    }
    return current;
  }
  const range = ranges[0]!;
  // Table field IDs come from the table model, so read the source columns with
  // their physical identities first and remap once below.
  const table = readRange(workbook.getSheet(range.sheetId), range, source, 0, source.kind === 'table' ? undefined : catalog, formula);
  if (source.kind === 'table') {
    return remapTableSourceFields(table, resolvePivotTable(workbook, source.tableId).fields);
  }
  return table;
}

async function buildSourceTableAsync(
  workbook: WorkbookModel,
  pivot: PivotModel,
  catalog: PivotFieldCatalog | undefined,
  formula: FormulaEngine | undefined,
  options: PivotSourceAcquireOptions,
): Promise<SourceTable> {
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') throw new Error(`Block-backed data source ${source.dataSourceId} requires asynchronous content acquisition`);
  const ranges = sourceRanges(workbook, pivot, formula);
  if (source.kind === 'worksheet-ranges') {
    const nodes: LocalSourceNode[] = [];
    for (let index = 0; index < source.ranges.length; index += 1) {
      const sourceRange = source.ranges[index]!;
      nodes.push({
        sourceId: sourceRange.sourceId,
        range: sourceRange.range,
        table: await readRangeAsync(workbook.getSheet(sourceRange.range.sheetId), sourceRange.range, source, index, catalog, formula, options),
      });
    }
    if (new Set(nodes.map((node) => node.sourceId)).size !== nodes.length || nodes.some((node) => !node.sourceId.trim())) {
      throw new Error('Every local worksheet range must have a unique stable sourceId');
    }
    const plan = validateRelationshipGraph(nodes, source.relationships);
    let current = nodes.find((node) => node.sourceId === plan.rootId)!.table;
    const visited = new Set<string>([plan.rootId]);
    while (visited.size < nodes.length) {
      assertPivotAcquireActive(options.signal);
      const candidate = plan.edges.find((edge) => (visited.has(edge.left.sourceId) && !visited.has(edge.right.sourceId)) || (visited.has(edge.right.sourceId) && !visited.has(edge.left.sourceId)));
      if (!candidate) throw new Error('Pivot relationship graph cannot be planned from its root');
      if (visited.has(candidate.left.sourceId)) {
        const attached = nodes.find((node) => node.sourceId === candidate.right.sourceId)!;
        current = joinSourceTables(current, attached.table, candidate.left.fieldId, candidate.right.fieldId, candidate.join);
        visited.add(candidate.right.sourceId);
      } else {
        if (candidate.join === 'left') throw new Error(`Left relationship ${candidate.id} cannot be traversed from its lookup side`);
        const attached = nodes.find((node) => node.sourceId === candidate.left.sourceId)!;
        current = joinSourceTables(current, attached.table, candidate.right.fieldId, candidate.left.fieldId, 'inner');
        visited.add(candidate.left.sourceId);
      }
      await yieldPivotAcquire();
    }
    return current;
  }
  const range = ranges[0]!;
  const table = await readRangeAsync(workbook.getSheet(range.sheetId), range, source, 0, source.kind === 'table' ? undefined : catalog, formula, options);
  return source.kind === 'table' ? remapTableSourceFields(table, resolvePivotTable(workbook, source.tableId).fields) : table;
}

interface PivotSourceTableCacheEntry {
  revision: string;
  table: SourceTable;
}

const pivotSourceTableCaches = new WeakMap<WorkbookModel, Map<string, PivotSourceTableCacheEntry>>();
const pendingPivotSourceTables = new WeakMap<WorkbookModel, Map<string, { revision: string; promise: Promise<SourceTable> }>>();
const MAX_PIVOT_SOURCE_CACHE_ENTRIES = 8;

function sourceIndexAttached(table: SourceTable): boolean {
  return table.index.columns.every((column) => (column.kind === 'dictionary' ? column.codes.length : column.values.length) === table.index.rowCount);
}

function sourceTable(workbook: WorkbookModel, pivot: PivotModel, catalog?: PivotFieldCatalog, formula?: FormulaEngine): SourceTable {
  if (pivot.source.kind === 'data-source') return buildSourceTable(workbook, pivot, catalog, formula);
  const cacheKey = fingerprint(canonicalPivotSource(pivot.source));
  const revision = sourceRevision(workbook, pivot, formula);
  const cache = pivotSourceTableCaches.get(workbook) ?? new Map<string, PivotSourceTableCacheEntry>();
  const current = cache.get(cacheKey);
  if (current?.revision === revision && sourceIndexAttached(current.table)) return current.table;
  const table = buildSourceTable(workbook, pivot, catalog, formula);
  cache.delete(cacheKey);
  cache.set(cacheKey, { revision, table });
  while (cache.size > MAX_PIVOT_SOURCE_CACHE_ENTRIES) cache.delete(cache.keys().next().value!);
  if (!pivotSourceTableCaches.has(workbook)) pivotSourceTableCaches.set(workbook, cache);
  return table;
}

async function sourceTableAsync(
  workbook: WorkbookModel,
  pivot: PivotModel,
  catalog: PivotFieldCatalog | undefined,
  formula: FormulaEngine | undefined,
  options: PivotSourceAcquireOptions,
): Promise<SourceTable> {
  const cacheKey = fingerprint(canonicalPivotSource(pivot.source));
  const revision = sourceRevision(workbook, pivot, formula);
  const cache = pivotSourceTableCaches.get(workbook) ?? new Map<string, PivotSourceTableCacheEntry>();
  const current = cache.get(cacheKey);
  if (current?.revision === revision && sourceIndexAttached(current.table)) return current.table;
  const pending = pendingPivotSourceTables.get(workbook) ?? new Map<string, { revision: string; promise: Promise<SourceTable> }>();
  const existing = pending.get(cacheKey);
  if (existing?.revision === revision) return existing.promise;
  const promise = buildSourceTableAsync(workbook, pivot, catalog, formula, options).then((table) => {
    cache.delete(cacheKey);
    cache.set(cacheKey, { revision, table });
    while (cache.size > MAX_PIVOT_SOURCE_CACHE_ENTRIES) cache.delete(cache.keys().next().value!);
    if (!pivotSourceTableCaches.has(workbook)) pivotSourceTableCaches.set(workbook, cache);
    return table;
  }).finally(() => {
    if (pending.get(cacheKey)?.promise === promise) pending.delete(cacheKey);
  });
  pending.set(cacheKey, { revision, promise });
  if (!pendingPivotSourceTables.has(workbook)) pendingPivotSourceTables.set(workbook, pending);
  return promise;
}

function pivotSourceCalculator(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): FormulaEngine | undefined {
  if (formula) return formula;
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') return undefined;
  if (source.kind === 'named-range') return createPivotFormulaEngine(workbook);
  const ranges = sourceRanges(workbook, pivot);
  let requiresFormula = false;
  for (const range of ranges) {
    workbook.getSheet(range.sheetId).cells.forEachInRange(
      range.startRow,
      range.endRow,
      range.startColumn,
      range.endColumn,
      (cell) => {
        if (cell.formula !== undefined && !cell.formulaMetadata?.preservedOnly) requiresFormula = true;
      },
    );
  }
  return requiresFormula ? createPivotFormulaEngine(workbook) : undefined;
}

export function canonicalPivotMembers(values: readonly PivotScalar[]): PivotScalar[] {
  const members = [...new Map(values.map((value) => {
    // Empty text and null are one semantic blank member. Keep typed values
    // distinct so number 1 and text "1" remain independently filterable.
    const canonical = value === '' ? null : value;
    return [pivotMemberKey(createPivotMemberKey(canonical)), canonical] as const;
  })).values()];
  if (members.length > PIVOT_MAX_MEMBER_COUNT) {
    throw new Error(`Pivot field member domain exceeds ${PIVOT_MAX_MEMBER_COUNT} unique members`);
  }
  return members;
}

function normalizeFieldCatalog(sourceTableValue: SourceTable, persisted?: PivotFieldCatalog): PivotFieldCatalog {
  const fields = sourceTableValue.fields.map((field, ordinal) => {
    const values = sourceColumnValues(sourceTableValue, field.fieldId);
    const persistedField = persisted?.fields.find((candidate) => candidate.fieldId === field.fieldId);
    const fieldId = persistedField?.fieldId ?? field.fieldId ?? `field:${ordinal}`;
    const members = canonicalPivotMembers(values);
    return { fieldId, name: field.name, dataType: inferPivotSourceFieldType(values), ordinal, values: members };
  });
  return { schema: 'PivotFieldCatalog', fields };
}

interface PivotFieldCatalogCacheEntry {
  identity: string;
  catalog: PivotFieldCatalog;
}

const pivotFieldCatalogCache = new WeakMap<SourceTable, PivotFieldCatalogCacheEntry>();

function normalizedFieldCatalog(sourceTableValue: SourceTable, persisted?: PivotFieldCatalog): PivotFieldCatalog {
  const identity = fingerprint((persisted?.fields ?? []).map(({ fieldId, name, dataType, ordinal }) => ({ fieldId, name, dataType, ordinal })));
  const cached = pivotFieldCatalogCache.get(sourceTableValue);
  if (cached?.identity === identity) return structuredClone(cached.catalog);
  const catalog = normalizeFieldCatalog(sourceTableValue, persisted);
  pivotFieldCatalogCache.set(sourceTableValue, { identity, catalog: structuredClone(catalog) });
  return catalog;
}

function resolveFieldId(reference: string | undefined, catalog: PivotFieldCatalog): string | undefined {
  if (!reference) return undefined;
  return catalog.fields.find((field) => field.fieldId === reference || field.name === reference)?.fieldId;
}

function fieldName(fieldId: string, catalog: PivotFieldCatalog): string {
  return catalog.fields.find((field) => field.fieldId === fieldId)?.name ?? fieldId;
}

/**
 * Build the canonical, locale-independent summary for one report field.
 * Report filters are grouped by field so allowing multiple filter families
 * cannot silently select one family and render the others as `All`.
 */
export function summarizePivotReportFilters(
  filters: readonly PivotFilter[],
  catalog: PivotFieldCatalog,
  fieldId: string,
  values: readonly PivotValueField[] = [],
): PivotReportFilterSummary {
  const valueName = (valueId: string): string => {
    const value = values.find((entry) => entry.valueId === valueId);
    return fieldName(value?.fieldId ?? valueId, catalog);
  };
  const fieldFilters = filters.filter((filter) => filter.fieldId === fieldId && (filter.scope ?? 'report') !== 'field');
  const entries: PivotReportFilterSummaryEntry[] = fieldFilters.map((filter) => {
    if (filter.kind === 'manual') {
      const memberValues = filter.memberKeys.map((member) => pivotScalarFromMemberKey(member));
      return {
        kind: 'manual',
        family: 'manual',
        active: filter.mode === 'include' || memberValues.length > 0,
        mode: filter.mode,
        count: memberValues.length,
        memberValues,
      };
    }
    if (filter.kind === 'top-items') {
      return {
        kind: 'top-items',
        family: 'top-items',
        active: true,
        mode: filter.mode,
        threshold: filter.threshold,
        direction: filter.direction,
        valueFieldName: valueName(filter.valueId),
      };
    }
    return {
      kind: 'condition',
      family: filter.family,
      active: true,
      operator: filter.operator,
      value: filter.value,
      ...(filter.value2 === undefined ? {} : { value2: filter.value2 }),
      ...(filter.dynamic === undefined ? {} : { dynamic: filter.dynamic }),
      ...(filter.valueId === undefined ? {} : { valueFieldName: valueName(filter.valueId) }),
    } as PivotReportFilterSummaryEntry;
  });
  return {
    fieldName: fieldName(fieldId, catalog),
    active: entries.some((entry) => entry.active),
    entries,
  };
}

function normalizePlacement(placement: PivotFieldPlacement, catalog: PivotFieldCatalog, valueIds: ReadonlySet<string>): PivotFieldPlacement {
  const fieldId = resolveFieldId(placement.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot field: ${placement.fieldId}`);
  let sort: PivotSort | undefined;
  if (placement.sort) {
    if (placement.sort.by === 'value') {
      const valueId = placement.sort.valueId;
      if (!valueId) throw new Error(`Pivot value sort requires a valueId for ${fieldId}`);
      if (!valueIds.has(valueId)) throw new Error(`Pivot value sort placement is not in Values: ${valueId}`);
      sort = { direction: placement.sort.direction, by: 'value', valueId };
    } else if (placement.sort.by === 'label') {
      if (Object.prototype.hasOwnProperty.call(placement.sort, 'valueId')) throw new Error(`Pivot label sort cannot carry a Values placement identity for ${fieldId}`);
      sort = { direction: placement.sort.direction, by: 'label' };
    } else {
      throw new Error(`Pivot sort mode is invalid for ${fieldId}`);
    }
  }
  return { fieldId, sort, group: placement.group, subtotal: placement.subtotal ? structuredClone(placement.subtotal) : undefined };
}

function validateTopBottomThreshold(mode: PivotTopBottomMode, threshold: number): void {
  if (!['items', 'percent', 'sum'].includes(mode) || !Number.isFinite(threshold) || threshold <= 0
    || (mode === 'items' && (!Number.isSafeInteger(threshold) || threshold < 1))
    || (mode === 'percent' && threshold > 100)) {
    throw new Error('Pivot top-items threshold is invalid');
  }
}

function validateTopBottomDirection(direction: 'top' | 'bottom'): void {
  if (direction !== 'top' && direction !== 'bottom') throw new Error('Pivot top-items direction is invalid');
}

function normalizeFilter(filter: PivotFilter, catalog: PivotFieldCatalog, valueIds: ReadonlySet<string>): PivotFilter {
  const fieldId = resolveFieldId(filter.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot field: ${filter.fieldId}`);
  if (filter.kind === 'manual') {
    if (filter.family !== 'manual') throw new Error(`Pivot manual filter family is invalid: ${fieldId}`);
    return { kind: 'manual', family: 'manual', fieldId, scope: filter.scope, mode: filter.mode, memberKeys: structuredClone(filter.memberKeys) };
  }
  if (filter.kind === 'top-items') {
    if (filter.family !== 'top-items') throw new Error(`Pivot top-items filter family is invalid: ${fieldId}`);
    if (!valueIds.has(filter.valueId)) throw new Error(`Unknown Pivot Values placement: ${filter.valueId}`);
    if (Object.prototype.hasOwnProperty.call(filter, 'count')) throw new Error('Pivot top-items count is no longer supported; use threshold');
    validateTopBottomDirection(filter.direction);
    validateTopBottomThreshold(filter.mode, filter.threshold);
    return { ...filter, fieldId, valueId: filter.valueId };
  }
  if (!['label', 'date', 'value'].includes(filter.family)) throw new Error(`Pivot condition filter family is invalid: ${fieldId}`);
  if (filter.family === 'value') {
    if (!filter.valueId || !valueIds.has(filter.valueId)) throw new Error(`Unknown Pivot Values placement: ${filter.valueId ?? '(missing)'}`);
    return { ...filter, fieldId, valueId: filter.valueId };
  }
  if (filter.valueId !== undefined) throw new Error(`Pivot condition valueId is only valid for value filters: ${fieldId}`);
  return { ...filter, fieldId };
}

function normalizeValueField(field: PivotValueField, catalog: PivotFieldCatalog): PivotValueField {
  if (!field.valueId) throw new Error('Pivot Values placement identity is required');
  const fieldId = resolveFieldId(field.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot value field: ${field.fieldId}`);
  const numberFormat = normalizePivotNumberFormat(field.numberFormat);
  if (Object.prototype.hasOwnProperty.call(field as object, 'baseFieldId') || Object.prototype.hasOwnProperty.call(field as object, 'baseItem')) {
    throw new Error('Pivot value field baseFieldId/baseItem are no longer accepted; configure them inside showAs');
  }
  return { ...field, fieldId, ...(numberFormat === undefined ? {} : { numberFormat }) };
}

function normalizeShowAsBaseItem(value: unknown): PivotShowAsBaseItem {
  if (value === 'previous' || value === 'next') return value;
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('Pivot showAs baseItem is invalid');
  const member = value as Record<string, unknown>;
  if (!['text', 'number', 'boolean', 'blank', 'error'].includes(String(member.type))) throw new Error('Pivot showAs baseItem type is invalid');
  if (member.type === 'blank' && member.value !== null) throw new Error('Pivot showAs blank baseItem must have null value');
  if (member.type === 'text' && typeof member.value !== 'string') throw new Error('Pivot showAs text baseItem is invalid');
  if (member.type === 'number' && (typeof member.value !== 'number' || !Number.isFinite(member.value))) throw new Error('Pivot showAs number baseItem is invalid');
  if (member.type === 'boolean' && typeof member.value !== 'boolean') throw new Error('Pivot showAs boolean baseItem is invalid');
  if (member.type === 'error' && typeof member.value !== 'string') throw new Error('Pivot showAs error baseItem is invalid');
  return { type: member.type as 'text' | 'number' | 'boolean' | 'blank' | 'error', value: member.value as string | number | boolean | null };
}

function normalizeValueShowAs(field: PivotValueField, catalog: PivotFieldCatalog, axisFieldIds: ReadonlySet<string>): PivotValueField {
  const raw = field.showAs;
  if (raw === undefined) return field;
  const showAs = raw as unknown as Record<string, unknown>;
  const kind = showAs.kind;
  const totalKinds = new Set(['normal', 'grand-percentage', 'row-percentage', 'column-percentage', 'parent-percentage', 'index']);
  if (typeof kind !== 'string') throw new Error('Pivot showAs kind is invalid');
  if (totalKinds.has(kind)) {
    if (Object.keys(showAs).some((key) => key !== 'kind')) throw new Error('Pivot showAs contains unknown fields');
    return { ...field, showAs: { kind } as PivotValueField['showAs'] };
  }
  const baseFieldId = showAs.baseFieldId;
  if (typeof baseFieldId !== 'string' || !baseFieldId.trim()) throw new Error(`Pivot ${kind} showAs requires baseFieldId`);
  const resolvedBaseFieldId = resolveFieldId(baseFieldId, catalog);
  if (!resolvedBaseFieldId || !axisFieldIds.has(resolvedBaseFieldId)) throw new Error(`Pivot ${kind} showAs baseFieldId must target a row or column field: ${baseFieldId}`);
  if (kind === 'difference' || kind === 'percentage-difference') {
    if (!Object.prototype.hasOwnProperty.call(showAs, 'baseItem')) throw new Error(`Pivot ${kind} showAs requires baseItem`);
    return { ...field, showAs: { kind, baseFieldId: resolvedBaseFieldId, baseItem: normalizeShowAsBaseItem(showAs.baseItem) } as PivotValueField['showAs'] };
  }
  if (kind === 'running-total' || kind === 'percentage-running-total') {
    if (Object.keys(showAs).some((key) => key !== 'kind' && key !== 'baseFieldId')) throw new Error('Pivot running-total showAs contains unknown fields');
    return { ...field, showAs: { kind, baseFieldId: resolvedBaseFieldId } as PivotValueField['showAs'] };
  }
  if (kind === 'rank') {
    if (!['ascending', 'descending'].includes(String(showAs.direction)) || Object.keys(showAs).some((key) => !['kind', 'baseFieldId', 'direction'].includes(key))) throw new Error('Pivot rank showAs is invalid');
    return { ...field, showAs: { kind, baseFieldId: resolvedBaseFieldId, direction: showAs.direction as 'ascending' | 'descending' } };
  }
  throw new Error(`Pivot showAs kind is unsupported: ${String(kind)}`);
}

function normalizeLayout(layout: PivotLayout, catalog: PivotFieldCatalog): PivotLayout {
  if (!['compact', 'outline', 'tabular'].includes(layout.reportLayout)) throw new Error('Pivot report layout is invalid');
  const rawValues = layout.values.map((entry) => normalizeValueField(entry, catalog));
  const valueIds = new Set<string>();
  for (const entry of rawValues) {
    if (valueIds.has(entry.valueId)) throw new Error(`Duplicate Pivot Values placement identity: ${entry.valueId}`);
    valueIds.add(entry.valueId);
  }
  const normalizedRows = layout.rows.map((entry) => normalizePlacement(entry, catalog, valueIds));
  const normalizedColumns = layout.columns.map((entry) => normalizePlacement(entry, catalog, valueIds));
  const axisFieldIds = new Set([...normalizedRows, ...normalizedColumns].map((entry) => entry.fieldId));
  const values = rawValues.map((entry) => normalizeValueShowAs(entry, catalog, axisFieldIds));
  const filters = layout.filters.map((entry) => normalizeFilter(entry, catalog, valueIds));
  const scopedFilters = filters.map((filter) => {
    const scope = filter.scope ?? (axisFieldIds.has(filter.fieldId) ? 'field' : 'report');
    if (scope === 'field' && !axisFieldIds.has(filter.fieldId)) {
      throw new Error(`Pivot field filter must target a row or column field: ${filter.fieldId}`);
    }
    return { ...filter, scope };
  });
  for (const filter of scopedFilters) {
    if (filter.kind !== 'manual' || (filter.scope ?? 'report') !== 'field') continue;
    const placement = [...normalizedRows, ...normalizedColumns].find((entry) => entry.fieldId === filter.fieldId && entry.group);
    const field = catalog.fields.find((entry) => entry.fieldId === filter.fieldId);
    if (!placement?.group || !field?.values?.length) continue;
    const validKeys = new Set(buildPivotGroupedFilterMembers(field.values, placement.group).map((member) => pivotMemberKey(member.key)));
    const invalid = filter.memberKeys.find((member) => !validKeys.has(pivotMemberKey(member)));
    if (invalid) throw new Error(`Pivot grouped filter member is incompatible with grouping for ${filter.fieldId}`);
  }
  const identities = new Set<string>();
  const fields = new Set<string>();
  for (const filter of scopedFilters) {
    const identity = `${filter.fieldId}|${filter.scope ?? 'report'}|${filter.family}`;
    if (identities.has(identity)) throw new Error(`Duplicate Pivot filter family: ${identity}`);
    identities.add(identity);
    const fieldScope = `${filter.fieldId}|${filter.scope ?? 'report'}`;
    if (!layout.allowMultipleFiltersPerField && fields.has(fieldScope)) throw new Error(`Multiple Pivot filters are disabled for ${fieldScope}`);
    fields.add(fieldScope);
  }
  return {
    ...structuredClone(layout),
    rows: normalizedRows,
    columns: normalizedColumns,
    filters: scopedFilters,
    values,
    expansion: layout.expansion ? {
      expandedNodeIds: [...layout.expansion.expandedNodeIds],
      collapsedNodeIds: [...layout.expansion.collapsedNodeIds],
      showButtons: layout.expansion.showButtons,
    } : {
      expandedNodeIds: [],
      collapsedNodeIds: [],
      showButtons: true,
    },
  };
}

/**
 * Validate calculated-item definitions without widening the source field
 * catalogue. A calculated item is a derived member identity owned by its
 * target field; it is never a field that can be placed on an axis or in
 * Values. This runtime check mirrors the protocol/backend effective-field
 * boundary so direct local projection cannot accept a shape the server will
 * reject.
 */
function normalizeCalculatedItemDefinitions(
  calculatedItems: PivotLayout['calculatedItems'] = [],
  catalog: PivotFieldCatalog,
  calculatedFieldIds: ReadonlySet<string> = new Set(),
): void {
  const fieldIds = new Set(catalog.fields.map((field) => field.fieldId));
  const itemIds = new Set<string>();
  const itemNames = new Set<string>();
  for (const item of calculatedItems) {
    if (!item || typeof item !== 'object' || typeof item.fieldId !== 'string' || !item.fieldId.trim()
      || typeof item.targetFieldId !== 'string' || !item.targetFieldId.trim()
      || typeof item.name !== 'string' || !item.name.trim()
      || typeof item.formula !== 'string' || !item.formula.trim()) {
      throw new Error('Pivot calculated item definition is invalid');
    }
    if (fieldIds.has(item.fieldId) || calculatedFieldIds.has(item.fieldId) || !itemIds.add(item.fieldId)) {
      throw new Error(`Pivot calculated item identity collides with an effective field or another item: ${item.fieldId}`);
    }
    if (calculatedFieldIds.has(item.targetFieldId)) {
      throw new Error(`Pivot calculated item target field cannot be a calculated field: ${item.targetFieldId}`);
    }
    const targetField = catalog.fields.find((field) => field.fieldId === item.targetFieldId);
    if (!targetField) throw new Error(`Pivot calculated item target field is unknown: ${item.targetFieldId}`);
    const nameKey = `${targetField.fieldId}|${item.name}`;
    if (!itemNames.add(nameKey)) throw new Error(`Pivot calculated item member is duplicated: ${item.name}`);
    if ((targetField.values ?? []).some((value) => same(value, item.name))) {
      throw new Error(`Pivot calculated item member already exists in source data: ${item.name}`);
    }
  }
}

function appendCalculatedItemMembers(catalog: PivotFieldCatalog, calculatedItems: readonly PivotCalculatedItem[] = []): void {
  for (const item of calculatedItems) {
    const target = catalog.fields.find((field) => field.fieldId === item.targetFieldId);
    if (!target) throw new Error(`Pivot calculated item target field is unknown: ${item.targetFieldId}`);
    const values = target.values ?? [];
    if (!values.some((value) => same(value, item.name))) target.values = [...values, item.name];
  }
}

function normalizePivotDefinitionWithCalculator(workbook: WorkbookModel, pivot: PivotModel, calculator: FormulaEngine | undefined): PivotDefinition {
  const source = getPivotSource(pivot);
  const fieldCatalog = source.kind === 'data-source'
    ? getPivotFieldCatalog(workbook, pivot)
    : normalizedFieldCatalog(sourceTable(workbook, pivot, pivot.fieldCatalog, calculator), pivot.fieldCatalog);
  return normalizePivotDefinitionFromCatalog({ ...pivot, source, fieldCatalog });
}

/** Canonicalize a command/task definition from its already validated revision-owned field catalog. */
export function normalizePivotDefinitionFromCatalog(pivot: PivotModel): PivotDefinition {
  const source = getPivotSource(pivot);
  const fieldCatalog = structuredClone(pivot.fieldCatalog);
  const calculatedFields = (pivot.layout.calculatedFields ?? []).map((field) => ({ fieldId: field.fieldId, name: field.name }));
  for (const calculated of calculatedFields) {
    if (!fieldCatalog.fields.some((field) => field.fieldId === calculated.fieldId || field.name === calculated.name)) {
      fieldCatalog.fields.push({ fieldId: calculated.fieldId, name: calculated.name, dataType: 'mixed', ordinal: fieldCatalog.fields.length, values: [] });
    }
  }
  normalizeCalculatedItemDefinitions(pivot.layout.calculatedItems, fieldCatalog, new Set(calculatedFields.map((field) => field.fieldId)));
  appendCalculatedItemMembers(fieldCatalog, pivot.layout.calculatedItems);
  const layout = normalizeLayout(pivot.layout, fieldCatalog);
  return {
    schema: 'PivotDefinition',
    id: pivot.id,
    source,
    target: getPivotTarget(pivot),
    fieldCatalog,
    layout,
    refreshPolicy: normalizePivotRefreshPolicy(pivot.refreshPolicy),
    presentation: {
      ...(pivot.presentation?.styleName ? { styleName: pivot.presentation.styleName } : {}),
      styleOptions: { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(pivot.presentation?.styleOptions ?? {}) },
      displayOptions: normalizePivotDisplayOptions(pivot.presentation?.displayOptions),
    },
    ...(pivot.nativeMetadata ? { nativeMetadata: structuredClone(pivot.nativeMetadata) } : {}),
  };
}

/** Canonicalize field catalog values against the live source. Calculation has one model shape. */
export function normalizePivotDefinition(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotDefinition {
  return normalizePivotDefinitionWithCalculator(workbook, pivot, pivotSourceCalculator(workbook, pivot, formula));
}

export function getPivotFieldCatalog(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotFieldCatalog {
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') {
    const manifest = workbook.getDataSource(source.dataSourceId);
    return {
      schema: 'PivotFieldCatalog',
      fields: manifest.fields.map((field) => ({
        fieldId: field.id,
        name: field.name,
        dataType: field.type,
        ordinal: field.ordinal,
        values: [],
      })),
    };
  }
  const calculator = pivotSourceCalculator(workbook, pivot, formula);
  return normalizedFieldCatalog(sourceTable(workbook, { ...pivot, source }, pivot.fieldCatalog, calculator), pivot.fieldCatalog);
}

function formulaScalar(value: FormulaValue): PivotScalar | null {
  if (isFormulaError(value)) return { kind: 'error', code: value.code, ...(value.message ? { message: value.message } : {}) };
  if (Array.isArray(value)) throw new Error('Pivot calculated formula returned an array instead of a scalar');
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || value === null ? value : null;
}

function columnLabel(index: number): string {
  let value = '';
  let current = index + 1;
  while (current > 0) {
    const remainder = (current - 1) % 26;
    value = String.fromCharCode(65 + remainder) + value;
    current = Math.floor((current - 1) / 26);
  }
  return value;
}

const formulaFunctions = new Set(['SUM', 'COUNT', 'AVERAGE', 'MIN', 'MAX', 'IF', 'AND', 'OR', 'NOT', 'ROUND', 'ABS', 'CONCAT', 'LEFT', 'RIGHT', 'LEN']);

type PivotCalculatedField = NonNullable<PivotLayout['calculatedFields']>[number];

interface CalculatedFieldPlan {
  fields: SourceField[];
  definitions: Map<string, PivotCalculatedField>;
  ordered: PivotCalculatedField[];
}

function rewriteCalculatedFormula(formula: string, fields: SourceField[]): string {
  let rewritten = formula.trim().replace(/^=/, '');
  fields.flatMap((field, index) => [
    { field: field.name, index },
    { field: field.fieldId, index },
  ]).filter(({ field }) => field.length > 0).sort((left, right) => right.field.length - left.field.length).forEach(({ field, index }) => {
    const escaped = field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const reference = `${columnLabel(index)}1`;
    rewritten = rewritten.replace(new RegExp(`\\[${escaped}\\]`, 'g'), reference);
    if (!formulaFunctions.has(field.toUpperCase())) rewritten = rewritten.replace(new RegExp(`(?<![A-Za-z0-9_])${escaped}(?![A-Za-z0-9_])`, 'g'), reference);
  });
  return `=${rewritten}`;
}

interface CalculatedItemReference {
  fieldId: string;
  member: PivotScalar;
  itemId?: string;
}

interface CalculatedItemPlanEntry extends PivotCalculatedItem {
  references: CalculatedItemReference[];
  rewrittenFormula: string;
}

interface CalculatedItemPlan {
  ordered: CalculatedItemPlanEntry[];
}

function itemMemberMatches(value: PivotScalar, text: string): boolean {
  if (value === null || isPivotError(value)) return false;
  return String(value).toLocaleUpperCase() === text.toLocaleUpperCase();
}

function itemFieldLabel(field: PivotFieldDefinition): string {
  return field.name || field.fieldId;
}

function calculatedItemMemberKey(fieldId: string, value: PivotScalar): string {
  return `${fieldId}|${pivotMemberKey(createPivotMemberKey(value))}`;
}

/**
 * Resolve the deliberately small, canonical Pivot-item formula grammar.
 *
 * Item formulas are not worksheet formulas: bare names identify a member and
 * `Field[Member]` is the qualified form.  Resolving names before handing the
 * expression to FormulaEngine keeps worksheet cell references out of this
 * contract and makes ambiguous cross-field names fail closed.
 */
function resolveCalculatedItemReferences(
  formula: string,
  fields: readonly PivotFieldDefinition[],
  calculatedItems: readonly PivotCalculatedItem[],
): { references: CalculatedItemReference[]; rewritten: string } {
  const source = formula.trim().replace(/^=/, '').trim();
  const tokens = parsePivotCalculatedItemFormula(formula);
  const membersByField = new Map<string, PivotScalar[]>();
  for (const field of fields) membersByField.set(field.fieldId, [...(field.values ?? [])]);
  for (const item of calculatedItems) {
    const members = membersByField.get(item.targetFieldId);
    if (!members) throw new Error(`Pivot calculated item target field is unknown: ${item.targetFieldId}`);
    if (!members.some((value) => same(value, item.name))) members.push(item.name);
  }

  const fieldByReference = (reference: string): PivotFieldDefinition | undefined => fields.find((field) => field.fieldId === reference || itemFieldLabel(field) === reference);
  const resolveMember = (field: PivotFieldDefinition, text: string): PivotScalar => {
    const candidates = (membersByField.get(field.fieldId) ?? []).filter((value) => itemMemberMatches(value, text));
    if (candidates.length !== 1) {
      if (!candidates.length) throw new Error(`Pivot calculated item references unknown item ${field.name}[${text}]`);
      throw new Error(`Pivot calculated item reference is ambiguous: ${field.name}[${text}]`);
    }
    return candidates[0]!;
  };
  const itemByMember = new Map<string, PivotCalculatedItem>();
  for (const item of calculatedItems) itemByMember.set(calculatedItemMemberKey(item.targetFieldId, item.name), item);
  const references: CalculatedItemReference[] = [];
  const replacements: Array<{ start: number; end: number; cell: string }> = [];
  const referenceCells = new Map<string, string>();
  const cellFor = (reference: CalculatedItemReference): string => {
    const key = calculatedItemMemberKey(reference.fieldId, reference.member);
    const existing = referenceCells.get(key);
    if (existing) return existing;
    const cell = `${columnLabel(referenceCells.size)}1`;
    referenceCells.set(key, cell);
    references.push(reference);
    return cell;
  };
  const occupied: Array<{ start: number; end: number }> = [];
  const addReplacement = (start: number, end: number, reference: CalculatedItemReference): void => {
    if (occupied.some((range) => start < range.end && end > range.start)) return;
    occupied.push({ start, end });
    replacements.push({ start, end, cell: cellFor(reference) });
  };
  const referenceFor = (field: PivotFieldDefinition, member: PivotScalar): CalculatedItemReference => ({
    fieldId: field.fieldId,
    member,
    ...(itemByMember.get(calculatedItemMemberKey(field.fieldId, member))?.fieldId
      ? { itemId: itemByMember.get(calculatedItemMemberKey(field.fieldId, member))!.fieldId }
      : {}),
  });
  const resolveToken = (token: Extract<PivotCalculatedItemFormulaToken, { kind: 'item' }>): void => {
    if (token.fieldReference !== undefined) {
      const field = fieldByReference(token.fieldReference);
      if (!field) throw new Error(`Pivot calculated item references unknown field: ${token.fieldReference}`);
      addReplacement(token.start, token.end, referenceFor(field, resolveMember(field, token.member)));
      return;
    }
    // Numeric literals are represented by number tokens, so an unqualified
    // token can only be a textual item.  It must resolve to exactly one field;
    // the target field is not an implicit disambiguation rule.
    if (/^[A-Z]{1,3}[1-9][0-9]*$/i.test(token.member)) {
      throw new Error(`Pivot calculated item worksheet reference is unsupported: ${token.member}`);
    }
    const candidates = fields.flatMap((field) => (membersByField.get(field.fieldId) ?? [])
      .filter((value) => typeof value === 'string' && itemMemberMatches(value, token.member))
      .map((member) => ({ field, member })));
    if (!candidates.length) throw new Error(`Pivot calculated item references unknown item: ${token.member}`);
    if (candidates.length !== 1) throw new Error(`Pivot calculated item reference is ambiguous: ${token.member}`);
    addReplacement(token.start, token.end, referenceFor(candidates[0]!.field, candidates[0]!.member));
  };
  for (const token of tokens) if (token.kind === 'item') resolveToken(token);
  const rewritten = replacements.sort((left, right) => right.start - left.start).reduce((current, replacement) => `${current.slice(0, replacement.start)}${replacement.cell}${current.slice(replacement.end)}`, source);
  try {
    parseFormula(`=${rewritten}`);
  } catch (error) {
    throw new Error(`Pivot calculated item formula is invalid: ${formula}`, { cause: error });
  }
  return { references, rewritten: `=${rewritten}` };
}

function createCalculatedItemPlan(
  fields: readonly PivotFieldDefinition[],
  calculatedItems: readonly PivotCalculatedItem[] = [],
): CalculatedItemPlan {
  const entries = new Map<string, CalculatedItemPlanEntry>();
  for (const item of calculatedItems) {
    if (entries.has(item.fieldId)) throw new Error(`Pivot calculated item is duplicated: ${item.fieldId}`);
    if (!item.name.trim() || !item.formula.trim()) throw new Error(`Pivot calculated item definition is invalid: ${item.fieldId}`);
    if (!fields.some((field) => field.fieldId === item.targetFieldId)) throw new Error(`Pivot calculated item target field is unknown: ${item.targetFieldId}`);
    entries.set(item.fieldId, { ...item, references: [], rewrittenFormula: item.formula });
  }
  const orderedEntries = [...entries.values()];
  for (const entry of orderedEntries) {
    const resolved = resolveCalculatedItemReferences(entry.formula, fields, calculatedItems);
    const references = resolved.references.map((reference) => ({ ...reference, ...(entries.has(reference.itemId ?? '') ? { itemId: reference.itemId } : {}) }));
    entry.references = references;
    entry.rewrittenFormula = resolved.rewritten;
  }
  const state = new Map<string, 'visiting' | 'visited'>();
  const ordered: CalculatedItemPlanEntry[] = [];
  const visit = (fieldId: string, path: string[]): void => {
    const current = state.get(fieldId);
    if (current === 'visited') return;
    if (current === 'visiting') throw new Error(`Pivot calculated item dependency cycle: ${[...path, fieldId].join(' -> ')}`);
    const entry = entries.get(fieldId);
    if (!entry) return;
    state.set(fieldId, 'visiting');
    for (const reference of entry.references) if (reference.itemId) visit(reference.itemId, [...path, fieldId]);
    state.set(fieldId, 'visited');
    ordered.push(entry);
  };
  for (const entry of orderedEntries) visit(entry.fieldId, []);
  return { ordered };
}

function calculatedItemContextKey(row: SourceRow, contextFieldIds: readonly string[], placements: readonly PivotFieldPlacement[]): string {
  return JSON.stringify(contextFieldIds.map((fieldId) => {
    const placement = placements.find((candidate) => candidate.fieldId === fieldId);
    return createPivotMemberKey(grouped(sourceRowValue(row, fieldId), placement?.group));
  }));
}

function evaluateCalculatedItemFormula(
  entry: CalculatedItemPlanEntry,
  rows: readonly SourceRow[],
  valueFieldId: string,
): PivotScalar {
  const engine = new FormulaEngine({ defaultSheetId: 'pivot-calculated-item' });
  entry.references.forEach((reference, index) => {
    const memberRows = rows.filter((row) => same(sourceRowValue(row, reference.fieldId), reference.member));
    const value = aggregateSourceRows(memberRows, valueFieldId, 'sum');
    if (isPivotError(value)) throw new Error(`Pivot calculated item source aggregate failed: ${entry.fieldId} (${value.code})`);
    engine.setValue({ sheetId: 'pivot-calculated-item', row: 0, column: index }, value);
  });
  try {
    engine.setFormula({ sheetId: 'pivot-calculated-item', row: 1, column: 0 }, entry.rewrittenFormula);
  } catch (error) {
    throw new Error(`Pivot calculated item formula evaluation failed: ${entry.fieldId}`, { cause: error });
  }
  const value = formulaScalar(engine.getCellValue({ sheetId: 'pivot-calculated-item', row: 1, column: 0 }));
  if (value === null && entry.references.length > 0 && rows.length === 0) return null;
  return value;
}

function applyCalculatedItems(
  rows: SourceRow[],
  fields: PivotFieldDefinition[],
  layout: PivotLayout,
  valueFields: readonly PivotValueField[],
): SourceRow[] {
  const calculatedItems = layout.calculatedItems ?? [];
  if (!calculatedItems.length) return rows;
  const plan = createCalculatedItemPlan(fields, calculatedItems);
  let currentRows = rows;
  const axisPlacements = [...layout.rows, ...layout.columns];
  for (const entry of plan.ordered) {
    const contextFieldIds = axisPlacements
      .map((placement) => placement.fieldId)
      // A calculated item is a new member of its target field.  Every other
      // axis field remains part of the summary context, even when the formula
      // explicitly references one of its members; otherwise a cross-field
      // qualified reference would silently aggregate across that axis.
      .filter((fieldId, index, all) => fieldId !== entry.targetFieldId && all.indexOf(fieldId) === index);
    const contexts = new Map<string, SourceRow[]>();
    for (const row of rows) {
      const key = calculatedItemContextKey(row, contextFieldIds, axisPlacements);
      const context = contexts.get(key) ?? [];
      context.push(row);
      contexts.set(key, context);
    }
    const generated = [...contexts.values()].map((context) => {
      const template = context[0];
      if (!template) throw new Error(`Pivot calculated item context is empty: ${entry.fieldId}`);
      const contextKey = calculatedItemContextKey(template, contextFieldIds, axisPlacements);
      const candidateRows = currentRows.filter((row) => calculatedItemContextKey(row, contextFieldIds, axisPlacements) === contextKey);
      const overrides = new Map(template.overrides ?? []);
      overrides.set(entry.targetFieldId, entry.name);
      for (const valueField of valueFields) {
        if (valueField.fieldId === entry.targetFieldId) continue;
        overrides.set(valueField.fieldId, evaluateCalculatedItemFormula(entry, candidateRows, valueField.fieldId));
      }
      return {
        source: template.source,
        row: template.row,
        overrides,
        pathsOverride: [...new Map(context.flatMap((row) => sourceRowPaths(row)).map((path) => [stableSerialize(path), path])).values()],
      };
    });
    currentRows = [...currentRows, ...generated];
  }
  return currentRows;
}

function calculatedFieldReferenceIds(formula: string, fields: SourceField[], definitions: Map<string, PivotCalculatedField>, ownerId: string): string[] {
  let ast;
  try {
    ast = parseFormula(rewriteCalculatedFormula(formula, fields));
  } catch (error) {
    const bracketReference = formula.match(/\[([^\]]+)\]/)?.[1];
    if (bracketReference && !fields.some((field) => field.name.toUpperCase() === bracketReference.toUpperCase() || field.fieldId.toUpperCase() === bracketReference.toUpperCase())) {
      throw new Error(`Pivot calculated field references unknown field: ${bracketReference}`, { cause: error });
    }
    throw new Error(`Pivot calculated field formula is invalid: ${ownerId}`, { cause: error });
  }
  const fieldReferences = new Map<string, string>();
  for (const field of fields) {
    for (const fieldName of [field.fieldId, field.name]) {
      const key = fieldName.toUpperCase();
      const previous = fieldReferences.get(key);
      if (previous && previous !== field.fieldId) throw new Error(`Pivot calculated field reference is ambiguous: ${fieldName}`);
      fieldReferences.set(key, field.fieldId);
    }
  }
  const references: string[] = [];
  for (const field of fields) {
    for (const fieldName of [field.fieldId, field.name]) {
      const escaped = fieldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const referenced = new RegExp(`\\[${escaped}\\]`, 'i').test(formula)
        || (!formulaFunctions.has(fieldName.toUpperCase()) && new RegExp(`(?<![A-Za-z0-9_])${escaped}(?![A-Za-z0-9_])`, 'i').test(formula));
      if (referenced && definitions.has(field.fieldId) && !references.includes(field.fieldId)) references.push(field.fieldId);
    }
  }
  for (const name of collectNameReferences(ast)) {
    const fieldId = fieldReferences.get(name.toUpperCase());
    if (!fieldId) throw new Error(`Pivot calculated field references unknown field: ${name}`);
  }
  return references;
}

function createCalculatedFieldPlan(fields: PivotFieldDefinition[], calculatedFields: PivotLayout['calculatedFields'] = []): CalculatedFieldPlan {
  const descriptors: SourceField[] = fields
    .map((field) => ({ fieldId: field.fieldId, name: field.name, ordinal: field.ordinal, dataType: field.dataType }));
  const definitions = new Map<string, PivotCalculatedField>();
  for (const calculated of calculatedFields ?? []) {
    if (definitions.has(calculated.fieldId)) throw new Error(`Pivot calculated field is duplicated: ${calculated.fieldId}`);
    if (!calculated.name.trim() || !calculated.formula.trim()) throw new Error(`Pivot calculated field definition is invalid: ${calculated.fieldId}`);
    definitions.set(calculated.fieldId, calculated);
    if (!descriptors.some((field) => field.fieldId === calculated.fieldId)) {
      descriptors.push({ fieldId: calculated.fieldId, name: calculated.name, ordinal: descriptors.length, dataType: 'mixed' });
    }
  }
  const dependencies = new Map<string, string[]>();
  for (const calculated of definitions.values()) dependencies.set(calculated.fieldId, calculatedFieldReferenceIds(calculated.formula, descriptors, definitions, calculated.fieldId));
  const state = new Map<string, 'visiting' | 'visited'>();
  const ordered: PivotCalculatedField[] = [];
  const visit = (fieldId: string, path: string[]): void => {
    const current = state.get(fieldId);
    if (current === 'visited') return;
    if (current === 'visiting') throw new Error(`Pivot calculated field dependency cycle: ${[...path, fieldId].join(' -> ')}`);
    state.set(fieldId, 'visiting');
    for (const dependency of dependencies.get(fieldId) ?? []) visit(dependency, [...path, fieldId]);
    state.set(fieldId, 'visited');
    ordered.push(definitions.get(fieldId)!);
  };
  for (const calculated of definitions.values()) visit(calculated.fieldId, []);
  return { fields: descriptors, definitions, ordered };
}

interface CalculatedFieldEvaluator {
  has(fieldId: string): boolean;
  evaluate(rows: ReadonlyArray<SourceRow>, fieldId: string): PivotScalar | null;
}

function createCalculatedFieldEvaluator(plan: CalculatedFieldPlan, aggregates: PivotAggregatePlanner): CalculatedFieldEvaluator {
  const calculatedIds = new Set(plan.definitions.keys());
  const evaluate = (rows: ReadonlyArray<SourceRow>, fieldId: string): PivotScalar | null => {
    if (!calculatedIds.has(fieldId)) return null;
    const engine = new FormulaEngine({ defaultSheetId: 'pivot-summary' });
    const values = new Map<string, PivotScalar | null>();
    plan.fields.forEach((field, index) => {
      if (!calculatedIds.has(field.fieldId)) {
        const value = aggregates.aggregate(rows, field.fieldId, 'sum');
        values.set(field.fieldId, value);
        engine.setValue({ sheetId: 'pivot-summary', row: 0, column: index }, isPivotError(value) ? null : value);
      }
    });
    for (const calculated of plan.ordered) {
      const index = plan.fields.findIndex((field) => field.fieldId === calculated.fieldId);
      if (index < 0) throw new Error(`Pivot calculated field descriptor is missing: ${calculated.fieldId}`);
      const address = { sheetId: 'pivot-summary', row: 1, column: index };
      engine.setFormula(address, rewriteCalculatedFormula(calculated.formula, plan.fields));
      const value = formulaScalar(engine.getCellValue(address));
      values.set(calculated.fieldId, value);
      engine.setValue({ sheetId: 'pivot-summary', row: 0, column: index }, isPivotError(value) ? null : value);
    }
    return values.get(fieldId) ?? null;
  };
  return { has: (fieldId) => calculatedIds.has(fieldId), evaluate };
}

function toNumber(value: PivotScalar): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string' || value.trim() === '') return null;
  const parsed = Number(value.replace(/[$,%]/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function compare(left: PivotScalar, right: PivotScalar, dataType: PivotFieldDataType | undefined, collator: Intl.Collator): number {
  if (same(left, right)) return 0;
  if (left == null || left === '') return -1;
  if (right == null || right === '') return 1;
  if (isPivotError(left) || isPivotError(right)) {
    if (isPivotError(left) && isPivotError(right)) return collator.compare(left.code, right.code);
    return isPivotError(left) ? 1 : -1;
  }
  if (dataType === 'boolean' && typeof left === 'boolean' && typeof right === 'boolean') return Number(left) - Number(right);
  if (dataType === 'date') {
    const leftDate = pivotTimelineInstant(left);
    const rightDate = pivotTimelineInstant(right);
    if (leftDate !== undefined && rightDate !== undefined) return leftDate - rightDate;
  }
  if (dataType === 'text') return collator.compare(String(left), String(right));
  const leftNumber = pivotNumericValue(left);
  const rightNumber = pivotNumericValue(right);
  if (leftNumber != null && rightNumber != null) return leftNumber - rightNumber;
  return collator.compare(String(left), String(right));
}

/** Every aggregate has its own semantics; no operation falls through to sum. */
export function aggregatePivotValues(rows: ReadonlyArray<{ values: Record<string, PivotScalar> }>, fieldId: string, operation: PivotAggregateFunction): PivotScalar {
  return aggregatePivotValueStream(rows.map((row) => row.values[fieldId] ?? null), operation);
}

function aggregateSourceRows(rows: ReadonlyArray<SourceRow>, fieldId: string, operation: PivotAggregateFunction): PivotScalar {
  return aggregatePivotValueStream(rows.map((row) => sourceRowValue(row, fieldId)), operation);
}

function aggregatePivotValueStream(values: readonly PivotScalar[], operation: PivotAggregateFunction): PivotScalar {
  return aggregatePivotState(buildPivotAggregateState(values), operation);
}

interface PivotAggregateState {
  count: number;
  numericCount: number;
  distinct: Set<string>;
  firstError: Extract<PivotScalar, { kind: 'error' }> | undefined;
  sum: number;
  product: number;
  minimum: number;
  maximum: number;
  mean: number;
  m2: number;
}

function buildPivotAggregateState(values: readonly PivotScalar[]): PivotAggregateState {
  const state: PivotAggregateState = {
    count: 0,
    numericCount: 0,
    distinct: new Set<string>(),
    firstError: undefined,
    sum: 0,
    product: 1,
    minimum: Number.POSITIVE_INFINITY,
    maximum: Number.NEGATIVE_INFINITY,
    mean: 0,
    m2: 0,
  };
  for (const raw of values) {
    if (raw != null && raw !== '') {
      state.count += 1;
      state.distinct.add(pivotMemberKey(createPivotMemberKey(raw)));
    }
    if (!state.firstError && isPivotError(raw)) state.firstError = raw;
    const number = pivotNumericValue(raw);
    if (number == null) continue;
    state.numericCount += 1;
    const delta = number - state.mean;
    state.mean += delta / state.numericCount;
    state.m2 += delta * (number - state.mean);
    state.sum += number;
    state.product *= number;
    state.minimum = Math.min(state.minimum, number);
    state.maximum = Math.max(state.maximum, number);
  }
  return state;
}

function aggregatePivotState(state: PivotAggregateState, operation: PivotAggregateFunction): PivotScalar {
  switch (operation) {
    case 'count': return state.count;
    case 'distinct-count': return state.distinct.size;
    case 'count-numbers': return state.numericCount;
    case 'sum': return state.firstError ?? state.sum;
    case 'average': return state.firstError ?? (state.numericCount ? state.sum / state.numericCount : null);
    case 'min': return state.firstError ?? (state.numericCount ? state.minimum : null);
    case 'max': return state.firstError ?? (state.numericCount ? state.maximum : null);
    case 'product': return state.firstError ?? (state.numericCount ? state.product : null);
    case 'stdev': return state.firstError ?? (state.numericCount < 2 ? null : Math.sqrt(state.m2 / (state.numericCount - 1)));
    case 'stdevp': return state.firstError ?? (!state.numericCount ? null : Math.sqrt(state.m2 / state.numericCount));
    case 'var': return state.firstError ?? (state.numericCount < 2 ? null : state.m2 / (state.numericCount - 1));
    case 'varp': return state.firstError ?? (!state.numericCount ? null : state.m2 / state.numericCount);
    default: return assertNever(operation);
  }
}

class PivotAggregatePlanner {
  private readonly states = new WeakMap<ReadonlyArray<SourceRow>, Map<string, PivotAggregateState>>();

  aggregate(rows: ReadonlyArray<SourceRow>, fieldId: string, operation: PivotAggregateFunction): PivotScalar {
    let fields = this.states.get(rows);
    if (!fields) {
      fields = new Map<string, PivotAggregateState>();
      this.states.set(rows, fields);
    }
    let state = fields.get(fieldId);
    if (!state) {
      state = buildPivotAggregateState(rows.map((row) => sourceRowValue(row, fieldId)));
      fields.set(fieldId, state);
    }
    return aggregatePivotState(state, operation);
  }
}

function assertNever(value: never): never {
  throw new Error(`Unsupported pivot aggregate: ${String(value)}`);
}

function grouped(value: PivotScalar, group?: PivotGroup): PivotScalar {
  if (!group || value == null || value === '') return value;
  if (group.kind === 'manual') {
    const key = createPivotMemberKey(value);
    return group.groups.find((candidate) => candidate.items.some((item) => pivotMemberKeyEquals(item, key)))?.name ?? value;
  }
  if (group.kind === 'number') {
    const number = pivotNumericValue(value);
    if (number == null) return value;
    if (!Number.isFinite(group.interval) || group.interval <= 0) throw new Error('Pivot number grouping interval must be positive');
    const start = group.start ?? 0;
    const result = start + Math.floor((number - start) / group.interval) * group.interval;
    return group.end !== undefined && result > group.end ? group.end : result;
  }
  const date = pivotDate(value);
  if (Number.isNaN(date.getTime())) return value;
  const start = group.start === undefined ? undefined : pivotDate(group.start);
  const end = group.end === undefined ? undefined : pivotDate(group.end);
  if (start && !Number.isNaN(start.getTime()) && date < start) return group.autoStart ? dateGroupLabel(start, group) : value;
  if (end && !Number.isNaN(end.getTime()) && date > end) return group.autoEnd ? dateGroupLabel(end, group) : value;
  return dateGroupLabel(date, group);
}

/** A grouped item keeps its canonical selection key separate from its display caption. */
export interface PivotGroupedFilterMember {
  key: PivotMemberKey;
  value: PivotScalar;
  label: string;
}

function groupedMemberKey(value: PivotScalar, group: PivotGroup): PivotMemberKey {
  if (group.kind === 'manual') {
    const rawKey = createPivotMemberKey(value);
    const owner = group.groups.find((candidate) => candidate.items.some((item) => pivotMemberKeyEquals(item, rawKey)));
    if (owner) return { type: 'text', value: `__pivot_group__:${owner.groupId}` };
  }
  return createPivotMemberKey(grouped(value, group));
}

/** Build the same grouped member domain used by axisGroups for filter surfaces. */
export function buildPivotGroupedFilterMembers(values: readonly PivotScalar[], group: PivotGroup): PivotGroupedFilterMember[] {
  const members = new Map<string, PivotGroupedFilterMember>();
  for (const value of values) {
    const projected = grouped(value, group);
    const key = groupedMemberKey(value, group);
    const identity = pivotMemberKey(key);
    if (!members.has(identity)) members.set(identity, { key, value: projected, label: formatPivotMember(projected) });
  }
  return [...members.values()];
}

function pivotDate(value: PivotScalar): Date {
  if (typeof value === 'number' && Number.isFinite(value)) return new Date(Date.UTC(1899, 11, 30) + value * 86_400_000);
  return new Date(String(value));
}

function dateGroupLabel(date: Date, group: Extract<PivotGroup, { kind: 'date' }>): PivotScalar {
  const units: PivotDateGroupUnit[] = group.units?.length ? group.units : [group.unit];
  const labels = units.map((unit) => {
    if (unit === 'year') return String(date.getFullYear());
    if (unit === 'quarter') return `${date.getFullYear()} Q${Math.floor(date.getMonth() / 3) + 1}`;
    if (unit === 'month') return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    if (unit === 'week') {
      const startOfWeek = group.startOfWeek ?? 0;
      const first = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
      const offset = (first.getUTCDay() - startOfWeek + 7) % 7;
      return `W${Math.floor((Math.floor((date.getTime() - first.getTime()) / 86_400_000) + offset) / 7) + 1}`;
    }
    return date.toISOString().slice(0, 10);
  });
  return labels.length === 1 && units[0] === 'year' ? Number(labels[0]) : labels.join(' / ');
}

function valueSourceFieldId(valueId: string, values: readonly Pick<PivotValueField, 'valueId' | 'fieldId'>[]): string {
  const value = values.find((entry) => entry.valueId === valueId);
  if (!value) throw new Error(`Pivot Values placement is missing: ${valueId}`);
  return value.fieldId;
}

function axisGroups(rows: SourceRow[], placements: PivotFieldPlacement[], fieldCatalog: PivotFieldCatalog, collator: Intl.Collator, values: readonly PivotResultValueField[] = [], calculatedFields?: CalculatedFieldEvaluator, aggregates?: PivotAggregatePlanner): AxisGroup[] {
  const map = new Map<string, AxisGroup>();
  for (const row of rows) {
    const values = placements.map((placement) => grouped(sourceRowValue(row, placement.fieldId), placement.group));
    const key = JSON.stringify(values.map(createPivotMemberKey));
    const group = map.get(key) ?? { values, rows: [], rowSet: new Set<SourceRow>() };
    group.rows.push(row);
    group.rowSet.add(row);
    map.set(key, group);
  }
  const placement = placements[placements.length - 1];
  const dataType = placement ? fieldCatalog.fields.find((field) => field.fieldId === placement.fieldId)?.dataType : undefined;
  const valueSort = placement?.sort?.by === 'value' ? placement.sort : undefined;
  const compareLabels = (left: AxisGroup, right: AxisGroup): number => {
    for (let index = 0; index < left.values.length; index += 1) {
      const fieldType = fieldCatalog.fields.find((field) => field.fieldId === placements[index]?.fieldId)?.dataType ?? dataType;
      const order = compare(left.values[index] ?? null, right.values[index] ?? null, fieldType, collator);
      if (order) return order;
    }
    return 0;
  };
  const result = [...map.values()].sort((left, right) => {
    if (valueSort) {
      const valueField = values.find((value) => value.valueId === valueSort.valueId);
      if (!valueField) throw new Error(`Pivot value sort placement is not in Values: ${valueSort.valueId}`);
      const leftValue = pivotNumericValue(resultValue(left.rows, valueField, valueField.summarizeBy, calculatedFields, aggregates)) ?? 0;
      const rightValue = pivotNumericValue(resultValue(right.rows, valueField, valueField.summarizeBy, calculatedFields, aggregates)) ?? 0;
      const valueOrder = leftValue - rightValue;
      if (valueOrder) return valueSort.direction === 'descending' ? -valueOrder : valueOrder;
      const labelOrder = compareLabels(left, right);
      return valueSort.direction === 'descending' ? -labelOrder : labelOrder;
    }
    return compareLabels(left, right);
  });
  if (!valueSort && placement?.sort?.direction === 'descending') result.reverse();
  return result;
}

function countPivotResultNodes(
  rows: SourceRow[],
  placements: PivotFieldPlacement[],
  depth: number,
  fieldCatalog: PivotFieldCatalog,
  collator: Intl.Collator,
  values: readonly PivotResultValueField[],
  calculatedFields: CalculatedFieldEvaluator,
  aggregates: PivotAggregatePlanner,
  limit: number,
): number {
  if (placements.length === 0) return 1;
  if (depth >= placements.length) return 0;
  const groups = axisGroups(rows, [placements[depth]!], fieldCatalog, collator, values, calculatedFields, aggregates);
  let count = groups.length;
  if (count > limit) return count;
  for (const group of groups) {
    count += countPivotResultNodes(group.rows, placements, depth + 1, fieldCatalog, collator, values, calculatedFields, aggregates, limit - count);
    if (count > limit) return count;
  }
  return count;
}

function assertPivotTaskFootprint(
  definition: PivotDefinition,
  filtered: SourceRow[],
  columns: readonly AxisGroup[],
  values: readonly PivotResultValueField[],
  collator: Intl.Collator,
  calculatedFields: CalculatedFieldEvaluator,
  aggregates: PivotAggregatePlanner,
  targetBounds: { rowCount: number; columnCount: number },
): void {
  const displayOptions = normalizePivotDisplayOptions(definition.presentation?.displayOptions);
  const rowHeaderCount = definition.layout.reportLayout === 'compact' ? 1 : Math.max(definition.layout.rows.length, 1);
  const valueCount = Math.max(values.length, 1);
  const projectedColumns = rowHeaderCount + Math.max(columns.length, 1) * valueCount + (definition.layout.showRowGrandTotals ? valueCount : 0);
  const availableRows = targetBounds.rowCount - definition.target.anchor.row;
  const availableColumns = targetBounds.columnCount - definition.target.anchor.column;
  if (projectedColumns > availableColumns) throw new Error('Pivot target range exceeds the destination worksheet boundary');
  const reportFilterRows = displayOptions.showFieldHeaders
    ? new Set(definition.layout.filters.filter((entry) => entry.scope !== 'field').map((entry) => entry.fieldId)).size
    : 0;
  const fixedRows = 1 + reportFilterRows + (displayOptions.showFieldHeaders ? 1 : 0) + (definition.layout.showColumnGrandTotals ? 1 : 0);
  const nodeLimit = Math.max(0, availableRows - fixedRows);
  const nodeCount = countPivotResultNodes(filtered, definition.layout.rows, 0, definition.fieldCatalog, collator, values, calculatedFields, aggregates, nodeLimit);
  if (nodeCount > nodeLimit) throw new Error('Pivot target range exceeds the destination worksheet boundary');
  const resultCellCount = nodeCount * Math.max(columns.length, 1)
    + (definition.layout.showRowGrandTotals ? nodeCount : 0)
    + Math.max(columns.length, 1)
    + 1;
  if (resultCellCount > PIVOT_MAX_RESULT_CELL_COUNT) {
    throw new Error(`Pivot result cell limit exceeded: ${String(resultCellCount)} > ${String(PIVOT_MAX_RESULT_CELL_COUNT)}`);
  }
  const provenanceReferences = filtered.length * (4 + Math.max(1, definition.layout.rows.length) * 2);
  if (provenanceReferences > PIVOT_MAX_PROVENANCE_REFERENCE_COUNT) {
    throw new Error(`Pivot result provenance limit exceeded: ${String(provenanceReferences)} > ${String(PIVOT_MAX_PROVENANCE_REFERENCE_COUNT)}`);
  }
}

function manualFilterMatches(value: PivotScalar, filter: Extract<PivotFilter, { kind: 'manual' }>, group?: PivotGroup): boolean {
  if (filter.mode === 'all') return true;
  const key = group ? groupedMemberKey(value, group) : createPivotMemberKey(value);
  const included = (filter.memberKeys ?? []).some((candidate) => pivotMemberKeyEquals(candidate, key));
  return filter.mode === 'include' ? included : !included;
}

function dynamicDateBounds(kind: NonNullable<Extract<PivotFilter, { kind: 'condition'; family: 'date' }>['dynamic']>, now = Date.now()): [number, number] {
  const today = new Date(Math.floor(now / 86_400_000) * 86_400_000);
  const startOfWeek = new Date(today);
  startOfWeek.setUTCDate(today.getUTCDate() - ((today.getUTCDay() + 6) % 7));
  const startOfMonth = new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth(), 1));
  const startOfQuarter = new Date(Date.UTC(today.getUTCFullYear(), Math.floor(today.getUTCMonth() / 3) * 3, 1));
  const startOfYear = new Date(Date.UTC(today.getUTCFullYear(), 0, 1));
  const shift = (base: Date, months: number, days = 0): [number, number] => {
    const start = new Date(base);
    start.setUTCMonth(start.getUTCMonth() + months);
    start.setUTCDate(start.getUTCDate() + days);
    const end = new Date(start);
    if (months === 0) end.setUTCDate(end.getUTCDate() + 1);
    else end.setUTCMonth(end.getUTCMonth() + months);
    return [start.getTime(), end.getTime()];
  };
  if (kind === 'today') return [today.getTime(), today.getTime() + 86_400_000];
  if (kind === 'yesterday') return shift(today, 0, -1);
  if (kind === 'tomorrow') return shift(today, 0, 1);
  if (kind === 'this-week') return [startOfWeek.getTime(), startOfWeek.getTime() + 7 * 86_400_000];
  if (kind === 'last-week') return [startOfWeek.getTime() - 7 * 86_400_000, startOfWeek.getTime()];
  if (kind === 'next-week') return [startOfWeek.getTime() + 7 * 86_400_000, startOfWeek.getTime() + 14 * 86_400_000];
  if (kind === 'this-month') return [startOfMonth.getTime(), new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() + 1, 1)).getTime()];
  if (kind === 'last-month') return [new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() - 1, 1)).getTime(), startOfMonth.getTime()];
  if (kind === 'next-month') return [new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() + 1, 1)).getTime(), new Date(Date.UTC(today.getUTCFullYear(), today.getUTCMonth() + 2, 1)).getTime()];
  if (kind === 'this-quarter') return [startOfQuarter.getTime(), new Date(Date.UTC(today.getUTCFullYear(), startOfQuarter.getUTCMonth() + 3, 1)).getTime()];
  if (kind === 'last-quarter') return [new Date(Date.UTC(today.getUTCFullYear(), startOfQuarter.getUTCMonth() - 3, 1)).getTime(), startOfQuarter.getTime()];
  if (kind === 'next-quarter') return [new Date(Date.UTC(today.getUTCFullYear(), startOfQuarter.getUTCMonth() + 3, 1)).getTime(), new Date(Date.UTC(today.getUTCFullYear(), startOfQuarter.getUTCMonth() + 6, 1)).getTime()];
  if (kind === 'this-year') return [startOfYear.getTime(), new Date(Date.UTC(today.getUTCFullYear() + 1, 0, 1)).getTime()];
  if (kind === 'last-year') return [new Date(Date.UTC(today.getUTCFullYear() - 1, 0, 1)).getTime(), startOfYear.getTime()];
  if (kind === 'next-year') return [new Date(Date.UTC(today.getUTCFullYear() + 1, 0, 1)).getTime(), new Date(Date.UTC(today.getUTCFullYear() + 2, 0, 1)).getTime()];
  return [startOfYear.getTime(), today.getTime() + 86_400_000];
}

function dateFilterMatches(value: PivotScalar, filter: Extract<PivotFilter, { kind: 'condition'; family: 'date' }>): boolean {
  const instant = pivotTimelineInstant(value);
  if (instant === undefined) return false;
  if (filter.dynamic) {
    const [start, end] = dynamicDateBounds(filter.dynamic);
    return instant >= start && instant < end;
  }
  const first = pivotTimelineInstant(filter.value);
  const second = filter.value2 === undefined ? undefined : pivotTimelineInstant(filter.value2);
  if (first === undefined || ((filter.operator === 'between' || filter.operator === 'not-between') && second === undefined)) return false;
  const left = filter.wholeDay ? Math.floor(instant / 86_400_000) : instant;
  const right = filter.wholeDay ? Math.floor(first / 86_400_000) : first;
  if (filter.operator === 'equals') return left === right;
  if (filter.operator === 'not-equals') return left !== right;
  if (filter.operator === 'before') return left < right;
  if (filter.operator === 'after') return left > right;
  const upper = filter.wholeDay ? Math.floor(second! / 86_400_000) : second!;
  const inside = left >= Math.min(right, upper) && left <= Math.max(right, upper);
  return filter.operator === 'between' ? inside : !inside;
}

function labelFilterMatches(value: PivotScalar, filter: Extract<PivotFilter, { kind: 'condition'; family: 'label' }>, collator: Intl.Collator): boolean {
  const text = String(value ?? '');
  const operand = String(filter.value ?? '');
  if (filter.operator === 'equals') return text === operand;
  if (filter.operator === 'not-equals') return text !== operand;
  if (filter.operator === 'begins-with') return text.startsWith(operand);
  if (filter.operator === 'not-begins-with') return !text.startsWith(operand);
  if (filter.operator === 'ends-with') return text.endsWith(operand);
  if (filter.operator === 'not-ends-with') return !text.endsWith(operand);
  if (filter.operator === 'contains') return text.includes(operand);
  if (filter.operator === 'not-contains') return !text.includes(operand);
  const order = collator.compare(text, operand);
  if (filter.operator === 'greater-than') return order > 0;
  if (filter.operator === 'greater-or-equal') return order >= 0;
  if (filter.operator === 'less-than') return order < 0;
  if (filter.operator === 'less-or-equal') return order <= 0;
  const upper = String(filter.value2 ?? '');
  const inside = collator.compare(text, operand) >= 0 && collator.compare(text, upper) <= 0;
  return filter.operator === 'between' ? inside : !inside;
}

function groupedPlacementForFilter(definition: PivotDefinition, filter: PivotFilter): PivotFieldPlacement | undefined {
  if ((filter.scope ?? 'report') !== 'field' || filter.kind === 'top-items') return undefined;
  if (filter.kind === 'condition' && filter.valueId !== undefined) return undefined;
  return [...definition.layout.rows, ...definition.layout.columns].find((placement) => placement.fieldId === filter.fieldId && placement.group);
}

function groupedDateFilterMatches(value: PivotScalar, filter: Extract<PivotFilter, { kind: 'condition'; family: 'date' }>, group: PivotGroup, collator: Intl.Collator): boolean {
  if (filter.dynamic) return dateFilterMatches(value, filter);
  const projectedValue = grouped(value, group);
  const projectedFilter = { ...filter, value: grouped(filter.value, group), ...(filter.value2 === undefined ? {} : { value2: grouped(filter.value2, group) }) };
  const left = String(projectedValue ?? '');
  const right = String(projectedFilter.value ?? '');
  const order = collator.compare(left, right);
  if (filter.operator === 'equals') return order === 0;
  if (filter.operator === 'not-equals') return order !== 0;
  if (filter.operator === 'before') return order < 0;
  if (filter.operator === 'after') return order > 0;
  const upper = String(projectedFilter.value2 ?? '');
  const inside = collator.compare(left, right) >= 0 && collator.compare(left, upper) <= 0;
  return filter.operator === 'between' ? inside : !inside;
}

type PivotSourceFilter = Exclude<PivotFilter, { kind: 'condition'; family: 'value' }>;

function matchesFilter(row: SourceRow, filter: PivotFilter, collator: Intl.Collator, definition?: PivotDefinition): boolean {
  if (filter.kind === 'condition' && filter.family === 'value') throw new Error('Pivot value filters must be evaluated against aggregated Pivot items');
  const sourceFilter = filter as PivotSourceFilter;
  const fieldId = sourceFilter.fieldId;
  const rawValue = sourceRowValue(row, fieldId);
  const placement = definition ? groupedPlacementForFilter(definition, sourceFilter) : undefined;
  const value = placement?.group ? grouped(rawValue, placement.group) : rawValue;
  if (sourceFilter.kind === 'top-items') return true;
  if (sourceFilter.kind === 'manual') return manualFilterMatches(rawValue, sourceFilter, placement?.group);
  if (sourceFilter.family === 'date') return placement?.group ? groupedDateFilterMatches(rawValue, sourceFilter, placement.group, collator) : dateFilterMatches(value, sourceFilter);
  if (sourceFilter.family === 'label') return labelFilterMatches(value, sourceFilter, collator);
  throw new Error('Unsupported Pivot source filter family');
}

function matchesValueFilter(value: PivotScalar, filter: Extract<PivotFilter, { kind: 'condition'; family: 'value' }>, collator: Intl.Collator): boolean {
  const leftNumber = pivotNumericValue(value);
  const rightNumber = pivotNumericValue(filter.value);
  const upperNumber = filter.value2 === undefined ? null : pivotNumericValue(filter.value2);
  const order = leftNumber != null && rightNumber != null ? leftNumber - rightNumber : compare(value, filter.value, undefined, collator);
  switch (filter.operator) {
    case 'equals': return same(value, filter.value);
    case 'not-equals': return !same(value, filter.value);
    case 'greater-than': return order > 0;
    case 'greater-or-equal': return order >= 0;
    case 'less-than': return order < 0;
    case 'less-or-equal': return order <= 0;
    case 'between': return filter.value2 !== undefined && order >= 0 && (leftNumber != null && upperNumber != null ? leftNumber <= upperNumber : compare(value, filter.value2, undefined, collator) <= 0);
    case 'not-between': return filter.value2 !== undefined && !(order >= 0 && (leftNumber != null && upperNumber != null ? leftNumber <= upperNumber : compare(value, filter.value2, undefined, collator) <= 0));
    default: return false;
  }
}

/**
 * Apply aggregate value predicates to Pivot item buckets.
 *
 * Label/manual/date predicates intentionally run before this stage because
 * they restrict source members. A value predicate is different: its left
 * operand is the selected Values placement's configured aggregate for one
 * item, never a raw source-row member. The preceding row/column placements
 * form the parent context for field-scoped filters; report-scoped filters
 * aggregate the target field globally.
 */
function applyValueFilters(
  rows: SourceRow[],
  filters: readonly PivotFilter[],
  definition: PivotDefinition,
  calculatedFields: CalculatedFieldEvaluator,
  collator: Intl.Collator,
  aggregates: PivotAggregatePlanner,
): SourceRow[] {
  const source = rows;
  let result = rows;
  for (const rawFilter of filters) {
    if (rawFilter.kind !== 'condition' || rawFilter.family !== 'value') continue;
    const filter = rawFilter;
    if (!filter.valueId) throw new Error(`Pivot value filter requires valueId for ${filter.fieldId}`);
    const valueField = definition.layout.values.find((entry) => entry.valueId === filter.valueId);
    if (!valueField) throw new Error(`Unknown Pivot Values placement: ${filter.valueId}`);

    const rowPlacements = definition.layout.rows.filter((placement) => placement.fieldId === filter.fieldId);
    const columnPlacements = definition.layout.columns.filter((placement) => placement.fieldId === filter.fieldId);
    const fieldScoped = (filter.scope ?? 'report') === 'field';
    if (fieldScoped && rowPlacements.length + columnPlacements.length !== 1) {
      throw new Error(`Pivot value filter field must resolve to exactly one axis placement: ${filter.fieldId}`);
    }
    const axis = !fieldScoped ? undefined : rowPlacements.length === 1 && columnPlacements.length === 0
      ? definition.layout.rows
      : columnPlacements.length === 1 && rowPlacements.length === 0
        ? definition.layout.columns
        : undefined;
    const targetIndex = axis?.findIndex((placement) => placement.fieldId === filter.fieldId) ?? -1;
    const contextPlacements = axis && targetIndex >= 0 ? axis.slice(0, targetIndex + 1) : undefined;
    const targetPlacement = [...rowPlacements, ...columnPlacements][0];
    const buckets = new Map<string, SourceRow[]>();
    for (const row of source) {
      const keyValues = contextPlacements?.map((placement) => grouped(sourceRowValue(row, placement.fieldId), placement.group))
        ?? [grouped(sourceRowValue(row, filter.fieldId), targetPlacement?.group)];
      const key = JSON.stringify(keyValues.map(createPivotMemberKey));
      const bucket = buckets.get(key) ?? [];
      bucket.push(row);
      buckets.set(key, bucket);
    }
    const aggregateField: PivotResultValueField = { ...valueField, sourceFieldId: valueField.fieldId };
    const accepted = new Set<SourceRow>();
    for (const bucket of buckets.values()) {
      const aggregate = resultValue(bucket, aggregateField, valueField.summarizeBy, calculatedFields, aggregates);
      if (matchesValueFilter(aggregate, filter, collator)) bucket.forEach((row) => accepted.add(row));
    }
    result = result.filter((row) => accepted.has(row));
  }
  return result;
}

function topItems(
  rows: SourceRow[],
  filters: PivotFilter[],
  values: readonly PivotValueField[],
  calculatedFields?: CalculatedFieldEvaluator,
  definition?: PivotDefinition,
  aggregates?: PivotAggregatePlanner,
): SourceRow[] {
  let result = rows;
  for (const filter of filters) {
    if (filter.kind !== 'top-items') continue;
    const fieldId = filter.fieldId;
    const valueField = values.find((value) => value.valueId === filter.valueId);
    if (!valueField) throw new Error(`Pivot top-items references an unknown Values placement: ${filter.valueId}`);
    validateTopBottomDirection(filter.direction);
    validateTopBottomThreshold(filter.mode, filter.threshold);
    const groupedField = definition
      ? [...definition.layout.rows, ...definition.layout.columns].find((placement) => placement.fieldId === fieldId)?.group
      : undefined;
    const buckets = new Map<string, SourceRow[]>();
    for (const row of result) {
      const member = groupedField ? grouped(sourceRowValue(row, fieldId), groupedField) : sourceRowValue(row, fieldId);
      const key = pivotMemberKey(createPivotMemberKey(member));
      const bucket = buckets.get(key) ?? [];
      bucket.push(row);
      buckets.set(key, bucket);
    }
    const ranked = [...buckets.entries()].map(([key, bucket]) => ({
      key,
      bucket,
      aggregate: pivotNumericValue(resultValue(bucket, { ...valueField, sourceFieldId: valueField.fieldId }, valueField.summarizeBy, calculatedFields, aggregates)),
    }));
    ranked.sort((left, right) => {
      const leftValue = left.aggregate ?? 0;
      const rightValue = right.aggregate ?? 0;
      const valueOrder = filter.direction === 'top' ? rightValue - leftValue : leftValue - rightValue;
      return valueOrder || left.key.localeCompare(right.key);
    });
    let selected: typeof ranked;
    if (filter.mode === 'items') {
      selected = ranked.slice(0, filter.threshold);
    } else {
      // Excel's Percent and Sum modes select the ranked prefix whose
      // aggregate reaches the requested target; they do not compare every
      // member independently with the threshold.  This is also what keeps a
      // selected Average/Count/Min/Max Values placement authoritative.
      const target = filter.mode === 'percent'
        ? ranked.reduce((total, entry) => total + (entry.aggregate ?? 0), 0) * filter.threshold / 100
        : filter.threshold;
      let accumulated = 0;
      selected = [];
      for (const entry of ranked) {
        selected.push(entry);
        accumulated += entry.aggregate ?? 0;
        if (accumulated >= target) break;
      }
    }
    result = selected.flatMap((entry) => entry.bucket);
  }
  return result;
}

function matchesSlicer(row: SourceRow, slicer: PivotSlicerDrawingPayload, fieldId: string): boolean {
  const { filter } = slicer;
  if (filter.mode === 'all') return true;
  const included = filter.memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, createPivotMemberKey(sourceRowValue(row, fieldId))));
  return filter.mode === 'include' ? included : !included;
}

export interface PivotTaskControl {
  drawingId: string;
  payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload;
  fieldId: string;
}

export function collectPivotTaskControls(workbook: WorkbookModel, pivot: PivotModel): PivotTaskControl[] {
  return workbook.getSheets().flatMap((sheet) => sheet.drawings.flatMap((drawing) => {
    if (drawing.kind !== 'slicer' && drawing.kind !== 'timeline') return [];
    const payload = sheet.drawingPayloads.get(drawing.payloadId);
    if (!payload || (payload.kind !== 'slicer' && payload.kind !== 'timeline')) return [];
    if (payload.pivotId === pivot.id) return [{ drawingId: drawing.id, payload, fieldId: payload.fieldId }];
    const connection = payload.connections?.find((candidate) => candidate.pivotId === pivot.id);
    return connection ? [{ drawingId: drawing.id, payload, fieldId: connection.fieldId }] : [];
  }));
}

function matchesTimeline(row: SourceRow, timeline: PivotTimelineDrawingPayload, fieldId: string, bounds: PivotTimelinePeriodBounds): boolean {
  const raw = sourceRowValue(row, fieldId);
  if (raw == null || raw === '') return false;
  const instant = pivotTimelineInstant(raw);
  if (instant === undefined) return false;
  return instant >= (bounds.start ?? Number.NEGATIVE_INFINITY)
    && instant < (bounds.endExclusive ?? Number.POSITIVE_INFINITY);
}

function matchesControls(rows: SourceRow[], controls: readonly PivotTaskControl[], excludedSlicerDrawingId?: string): SourceRow[] {
  const activeControls = controls.filter((entry) => entry.drawingId !== excludedSlicerDrawingId);
  const slicers = activeControls.filter((entry): entry is PivotTaskControl & { payload: PivotSlicerDrawingPayload } => entry.payload.kind === 'slicer');
  const timelines = activeControls.filter((entry): entry is PivotTaskControl & { payload: PivotTimelineDrawingPayload } => entry.payload.kind === 'timeline');
  const timelineBounds = timelines.map((entry) => normalizePivotTimelinePeriod(entry.payload.period));
  return rows.filter((row) => slicers.every((entry) => matchesSlicer(row, entry.payload, entry.fieldId))
    && timelines.every((entry, index) => matchesTimeline(row, entry.payload, entry.fieldId, timelineBounds[index]!)));
}

function slicerItemProjection(
  definition: PivotDefinition,
  rows: SourceRow[],
  drawingId: string,
  payload: PivotSlicerDrawingPayload,
  collator: Intl.Collator,
  calculatedFields: CalculatedFieldEvaluator,
  controls: readonly PivotTaskControl[],
  aggregates: PivotAggregatePlanner,
): PivotSlicerItemProjection[] {
  const fieldValues = rows.map((row) => sourceRowValue(row, payload.fieldId));
  const members = new Map<string, PivotSlicerItemProjection>();
  for (const value of fieldValues) {
    const key = createPivotMemberKey(value);
    const identity = pivotMemberKey(key);
    if (!members.has(identity)) members.set(identity, { key, value, label: formatPivotMember(value), selected: false, hasData: false });
  }
  const filteredRows = matchesControls(rows, controls, drawingId)
    .filter((row) => definition.layout.filters.filter((filter) => filter.kind !== 'top-items' && !(filter.kind === 'condition' && filter.family === 'value')).every((filter) => matchesFilter(row, filter, collator, definition)));
  const valueFilteredRows = applyValueFilters(filteredRows, definition.layout.filters, definition, calculatedFields, collator, aggregates);
  const availableRows = topItems(valueFilteredRows, definition.layout.filters, definition.layout.values, calculatedFields, definition, aggregates);
  const available = new Set(availableRows.map((row) => pivotMemberKey(createPivotMemberKey(sourceRowValue(row, payload.fieldId)))));
  for (const item of members.values()) {
    item.hasData = available.has(pivotMemberKey(item.key));
    const included = payload.filter.memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, item.key));
    item.selected = payload.filter.mode === 'all' || (payload.filter.mode === 'include' ? included : !included);
  }
  const sorted = [...members.values()].sort((left, right) => collator.compare(left.label, right.label));
  if (payload.settings.sort === 'descending') sorted.reverse();
  if (payload.settings.noDataItemsLast) sorted.sort((left, right) => Number(right.hasData) - Number(left.hasData));
  return sorted;
}

function resultValueFields(layout: PivotLayout): PivotResultValueField[] {
  const customFunctions = [...layout.rows, ...layout.columns].flatMap((placement) => placement.subtotal?.mode === 'custom'
    ? placement.subtotal.functions.map((fn) => ({ fieldId: placement.fieldId, fn }))
    : []);
  if (!customFunctions.length) return layout.values.map((field) => ({ ...field, sourceFieldId: field.fieldId }));
  return layout.values.flatMap((field) => {
    const base = { ...field, sourceFieldId: field.fieldId };
    const extras = customFunctions.filter(({ fieldId, fn }, index, all) => fn !== field.summarizeBy && all.findIndex((candidate) => candidate.fieldId === fieldId && candidate.fn === fn) === index).map(({ fieldId, fn }) => ({
      ...field,
      valueId: `${field.valueId}:subtotal:${fn}`,
      sourceFieldId: field.fieldId,
      subtotalFunction: fn,
      subtotalFieldId: fieldId,
      displayName: `${field.displayName ?? field.fieldId} (${fn})`,
    }));
    return [base, ...extras];
  });
}

function resultValue(rows: ReadonlyArray<SourceRow>, value: PivotResultValueField, operation: PivotAggregateFunction, calculatedFields?: CalculatedFieldEvaluator, aggregates?: PivotAggregatePlanner): PivotScalar {
  if (calculatedFields?.has(value.sourceFieldId)) return calculatedFields.evaluate(rows, value.sourceFieldId);
  return aggregates?.aggregate(rows, value.sourceFieldId, operation) ?? aggregateSourceRows(rows, value.sourceFieldId, operation);
}

function resultCells(rows: SourceRow[], columns: AxisGroup[], values: PivotResultValueField[], nodePath: string[], kind: PivotResultCell['kind'] = 'detail', subtotalFieldId?: string, calculatedFields?: CalculatedFieldEvaluator, aggregates?: PivotAggregatePlanner): PivotResultCell[] {
  return columns.map((column, columnIndex) => {
    const columnRows = rows.length <= column.rows.length
      ? rows.filter((candidate) => column.rowSet.has(candidate))
      : (() => {
          const nodeRows = new Set(rows);
          return column.rows.filter((candidate) => nodeRows.has(candidate));
        })();
    return {
      id: `${nodePath.join('/') || 'root'}|column:${columnIndex}`,
      nodePath,
      kind,
      columnPath: column.values,
      sourceRowPaths: columnRows.flatMap((row) => sourceRowPaths(row)),
      values: values.map((value) => resultValue(columnRows, value, kind === 'subtotal' && value.subtotalFieldId === subtotalFieldId
        ? value.subtotalFunction ?? value.summarizeBy
        : value.summarizeBy, calculatedFields, aggregates)),
    };
  });
}

function resultGrandTotalCell(rows: SourceRow[], values: PivotResultValueField[], nodePath: string[], subtotalFieldId?: string, calculatedFields?: CalculatedFieldEvaluator, aggregates?: PivotAggregatePlanner): PivotResultCell {
  return {
    id: `${nodePath.join('/') || 'root'}|grand-total:row`,
    nodePath,
    kind: 'grand-total',
    columnPath: [],
    sourceRowPaths: rows.flatMap((row) => sourceRowPaths(row)),
    values: values.map((value) => resultValue(rows, value, subtotalFieldId === value.subtotalFieldId
      ? value.subtotalFunction ?? value.summarizeBy
      : value.summarizeBy, calculatedFields, aggregates)),
  };
}

function resultNodes(rows: SourceRow[], placements: PivotFieldPlacement[], depth: number, columns: AxisGroup[], values: PivotResultValueField[], subtotalLocation: PivotLayout['subtotalLocation'], showRowGrandTotals: boolean, fieldCatalog: PivotFieldCatalog, collator: Intl.Collator, calculatedFields?: CalculatedFieldEvaluator, aggregates?: PivotAggregatePlanner, prefix: string[] = []): PivotResultNode[] {
  // A Pivot with no Row fields still owns one data row: the root aggregation
  // crossing every Column path and Values placement. Grand Total is a
  // separate axis total and must not stand in for this matrix row.
  if (depth >= placements.length) {
    if (placements.length !== 0 || depth !== 0) return [];
    const path = ['__root__'];
    return [{
      nodeId: path[0],
      path,
      kind: 'leaf',
      key: null,
      label: 'Values',
      depth: 0,
      children: [],
      values: resultCells(rows, columns, values, path, 'detail', undefined, calculatedFields, aggregates),
      ...(showRowGrandTotals ? { rowGrandTotal: resultGrandTotalCell(rows, values, path, undefined, calculatedFields, aggregates) } : {}),
      subtotal: false,
      sourceRowPaths: rows.flatMap((row) => sourceRowPaths(row)),
    }];
  }
  const placement = placements[depth]!;
  return axisGroups(rows, [placement], fieldCatalog, collator, values, calculatedFields, aggregates).map((group) => {
    const fieldId = placement.fieldId;
    const member = createPivotMemberKey(group.values[0] ?? null);
    const path = [...prefix, `${fieldId}=${pivotMemberKey(member)}`];
    const children = resultNodes(group.rows, placements, depth + 1, columns, values, subtotalLocation, showRowGrandTotals, fieldCatalog, collator, calculatedFields, aggregates, path);
    const leaf = children.length === 0;
    const subtotal = !leaf && subtotalLocation !== 'off' && placement.subtotal?.mode !== 'none';
    return {
      nodeId: path.join('/'),
      path,
      kind: subtotal ? 'subtotal' : 'leaf',
      fieldId,
      memberKey: member,
      key: group.values[0] ?? null,
      label: display(group.values[0] ?? null),
      depth,
      children,
      values: resultCells(group.rows, columns, values, path, subtotal ? 'subtotal' : 'detail', subtotal ? placement.fieldId : undefined, calculatedFields, aggregates),
      ...(showRowGrandTotals ? { rowGrandTotal: resultGrandTotalCell(group.rows, values, path, subtotal ? placement.fieldId : undefined, calculatedFields, aggregates) } : {}),
      subtotal,
      sourceRowPaths: group.rows.flatMap((row) => sourceRowPaths(row)),
    };
  });
}

interface PivotShowAsCellContext {
  cell: PivotResultCell;
  node?: PivotResultNode;
  parent?: PivotResultNode;
  columnIndex: number;
  kind: 'detail' | 'subtotal' | 'grand-total';
}

interface PivotShowAsAxisResolution {
  base: number | null;
  same: boolean;
  series: number[];
  position: number;
}

/**
 * Apply Show Values As from one immutable result matrix.
 *
 * The result tree contains three different calculation domains: detail rows,
 * subtotal rows and the grand-total row.  Keeping those contexts explicit is
 * important because subtotal values are valid Pivot members in their own
 * right; they must never be looked up in a leaf-only sequence.
 */
function applyShowAs(tree: PivotResultTree, fields: PivotValueField[], layout: PivotLayout): void {
  const raw = new Map<PivotResultCell, PivotScalar[]>();
  const contexts: PivotShowAsCellContext[] = [];
  const visit = (nodes: PivotResultNode[], parent?: PivotResultNode) => nodes.forEach((node) => {
    node.values.forEach((cell, columnIndex) => {
      raw.set(cell, [...cell.values]);
      contexts.push({ cell, node, parent, columnIndex, kind: node.subtotal ? 'subtotal' : 'detail' });
    });
    visit(node.children, node);
  });
  visit(tree.rows);
  if (tree.grandTotal) {
    raw.set(tree.grandTotal, [...tree.grandTotal.values]);
    contexts.push({ cell: tree.grandTotal, columnIndex: 0, kind: 'grand-total' });
  }
  for (const cell of tree.columnGrandTotals ?? []) {
    raw.set(cell, [...cell.values]);
    contexts.push({ cell, columnIndex: 0, kind: 'grand-total' });
  }
  const visitRowTotals = (nodes: PivotResultNode[]) => nodes.forEach((node) => {
    if (node.rowGrandTotal) {
      raw.set(node.rowGrandTotal, [...node.rowGrandTotal.values]);
      contexts.push({ cell: node.rowGrandTotal, node, columnIndex: 0, kind: 'grand-total' });
    }
    visitRowTotals(node.children);
  });
  visitRowTotals(tree.rows);

  const rawValue = (cell: PivotResultCell | undefined, index: number): PivotScalar | null => cell ? raw.get(cell)?.[index] ?? null : null;
  const grandValues = tree.grandTotal ? raw.get(tree.grandTotal) ?? [] : [];
  const rowContexts = contexts.filter((context) => context.kind !== 'grand-total');
  const leafContexts = rowContexts.filter((context) => context.node?.children.length === 0);

  const numericSum = (cells: PivotShowAsCellContext[], valueIndex: number, columnIndex: number): number => cells.reduce((sum, context) => {
    const cell = context.node?.values[columnIndex];
    return sum + (pivotNumericValue(rawValue(cell, valueIndex)) ?? 0);
  }, 0);

  const rowFieldIds = layout.rows.map((placement) => placement.fieldId);
  const columnFieldIds = layout.columns.map((placement) => placement.fieldId);
  const nodeContexts = new Map<string, PivotShowAsCellContext>();
  for (const context of rowContexts) {
    const id = context.node?.nodeId ?? context.node?.path?.join('/') ?? '';
    if (id && context.node && !nodeContexts.has(id)) nodeContexts.set(id, context);
  }
  const sameScalar = (left: PivotScalar | undefined, right: PivotScalar): boolean => same(left ?? null, right);
  const baseItemMatches = (value: PivotScalar | undefined, item: PivotShowAsBaseItem): boolean => {
    if (item === 'previous' || item === 'next') return false;
    return sameScalar(value, pivotScalarFromMemberKey(item));
  };

  const rowResolution = (context: PivotShowAsCellContext, spec: Extract<NonNullable<PivotValueField['showAs']>, { baseFieldId: string }>, valueIndex: number): PivotShowAsAxisResolution | undefined => {
    if (context.kind === 'grand-total' || !context.node?.path) return undefined;
    const depth = rowFieldIds.indexOf(spec.baseFieldId);
    if (depth < 0 || context.node.depth < depth) return undefined;
    const currentPath = context.node.path[depth];
    const prefix = context.node.path.slice(0, depth).join('\u001f');
    const candidates = [...nodeContexts.values()].filter((candidate) => candidate.node?.fieldId === spec.baseFieldId
      && candidate.node.path?.slice(0, depth).join('\u001f') === prefix);
    if (candidates.length === 0) return undefined;
    const currentIndex = candidates.findIndex((candidate) => candidate.node?.path?.[depth] === currentPath);
    let position = currentIndex;
    let target = currentIndex;
    if ('baseItem' in spec) {
      if (spec.baseItem === 'previous' || spec.baseItem === 'next') target = currentIndex + (spec.baseItem === 'previous' ? -1 : 1);
      else target = candidates.findIndex((candidate) => baseItemMatches(candidate.node?.memberKey?.value, spec.baseItem));
    }
    if (target < 0 || target >= candidates.length) return undefined;
    const series = candidates.map((candidate) => pivotNumericValue(rawValue(candidate.node?.values[context.columnIndex], valueIndex)) ?? 0);
    return { base: pivotNumericValue(rawValue(candidates[target]?.node?.values[context.columnIndex], valueIndex)), same: target === currentIndex, series, position };
  };

  const columnResolution = (context: PivotShowAsCellContext, spec: Extract<NonNullable<PivotValueField['showAs']>, { baseFieldId: string }>, valueIndex: number): PivotShowAsAxisResolution | undefined => {
    if (context.kind === 'grand-total' || !context.node) return undefined;
    const depth = columnFieldIds.indexOf(spec.baseFieldId);
    if (depth < 0) return undefined;
    const currentPath = tree.columnPaths[context.columnIndex];
    if (!currentPath) return undefined;
    const candidateIndexes = tree.columnPaths.map((path, index) => ({ path, index })).filter(({ path }) => path.length > depth
      && columnFieldIds.every((fieldId, index) => index === depth || same(path[index] ?? null, currentPath[index] ?? null)));
    const currentIndex = candidateIndexes.findIndex(({ index }) => index === context.columnIndex);
    let target = currentIndex;
    if ('baseItem' in spec) {
      if (spec.baseItem === 'previous' || spec.baseItem === 'next') target = currentIndex + (spec.baseItem === 'previous' ? -1 : 1);
      else target = candidateIndexes.findIndex(({ path }) => baseItemMatches(path[depth], spec.baseItem));
    }
    if (target < 0 || target >= candidateIndexes.length) return undefined;
    const series = candidateIndexes.map(({ index }) => pivotNumericValue(rawValue(context.node?.values[index], valueIndex)) ?? 0);
    const targetColumn = candidateIndexes[target]?.index;
    return { base: targetColumn === undefined ? null : pivotNumericValue(rawValue(context.node.values[targetColumn], valueIndex)), same: targetColumn === context.columnIndex, series, position: currentIndex };
  };

  const resolveAxis = (context: PivotShowAsCellContext, spec: Extract<NonNullable<PivotValueField['showAs']>, { baseFieldId: string }>, valueIndex: number): PivotShowAsAxisResolution | undefined => {
    if (rowFieldIds.includes(spec.baseFieldId)) return rowResolution(context, spec, valueIndex);
    if (columnFieldIds.includes(spec.baseFieldId)) return columnResolution(context, spec, valueIndex);
    return undefined;
  };

  const transform = (
    spec: NonNullable<PivotValueField['showAs']>,
    current: number,
    grand: number | null,
    rowTotal: number,
    columnTotal: number,
    parentTotal: number | null,
    context: PivotShowAsCellContext,
    valueIndex: number,
  ): number | null => {
    if (spec.kind === 'normal') return current;
    if (context.kind === 'grand-total') {
      // A grand total has no row/column member coordinate. It is nevertheless
      // part of the calculation domain: total-relative modes resolve to the
      // identity, differences to zero, running totals to the final aggregate,
      // and rank/index to the sole total member.
      if (spec.kind === 'grand-percentage' || spec.kind === 'row-percentage' || spec.kind === 'column-percentage' || spec.kind === 'parent-percentage') return grand ? current / grand : null;
      if (spec.kind === 'difference' || spec.kind === 'percentage-difference') return null;
      if (spec.kind === 'running-total') return current;
      if (spec.kind === 'percentage-running-total') return grand ? current / grand : null;
      if (spec.kind === 'rank') return 1;
      if (spec.kind === 'index') return grand != null && rowTotal && columnTotal ? current * grand / rowTotal / columnTotal : null;
    }
    if (spec.kind === 'grand-percentage') return grand ? current / grand : null;
    if (spec.kind === 'row-percentage') return rowTotal ? current / rowTotal : null;
    if (spec.kind === 'column-percentage') return columnTotal ? current / columnTotal : null;
    if (spec.kind === 'parent-percentage') return parentTotal == null ? null : parentTotal ? current / parentTotal : null;
    if (spec.kind === 'difference' || spec.kind === 'percentage-difference') {
      const resolved = resolveAxis(context, spec, valueIndex);
      if (!resolved || resolved.same || resolved.base == null) return null;
      return spec.kind === 'difference' ? current - resolved.base : resolved.base ? (current - resolved.base) / resolved.base : null;
    }
    if (spec.kind === 'running-total' || spec.kind === 'percentage-running-total') {
      const resolved = resolveAxis(context, spec, valueIndex);
      if (!resolved || resolved.position < 0) return null;
      const cumulative = resolved.series.slice(0, resolved.position + 1).reduce((sum, value) => sum + value, 0);
      if (spec.kind === 'running-total') return cumulative;
      const total = resolved.series.reduce((sum, value) => sum + value, 0);
      return total ? cumulative / total : null;
    }
    if (spec.kind === 'rank') {
      const resolved = resolveAxis(context, spec, valueIndex);
      if (!resolved || resolved.base == null) return null;
      const series = resolved.series;
      const ranked = series.filter((value): value is number => value != null).sort((left, right) => spec.direction === 'ascending' ? left - right : right - left);
      const rank = ranked.findIndex((value) => value === resolved.base);
      return rank < 0 ? null : rank + 1;
    }
    if (spec.kind === 'index') return grand != null && rowTotal && columnTotal ? current * grand / rowTotal / columnTotal : null;
    return null;
  };

  for (const context of contexts) {
    for (const [valueIndex, field] of fields.entries()) {
      const spec = field.showAs ?? { kind: 'normal' as const };
      const current = pivotNumericValue(rawValue(context.cell, valueIndex));
      if (current == null || spec.kind === 'normal') continue;
      const grand = pivotNumericValue(grandValues[valueIndex] ?? null);
      const rowTotal = context.kind === 'grand-total'
        ? (grand ?? current)
        : context.node?.values.reduce((sum, cell) => sum + (pivotNumericValue(rawValue(cell, valueIndex)) ?? 0), 0) ?? 0;
      const columnTotal = context.kind === 'grand-total'
        ? (grand ?? current)
        : numericSum(leafContexts, valueIndex, context.columnIndex);
      // Top-level members have the grand total as their parent context. This
      // is the only deterministic parent for a Pivot root member.
      const parentTotal = context.kind === 'grand-total'
        ? grand
        : context.parent ? pivotNumericValue(rawValue(context.parent.values[context.columnIndex], valueIndex)) : grand;
      context.cell.values[valueIndex] = transform(spec, current, grand, rowTotal, columnTotal, parentTotal, context, valueIndex);
    }
  }
}

function computePivotResultFromTable(
  definition: PivotDefinition,
  rawTable: SourceTable,
  controls: readonly PivotTaskControl[],
  revisions: PivotRevisionKey,
  targetBounds: { rowCount: number; columnCount: number },
): PivotResultTree {
  const collator = createPivotCollator(definition.layout.collation);
  const calculatedFieldIds = new Set((definition.layout.calculatedFields ?? []).map((field) => field.fieldId));
  const calculatedItemIds = new Set((definition.layout.calculatedItems ?? []).map((field) => field.fieldId));
  const structuralReferences: string[] = [
    ...definition.layout.rows.map((entry) => entry.fieldId),
    ...definition.layout.columns.map((entry) => entry.fieldId),
    ...definition.layout.filters.flatMap((filter) => {
      if (filter.kind === 'top-items') return [filter.fieldId, valueSourceFieldId(filter.valueId, definition.layout.values)];
      if (filter.kind === 'condition' && filter.valueId !== undefined) return [filter.fieldId, valueSourceFieldId(filter.valueId, definition.layout.values)];
      return [filter.fieldId];
    }),
  ];
  const calculatedStructuralReference = structuralReferences.find((field) => calculatedFieldIds.has(field));
  if (calculatedStructuralReference) throw new Error(`Pivot calculated field is only valid in Values: ${calculatedStructuralReference}`);
  const calculatedItemStructuralReference = structuralReferences.find((field) => calculatedItemIds.has(field));
  if (calculatedItemStructuralReference) throw new Error(`Pivot calculated item is only valid in Values: ${calculatedItemStructuralReference}`);
  const aggregates = new PivotAggregatePlanner();
  const calculatedPlan = createCalculatedFieldPlan(definition.fieldCatalog.fields, definition.layout.calculatedFields);
  const calculatedFields = createCalculatedFieldEvaluator(calculatedPlan, aggregates);
  const resultFields = resultValueFields(definition.layout);
  const rows = applyCalculatedItems(rawTable.rows, definition.fieldCatalog.fields, definition.layout, resultFields);
  const references = [
    ...definition.layout.rows.map((entry) => entry.fieldId),
    ...definition.layout.columns.map((entry) => entry.fieldId),
    ...definition.layout.filters.flatMap((filter) => {
      if (filter.kind === 'top-items') return [filter.fieldId, valueSourceFieldId(filter.valueId, definition.layout.values)];
      if (filter.kind === 'condition' && filter.valueId !== undefined) return [filter.fieldId, valueSourceFieldId(filter.valueId, definition.layout.values)];
      return [filter.fieldId];
    }),
    ...definition.layout.values.map((entry) => entry.fieldId),
  ];
  const known = new Set([...definition.fieldCatalog.fields.map((field) => field.fieldId), ...(definition.layout.calculatedFields ?? []).map((field) => field.fieldId), ...(definition.layout.calculatedItems ?? []).map((field) => field.fieldId)]);
  const unknown = references.find((field) => field && !known.has(field));
  if (unknown && rawTable.fields.length) throw new Error(`Unknown pivot field: ${unknown}`);
  let filtered = matchesControls(rows, controls);
  filtered = filtered.filter((row) => definition.layout.filters.filter((filter) => filter.kind !== 'top-items' && !(filter.kind === 'condition' && filter.family === 'value')).every((filter) => matchesFilter(row, filter, collator, definition)));
  filtered = applyValueFilters(filtered, definition.layout.filters, definition, calculatedFields, collator, aggregates);
  filtered = topItems(filtered, definition.layout.filters, definition.layout.values, calculatedFields, definition, aggregates);
  const columns = definition.layout.columns.length
    ? axisGroups(filtered, definition.layout.columns, definition.fieldCatalog, collator, resultFields, calculatedFields, aggregates)
    : [{ values: [], rows: filtered, rowSet: new Set(filtered) }];
  assertPivotTaskFootprint(definition, filtered, columns, resultFields, collator, calculatedFields, aggregates, targetBounds);
  const grandTotal: PivotResultCell = {
    id: `${definition.id}|grand-total`,
    kind: 'grand-total',
    columnPath: [],
    values: resultFields.map((field) => resultValue(filtered, field, field.summarizeBy, calculatedFields, aggregates)),
    sourceRowPaths: filtered.flatMap((row) => sourceRowPaths(row)),
  };
  const tree: PivotResultTree = {
    schema: PIVOT_RESULT_TREE_SCHEMA,
    pivotId: definition.id,
    fields: definition.fieldCatalog,
    columnPaths: columns.map((column) => column.values),
    valueFields: resultFields,
    rows: resultNodes(filtered, definition.layout.rows, 0, columns, resultFields, definition.layout.subtotalLocation, definition.layout.showRowGrandTotals, definition.fieldCatalog, collator, calculatedFields, aggregates),
    columnGrandTotals: resultCells(filtered, columns, resultFields, [`${definition.id}|grand-total`], 'grand-total', undefined, calculatedFields, aggregates),
    grandTotal,
    sourceRowPaths: filtered.flatMap((row) => sourceRowPaths(row)),
  };
  const slicerItems: Record<string, PivotSlicerItemProjection[]> = {};
  for (const control of controls) {
    if (control.payload.kind !== 'slicer') continue;
    slicerItems[control.drawingId] = slicerItemProjection(definition, rows, control.drawingId, control.payload, collator, calculatedFields, controls, aggregates);
  }
  if (Object.keys(slicerItems).length > 0) tree.slicerItems = slicerItems;
  applyShowAs(tree, resultFields, definition.layout);
  tree.sourceRevision = revisions.sourceRevision;
  tree.layoutRevision = revisions.layoutRevision;
  tree.filterRevision = revisions.filterRevision;
  return tree;
}

export function preparePivotTaskInput(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotTaskEvaluationInput {
  if (pivot.source.kind === 'data-source') throw new Error('Block-backed Pivot tasks require an asynchronously acquired source index');
  const calculator = pivotSourceCalculator(workbook, pivot, formula);
  const definition = normalizePivotDefinitionWithCalculator(workbook, pivot, calculator);
  const rawTable = sourceTable(workbook, definition, definition.fieldCatalog, calculator);
  return {
    definition,
    source: rawTable.index,
    controls: collectPivotTaskControls(workbook, pivot),
    revisions: getPivotRevisionKey(workbook, definition, calculator),
    targetBounds: pivotTargetBounds(workbook, definition),
  };
}

/**
 * Bind a validated block/columnar source to the normal Pivot task contract.
 * Data-source Pivots never fabricate a worksheet range or an empty table;
 * their source index and revision are explicit inputs to the worker.
 */
export function preparePivotTaskInputFromBlockSource(
  workbook: WorkbookModel,
  pivot: PivotModel,
  source: PivotSourceIndex,
  sourceRevision: string | number,
): PivotTaskEvaluationInput {
  if (pivot.source.kind !== 'data-source') throw new Error('Block source calculation requires a data-source Pivot');
  if (typeof sourceRevision === 'number' && !Number.isSafeInteger(sourceRevision)) throw new Error('Block source revision is invalid');
  assertPivotSourceIndex(source);
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  const sourceFieldIds = new Set(source.fields.map((field) => field.fieldId));
  const calculatedFieldIds = new Set((definition.layout.calculatedFields ?? []).map((field) => field.fieldId));
  const missing = definition.fieldCatalog.fields.find((field) => !sourceFieldIds.has(field.fieldId) && !calculatedFieldIds.has(field.fieldId));
  if (missing) throw new Error(`Block source is missing Pivot field ${missing.fieldId}`);
  const revision = String(sourceRevision);
  if (!revision.trim()) throw new Error('Block source revision is required');
  return {
    definition,
    source,
    controls: collectPivotTaskControls(workbook, pivot),
    revisions: { ...getPivotRevisionKey(workbook, definition), sourceRevision: revision },
    targetBounds: pivotTargetBounds(workbook, definition),
  };
}

/** Product source acquisition: bounded main-thread chunks, cancellation, then one transferable index. */
export async function preparePivotTaskInputAsync(
  workbook: WorkbookModel,
  pivot: PivotModel,
  formula: FormulaEngine,
  options: PivotSourceAcquireOptions = {},
): Promise<PivotTaskEvaluationInput> {
  if (pivot.source.kind === 'data-source') throw new Error('Block-backed Pivot tasks require an asynchronously acquired source index');
  const table = await sourceTableAsync(workbook, pivot, pivot.fieldCatalog, formula, options);
  const fieldCatalog = normalizedFieldCatalog(table, pivot.fieldCatalog);
  const definition = normalizePivotDefinitionFromCatalog({ ...pivot, fieldCatalog });
  return {
    definition,
    controls: collectPivotTaskControls(workbook, pivot),
    revisions: getPivotRevisionKey(workbook, definition, formula),
    targetBounds: pivotTargetBounds(workbook, definition),
    source: table.index,
  };
}

export type PivotTaskDescriptor = Omit<PivotTaskEvaluationInput, 'source'>;

/** Build a new layout/filter task without touching source cells or transferred source buffers. */
export function preparePivotTaskDescriptor(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotTaskDescriptor {
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  return {
    definition,
    controls: collectPivotTaskControls(workbook, pivot),
    revisions: getPivotRevisionKey(workbook, definition, formula),
    targetBounds: pivotTargetBounds(workbook, definition),
  };
}

function pivotTargetBounds(workbook: WorkbookModel, definition: PivotDefinition): { rowCount: number; columnCount: number } {
  const target = workbook.sheets.get(definition.target.sheetId);
  return target
    ? { rowCount: target.rowCount, columnCount: target.columnCount }
    : { rowCount: DEFAULT_SHEET_ROW_COUNT, columnCount: DEFAULT_SHEET_COLUMN_COUNT };
}

function computePivotResultUncached(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotResultTree {
  const input = preparePivotTaskInput(workbook, pivot, formula);
  return evaluatePivotTask({ ...input, targetBounds: { rowCount: MAX_SHEET_ROW_COUNT, columnCount: MAX_SHEET_COLUMN_COUNT } });
}

export function computePivotResultFromDefinition(workbook: WorkbookModel, definition: PivotDefinition, formula?: FormulaEngine): PivotResultTree {
  return evaluatePivotTask(preparePivotTaskInput(workbook, definition, formula));
}

export function computePivotResult(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotResultTree {
  return computePivotResultUncached(workbook, pivot, formula);
}

export interface PivotTaskEvaluationInput {
  definition: PivotDefinition;
  source: PivotSourceIndex;
  controls: PivotTaskControl[];
  revisions: PivotRevisionKey;
  targetBounds: { rowCount: number; columnCount: number };
}

/** Pure worker evaluator: no WorkbookModel, renderer, command runtime, or UI state is reachable here. */
export function evaluatePivotTask(input: PivotTaskEvaluationInput): PivotResultTree {
  if (input.definition.id !== input.revisions.pivotId) throw new Error('Pivot task revision identity does not match its definition');
  return computePivotResultFromTable(
    input.definition,
    openSourceTable(input.source),
    input.controls,
    input.revisions,
    input.targetBounds,
  );
}

/** Apply the normal Pivot calculation pipeline to asynchronously loaded block data. */
export function computePivotResultFromBlockSource(
  workbook: WorkbookModel,
  pivot: PivotModel,
  source: PivotSourceTableInput,
  sourceRevision: string,
): PivotResultTree {
  const input = preparePivotTaskInputFromBlockSource(workbook, pivot, source, sourceRevision);
  const cached = blockPivotResultCaches.get(workbook)?.get(pivot.id);
  if (cached && cached.sourceRevision === input.revisions.sourceRevision
    && cached.layoutRevision === input.revisions.layoutRevision
    && cached.filterRevision === input.revisions.filterRevision) return structuredClone(cached.result);
  const result = evaluatePivotTask(input);
  const cache = blockPivotResultCaches.get(workbook) ?? new Map<string, BlockPivotResultCacheEntry>();
  cache.set(pivot.id, { sourceRevision: input.revisions.sourceRevision, layoutRevision: input.revisions.layoutRevision, filterRevision: input.revisions.filterRevision, result: structuredClone(result) });
  if (!blockPivotResultCaches.has(workbook)) blockPivotResultCaches.set(workbook, cache);
  return result;
}

export function getCachedBlockPivotResult(workbook: WorkbookModel, pivotId: string): PivotResultTree | undefined {
  const result = blockPivotResultCaches.get(workbook)?.get(pivotId)?.result;
  return result ? structuredClone(result) : undefined;
}

export function clearBlockPivotResultCache(workbook: WorkbookModel, pivotId?: string): void {
  const cache = blockPivotResultCaches.get(workbook);
  if (!cache) return;
  if (pivotId === undefined) cache.clear();
  else cache.delete(pivotId);
}

function nodeExpanded(node: PivotResultNode, layout: PivotLayout): boolean {
  if (!node.children.length) return false;
  const nodeId = node.nodeId ?? '';
  const expansion = layout.expansion;
  if (!expansion) return true;
  // Expansion state controls traversal, never the existence of the current
  // row.  A collapsed node remains visible while only its descendants are
  // omitted from the projection. Explicit expanded IDs are retained as
  // stable overrides for restored/native Pivot state; the default is open.
  return !expansion.collapsedNodeIds.includes(nodeId) || expansion.expandedNodeIds.includes(nodeId);
}

interface FlatNode {
  node: PivotResultNode;
  labels: string[];
  visible: boolean;
}

function flattenNodes(nodes: PivotResultNode[], layout: PivotLayout, labels: string[] = [], parentVisible = true): FlatNode[] {
  const output: FlatNode[] = [];
  for (const node of nodes) {
    const currentLabels = [...labels, node.label];
    const visible = parentVisible;
    const includeNode = !node.children.length || node.subtotal;
    const children = visible && nodeExpanded(node, layout) ? flattenNodes(node.children, layout, currentLabels, true) : [];
    if (layout.subtotalLocation === 'bottom' && node.subtotal) {
      output.push(...children);
      output.push({ node, labels: currentLabels, visible });
    } else {
      if (includeNode) output.push({ node, labels: currentLabels, visible });
      output.push(...children);
    }
  }
  return output;
}

/**
 * Resolve the row-header projection from the one canonical report layout.
 * The result tree is shared by all layouts; only this presentation boundary
 * decides whether hierarchy is compacted, repeated, or shown as an outline.
 */
function projectionRowLabels(item: FlatNode, layout: PivotLayout, rowHeaderCount: number): string[] {
  if (layout.reportLayout === 'compact') {
    const label = item.labels.filter((entry) => entry.length > 0).join(' / ');
    return [label || item.node.label];
  }
  if (layout.reportLayout === 'tabular') {
    return Array.from({ length: rowHeaderCount }, (_, axis) => item.labels[axis] ?? '');
  }
  // Outline mode deliberately does not repeat an ancestor label on detail
  // rows. Subtotal rows own the label for their field and child rows occupy
  // the following lines, which is the distinction from tabular mode.
  return Array.from({ length: rowHeaderCount }, (_, axis) => axis === item.node.depth ? item.node.label : '');
}

function pivotNodeIds(nodes: readonly PivotResultNode[], target = new Set<string>()): Set<string> {
  for (const node of nodes) {
    if (node.nodeId) target.add(node.nodeId);
    pivotNodeIds(node.children, target);
  }
  return target;
}

function normalizeExpansionForTree(expansion: PivotLayout['expansion'], tree: PivotResultTree): NonNullable<PivotLayout['expansion']> {
  const known = pivotNodeIds(tree.rows);
  const source = expansion ?? { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true };
  const dedupeKnown = (ids: readonly string[]) => [...new Set(ids.filter((id) => known.has(id)))];
  return {
    expandedNodeIds: dedupeKnown(source.expandedNodeIds),
    collapsedNodeIds: dedupeKnown(source.collapsedNodeIds),
    showButtons: source.showButtons,
  };
}

function textForValue(value: PivotScalar, options = DEFAULT_PIVOT_DISPLAY_OPTIONS, numberFormat?: string): string {
  if (isPivotError(value)) return options.showErrorValues ? (options.errorCellText || value.code) : '';
  if (value == null || value === '') return options.fillEmptyCells ? options.emptyCellText : '';
  return numberFormat ? formatPivotValue(value, numberFormat) : display(value);
}

function formatPivotValue(value: Exclude<PivotScalar, null | PivotErrorValue>, numberFormat: string): string {
  if (typeof value === 'number' || typeof value === 'string' || typeof value === 'boolean') {
    return formatNumberValue(value, numberFormat);
  }
  return display(value);
}

function projectionCell(pivotId: string, row: number, column: number, kind: PivotProjectionCell['kind'], value: PivotScalar, text: string, extra: Partial<PivotProjectionCell> = {}): PivotProjectionCell {
  return { id: `${pivotId}|r${row}|c${column}`, pivotId, row, column, kind, value, text, ...extra };
}

function projectionRange(target: PivotTarget, rowCount: number, columnCount: number): RangeRef {
  return { sheetId: target.sheetId, startRow: target.anchor.row, endRow: target.anchor.row + Math.max(rowCount - 1, 0), startColumn: target.anchor.column, endColumn: target.anchor.column + Math.max(columnCount - 1, 0) };
}

interface PivotFootprint {
  pivotId: string;
  range: RangeRef;
  source: 'current' | 'last-valid';
}

function occupiedRangeForDefinition(definition: PivotDefinition, tree?: PivotResultTree): RangeRef {
  const displayOptions = normalizePivotDisplayOptions(definition.presentation?.displayOptions);
  const rowHeaderCount = definition.layout.reportLayout === 'compact' ? 1 : Math.max(definition.layout.rows.length, 1);
  const values = tree?.valueFields ?? definition.layout.values.map((field) => ({ ...field, sourceFieldId: field.fieldId }));
  const columnPathCount = Math.max(tree?.columnPaths.length ?? 0, 1);
  const valueColumnCount = Math.max(columnPathCount * Math.max(values.length, 1) + (definition.layout.showRowGrandTotals ? Math.max(values.length, 1) : 0), 1);
  let row = 1;
  const reportFilterFields = displayOptions.showFieldHeaders
    ? [...new Set(definition.layout.filters.filter((entry) => entry.scope !== 'field').map((entry) => entry.fieldId))]
    : [];
  row += reportFilterFields.length;
  if (displayOptions.showFieldHeaders) row += 1;
  if (tree) {
    const flat = flattenNodes(tree.rows, { ...definition.layout, expansion: normalizeExpansionForTree(definition.layout.expansion, tree) });
    row += flat.filter((item) => item.visible).length;
    if (tree.grandTotal && definition.layout.showColumnGrandTotals) row += 1;
  } else {
    row += 1;
  }
  return projectionRange(definition.target, Math.max(row, 1), Math.max(valueColumnCount + rowHeaderCount, 1));
}

/** Resolve the complete derived footprint from an already canonical definition. */
export function getPivotOccupiedRange(definition: PivotDefinition, tree?: PivotResultTree): RangeRef {
  return occupiedRangeForDefinition(definition, tree);
}

function resolvePivotFootprint(workbook: WorkbookModel, pivot: PivotModel): PivotFootprint | undefined {
  const last = lastValidPivotProjections.get(workbook)?.get(pivot.id);
  const targetMatches = last
    && last.projection.target.sheetId === pivot.target.sheetId
    && last.projection.target.anchor.row === pivot.target.anchor.row
    && last.projection.target.anchor.column === pivot.target.anchor.column;
  if (last && targetMatches && pivotResultMatchesRevision(workbook, pivot, last.result)) {
    return { pivotId: pivot.id, range: structuredClone(last.projection.occupiedRange), source: 'current' };
  }
  if (last && targetMatches && pivot.refreshPolicy.mode === 'manual' && pivotResultMatchesLayoutAndFilter(workbook, pivot, last.result)) {
    return { pivotId: pivot.id, range: structuredClone(last.projection.occupiedRange), source: 'last-valid' };
  }
  if (pivot.source.kind === 'data-source') return last ? { pivotId: pivot.id, range: structuredClone(last.projection.occupiedRange), source: 'last-valid' } : undefined;
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  return { pivotId: pivot.id, range: occupiedRangeForDefinition(definition), source: 'current' };
}

function cellMetadataRange(sheetId: string, key: string): RangeRef | undefined {
  const [row, column] = key.split(':').map(Number);
  if (typeof row !== 'number' || typeof column !== 'number' || !Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || column < 0) return undefined;
  return { sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column };
}

function drawingAnchorRange(sheetId: string, anchor: { kind: string; row?: number; column?: number; endRow?: number; endColumn?: number }): RangeRef | undefined {
  const row = anchor.row;
  const column = anchor.column;
  if (typeof row !== 'number' || typeof column !== 'number' || !Number.isSafeInteger(row) || !Number.isSafeInteger(column)) return undefined;
  const endRow = anchor.endRow;
  const endColumn = anchor.endColumn;
  return {
    sheetId,
    startRow: row,
    endRow: typeof endRow === 'number' && Number.isSafeInteger(endRow) ? Math.max(row, endRow) : row,
    startColumn: column,
    endColumn: typeof endColumn === 'number' && Number.isSafeInteger(endColumn) ? Math.max(column, endColumn) : column,
  };
}

function drawingBelongsToPivot(workbook: WorkbookModel, drawing: WorksheetModel['drawings'][number], pivotId: string): boolean {
  const payload = workbook.getSheet(drawing.sheetId).drawingPayloads.get(drawing.payloadId);
  if (!payload || !('pivotId' in payload)) return false;
  if (payload.pivotId === pivotId) return true;
  if (payload.kind !== 'slicer' && payload.kind !== 'timeline') return false;
  return payload.connections?.some((connection) => connection.pivotId === pivotId) === true;
}

export function detectPivotCollision(workbook: WorkbookModel, pivot: PivotModel, range: RangeRef): import('@react-sheets/core-model').PivotCollision {
  const sheet = workbook.getSheet(range.sheetId);
  const reasons = new Set<import('@react-sheets/core-model').PivotCollisionReason>();
  const conflictingRanges: RangeRef[] = [];
  const conflicts: import('@react-sheets/core-model').PivotCollisionConflict[] = [];
  const addConflict = (reason: import('@react-sheets/core-model').PivotCollisionReason, conflictRange: RangeRef, participantId?: string): void => {
    const normalized = structuredClone(conflictRange);
    reasons.add(reason);
    if (!conflictingRanges.some((existing) => existing.sheetId === normalized.sheetId && existing.startRow === normalized.startRow && existing.endRow === normalized.endRow && existing.startColumn === normalized.startColumn && existing.endColumn === normalized.endColumn)) {
      conflictingRanges.push(normalized);
    }
    if (!conflicts.some((existing) => existing.reason === reason && existing.participantId === participantId && existing.range.sheetId === normalized.sheetId && existing.range.startRow === normalized.startRow && existing.range.endRow === normalized.endRow && existing.range.startColumn === normalized.startColumn && existing.range.endColumn === normalized.endColumn)) {
      conflicts.push({ reason, range: normalized, ...(participantId ? { participantId } : {}) });
    }
  };
  const wholeSheet = { sheetId: sheet.id, startRow: 0, endRow: Math.max(sheet.rowCount - 1, 0), startColumn: 0, endColumn: Math.max(sheet.columnCount - 1, 0) };
  if (range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) addConflict('worksheet-bounds', range);
  sheet.cells.forEach((_cell, row, column) => {
    if (row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn) {
      addConflict('cell-data', { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column });
    }
  });
  for (const merge of sheet.merges) {
    if (rangesIntersect(range, merge.range)) addConflict('merge', merge.range);
  }
  for (const table of sheet.sheetTables) if (rangesIntersect(range, table.range)) addConflict('sheet-table', table.range, table.id);
  for (const region of sheet.dataRegions) if (rangesIntersect(range, region.range)) addConflict('data-region', region.range, region.id);
  for (const rule of sheet.conditionalFormats) for (const candidate of rule.ranges) if (rangesIntersect(range, candidate)) addConflict('conditional-format', candidate, rule.id);
  for (const rule of sheet.dataValidations) for (const candidate of rule.ranges) if (rangesIntersect(range, candidate)) addConflict('data-validation', candidate, rule.id);
  if (sheet.autoFilter && rangesIntersect(range, sheet.autoFilter.range)) addConflict('auto-filter', sheet.autoFilter.range);
  if (sheet.bandedRule && rangesIntersect(range, sheet.bandedRule.range)) addConflict('banded-rule', sheet.bandedRule.range);
  for (const spill of sheet.spillRanges) if (rangesIntersect(range, spill.range)) addConflict('spill', spill.range);
  for (const sparkline of sheet.sparklines) {
    const candidate = { sheetId: sheet.id, startRow: sparkline.anchor.row, endRow: sparkline.anchor.row, startColumn: sparkline.anchor.column, endColumn: sparkline.anchor.column };
    if (rangesIntersect(range, candidate)) addConflict('sparkline', candidate, sparkline.id);
  }
  for (const drawing of sheet.drawings) {
    if (drawingBelongsToPivot(workbook, drawing, pivot.id)) continue;
    const candidate = drawingAnchorRange(sheet.id, drawing.anchor);
    // Absolute floating drawings have no worksheet-cell occupancy. Only a
    // verifiable one-cell/two-cell anchor participates in structural overlap;
    // never invent a whole-sheet range from pixel-only geometry.
    if (candidate && rangesIntersect(range, candidate)) {
      addConflict('drawing', candidate, drawing.id);
    }
  }
  for (const { key } of sheet.review.noteEntries()) {
    const candidate = cellMetadataRange(sheet.id, key);
    if (candidate && rangesIntersect(range, candidate)) addConflict('note', candidate, key);
  }
  for (const [key] of sheet.hyperlinks) {
    const candidate = cellMetadataRange(sheet.id, key);
    if (candidate && rangesIntersect(range, candidate)) addConflict('hyperlink', candidate, key);
  }
  for (const comment of sheet.review.threadEntries()) {
    const candidate = { sheetId: sheet.id, startRow: comment.row, endRow: comment.row, startColumn: comment.column, endColumn: comment.column };
    if (rangesIntersect(range, candidate)) addConflict('comment', candidate, comment.id);
  }
  for (const rule of sheet.protectionRules) {
    if (!rule.locked) continue;
    const candidate = rule.range ?? wholeSheet;
    if (rangesIntersect(range, candidate)) addConflict('protection', candidate, rule.id);
  }
  for (const ownerSheet of workbook.getSheets()) {
    for (const candidate of ownerSheet.pivots) {
      if (candidate.id === pivot.id || getPivotTarget(candidate).sheetId !== range.sheetId) continue;
      const footprint = resolvePivotFootprint(workbook, candidate);
      if (!footprint) {
        addConflict('unresolved-pivot', wholeSheet, candidate.id);
      } else if (rangesIntersect(range, footprint.range)) {
        addConflict('pivot', footprint.range, candidate.id);
      }
    }
  }
  return { status: reasons.size ? 'collision' : 'clear', reasons: [...reasons], conflictingRanges, conflicts };
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId && left.startRow <= right.endRow && right.startRow <= left.endRow && left.startColumn <= right.endColumn && right.startColumn <= left.endColumn;
}

function refreshState(workbook: WorkbookModel, pivot: PivotModel, collision: import('@react-sheets/core-model').PivotCollision, status: PivotRefreshState['status'] = 'ready', error?: string, formula?: FormulaEngine): PivotRefreshState {
  const revisions = getPivotRevisionKey(workbook, pivot, formula);
  return {
    status: collision.status === 'collision' ? 'collision' : status,
    revision: Number.parseInt(revisions.sourceRevision.slice(-6), 16) || 0,
    sourceRevision: revisions.sourceRevision,
    completedAt: new Date().toISOString(),
    ...(error ? { error } : {}),
  };
}

export function getPivotRefreshState(workbook: WorkbookModel, pivot: PivotModel, collision?: import('@react-sheets/core-model').PivotCollision, status: PivotRefreshState['status'] = 'ready', error?: string): PivotRefreshState {
  const effectiveCollision = collision ?? { status: 'clear' as const, reasons: [], conflictingRanges: [], conflicts: [] };
  return refreshState(workbook, pivot, effectiveCollision, status, error);
}

/** Build one candidate worksheet overlay. It returns cells only; no workbook cell is mutated. */
function buildPivotGridProjectionCandidate(
  workbook: WorkbookModel,
  pivot: PivotModel,
  cachedResult?: PivotResultTree,
  options: PivotProjectionOptions = {},
): PivotGridProjection {
  const definition = options.canonicalDefinition ?? normalizePivotDefinitionFromCatalog(pivot);
  const target = definition.target;
  const displayOptions = normalizePivotDisplayOptions(definition.presentation?.displayOptions);
  let tree: PivotResultTree | undefined = cachedResult;
  let error: string | undefined;
  let loading = false;
  const sourceState = options.sourceState;
  if (!tree) {
    if (sourceState?.availability === 'error' || sourceState?.availability === 'missing') {
      error = sourceState.error ?? `PivotTable source ${sourceState.availability}`;
    } else {
      loading = true;
    }
  }
  if (tree && (sourceState?.availability === 'error' || sourceState?.availability === 'missing')) {
    error = sourceState.error ?? `PivotTable source ${sourceState.availability}`;
  } else if (tree && sourceState?.availability === 'loading') {
    loading = true;
  }
  const cells: PivotProjectionCell[] = [];
  const rowHeaderCount = definition.layout.reportLayout === 'compact' ? 1 : Math.max(definition.layout.rows.length, 1);
  const values = tree?.valueFields ?? definition.layout.values.map((field) => ({ ...field, sourceFieldId: field.fieldId }));
  const columnPathCount = Math.max(tree?.columnPaths.length ?? 0, 1);
  const valueColumnCount = Math.max(columnPathCount * Math.max(values.length, 1) + (definition.layout.showRowGrandTotals ? Math.max(values.length, 1) : 0), 1);
  let row = 0;
  cells.push(projectionCell(definition.id, row, 0, 'title', definition.id, definition.id));
  row += 1;
  const reportFilterFields = displayOptions.showFieldHeaders
    ? [...new Set(definition.layout.filters.filter((entry) => entry.scope !== 'field').map((entry) => entry.fieldId))]
    : [];
  for (const fieldId of reportFilterFields) {
    const filterSummary = summarizePivotReportFilters(definition.layout.filters, definition.fieldCatalog, fieldId, definition.layout.values);
    // The semantic summary is rendered by the presentation layer.  Keep the
    // projection text stable and locale-independent for export/replay.
    cells.push(projectionCell(definition.id, row, 0, 'filter', fieldId, filterSummary.fieldName, { fieldId, filterSummary }));
    row += 1;
  }
  if (displayOptions.showFieldHeaders) {
    for (let index = 0; index < rowHeaderCount; index += 1) {
      const fieldId = definition.layout.rows[index]?.fieldId ?? definition.layout.rows[0]?.fieldId;
      const label = definition.layout.reportLayout === 'compact'
        ? 'Row Labels'
        : fieldId ? fieldName(fieldId, definition.fieldCatalog) : 'Row Labels';
      cells.push(projectionCell(definition.id, row, index, 'column-header', null, label, { ...(index === 0 ? { captionKey: 'row-labels' as const } : {}), fieldId }));
    }
  }
  const columnPaths = tree?.columnPaths ?? [];
  for (let columnIndex = 0; columnIndex < columnPathCount; columnIndex += 1) {
    const path = columnPaths[columnIndex] ?? [];
    for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
      const column = rowHeaderCount + columnIndex * Math.max(values.length, 1) + valueIndex;
      const valueField = values[valueIndex];
      const valueCaption = valueField ? (valueField.displayName ?? fieldName(valueField.fieldId, definition.fieldCatalog)) : '';
      const label = path.length ? `${path.map(display).join(' / ')} ${valueCaption}`.trim() : valueCaption;
      if (displayOptions.showFieldHeaders) cells.push(projectionCell(definition.id, row, column, 'column-header', path[0] ?? null, label, { columnPath: path, fieldId: definition.layout.columns[definition.layout.columns.length - 1]?.fieldId, valueId: valueField?.valueId, isLastColumn: !definition.layout.showRowGrandTotals && columnIndex === columnPathCount - 1 }));
    }
  }
  if (definition.layout.showRowGrandTotals && displayOptions.showFieldHeaders) {
    for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
      const valueField = values[valueIndex];
      const column = rowHeaderCount + columnPathCount * Math.max(values.length, 1) + valueIndex;
      cells.push(projectionCell(definition.id, row, column, 'column-header', null, valueField ? `Grand Total ${valueField.displayName ?? fieldName(valueField.fieldId, definition.fieldCatalog)}` : 'Grand Total', { captionKey: 'grand-total', valueId: valueField?.valueId, isLastColumn: valueIndex === Math.max(values.length, 1) - 1 }));
    }
  }
  if (displayOptions.showFieldHeaders) row += 1;
  if (tree) {
    const expansion = normalizeExpansionForTree(definition.layout.expansion, tree);
    const projectionLayout: PivotLayout = { ...definition.layout, expansion };
    const flat = flattenNodes(tree.rows, projectionLayout);
    for (const item of flat) {
      if (!item.visible) continue;
      const node = item.node;
      const labels = projectionRowLabels(item, definition.layout, rowHeaderCount);
      for (let axis = 0; axis < rowHeaderCount; axis += 1) {
        const label = labels[axis] ?? '';
        const kind: PivotProjectionCell['kind'] = axis === 0 && node.children.length && expansion.showButtons ? 'expand-toggle' : node.subtotal ? 'subtotal' : 'row-header';
        cells.push(projectionCell(definition.id, row, axis, kind, axis === 0 ? node.key : null, label, { nodeId: node.nodeId, fieldId: definition.layout.rows[axis]?.fieldId ?? definition.layout.rows[0]?.fieldId, expandable: node.children.length > 0, expanded: nodeExpanded(node, projectionLayout) }));
      }
      for (let columnIndex = 0; columnIndex < columnPathCount; columnIndex += 1) {
        const resultCell = node.values[columnIndex];
        for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
          const column = rowHeaderCount + columnIndex * Math.max(values.length, 1) + valueIndex;
          const value = resultCell?.values[valueIndex] ?? null;
          const valueField = values[valueIndex];
          cells.push(projectionCell(definition.id, row, column, node.subtotal ? 'subtotal' : 'value', value, textForValue(value, displayOptions, valueField?.numberFormat), { nodeId: node.nodeId, resultCellId: resultCell?.id, columnPath: resultCell?.columnPath, valueId: valueField?.valueId, sourceRowPaths: resultCell?.sourceRowPaths, isLastColumn: !definition.layout.showRowGrandTotals && columnIndex === columnPathCount - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
        }
      }
      if (definition.layout.showRowGrandTotals) {
        const resultCell = node.rowGrandTotal;
        for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
          const column = rowHeaderCount + columnPathCount * Math.max(values.length, 1) + valueIndex;
          const value = resultCell?.values[valueIndex] ?? null;
          const valueField = values[valueIndex];
          cells.push(projectionCell(definition.id, row, column, 'grand-total', value, textForValue(value, displayOptions, valueField?.numberFormat), { nodeId: node.nodeId, resultCellId: resultCell?.id, columnPath: resultCell?.columnPath, valueId: valueField?.valueId, sourceRowPaths: resultCell?.sourceRowPaths, isLastColumn: valueIndex === Math.max(values.length, 1) - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
        }
      }
      row += 1;
    }
    if (tree.grandTotal && definition.layout.showColumnGrandTotals) {
      cells.push(projectionCell(definition.id, row, 0, 'grand-total', null, 'Grand Total', { captionKey: 'grand-total', resultCellId: tree.grandTotal.id, sourceRowPaths: tree.grandTotal.sourceRowPaths }));
      const columnGrandTotals = tree.columnGrandTotals ?? (tree.grandTotal ? [tree.grandTotal] : []);
      columnGrandTotals.forEach((resultCell, columnIndex) => resultCell.values.forEach((value, valueIndex) => {
        const column = rowHeaderCount + columnIndex * Math.max(values.length, 1) + valueIndex;
        const valueField = values[valueIndex];
        cells.push(projectionCell(definition.id, row, column, 'grand-total', value, textForValue(value, displayOptions, valueField?.numberFormat), { resultCellId: resultCell.id, columnPath: resultCell.columnPath, valueId: valueField?.valueId, sourceRowPaths: resultCell.sourceRowPaths, isLastColumn: !definition.layout.showRowGrandTotals && columnIndex === columnGrandTotals.length - 1 && valueIndex === resultCell.values.length - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
      }));
      if (definition.layout.showRowGrandTotals) {
        tree.grandTotal.values.forEach((value, valueIndex) => {
          const column = rowHeaderCount + columnPathCount * Math.max(values.length, 1) + valueIndex;
          const valueField = values[valueIndex];
          cells.push(projectionCell(definition.id, row, column, 'grand-total', value, textForValue(value, displayOptions, valueField?.numberFormat), { resultCellId: tree.grandTotal?.id, valueId: valueField?.valueId, sourceRowPaths: tree.grandTotal?.sourceRowPaths, isLastColumn: valueIndex === tree.grandTotal!.values.length - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
        });
      }
      row += 1;
    }
  } else {
    cells.push(projectionCell(definition.id, row, 0, error ? 'error' : 'loading', null, error ?? 'Loading PivotTable', error ? {} : { captionKey: 'loading' }));
    row += 1;
  }
  const occupiedRange = occupiedRangeForDefinition(definition, tree);
  const collision = detectPivotCollision(workbook, pivot, occupiedRange);
  return {
    schema: PIVOT_GRID_PROJECTION_SCHEMA,
    pivotId: definition.id,
    sheetId: target.sheetId,
    target,
    presentation: structuredClone(definition.presentation),
    occupiedRange,
    cells,
    collision,
    refresh: refreshState(workbook, pivot, collision, error ? 'error' : loading ? 'refreshing' : tree ? 'ready' : 'refreshing', error, options.formula),
  };
}

function projectionWithStatus(
  workbook: WorkbookModel,
  pivot: PivotModel,
  entry: LastValidPivotProjection,
  collision: import('@react-sheets/core-model').PivotCollision,
  status: PivotRefreshState['status'],
  error?: string,
  formula?: FormulaEngine,
): PivotGridProjection {
  const projection = structuredClone(entry.projection);
  projection.collision = structuredClone(collision);
  projection.refresh = refreshState(workbook, pivot, collision, status, error, formula);
  return projection;
}

/**
 * Build the production projection with a last-valid guard. A collision or
 * asynchronous source failure never replaces a successful result with an
 * empty/error grid, and ordinary worksheet cells remain untouched.
 */
export function buildPivotGridProjection(
  workbook: WorkbookModel,
  pivot: PivotModel,
  cachedResult?: PivotResultTree,
  options: PivotProjectionOptions = {},
): PivotGridProjection {
  const revision = getPivotRevisionKey(workbook, pivot, options.formula);
  const sourceRevisionMismatch = pivot.source.kind === 'data-source'
    && options.sourceState?.sourceRevision !== undefined
    && cachedResult !== undefined
    && cachedResult.sourceRevision !== String(options.sourceState.sourceRevision);
  const blockResultReady = pivot.source.kind === 'data-source'
    && options.sourceState?.availability === 'ready'
    && !sourceRevisionMismatch
    && pivotResultMatchesLayoutAndFilter(workbook, pivot, cachedResult, options.formula);
  const staleResult = pivotResultMatchesLayoutAndFilter(workbook, pivot, cachedResult, options.formula)
    && !sourceRevisionMismatch
    && cachedResult.sourceRevision !== revision.sourceRevision;
  let effectiveResult = pivotResultMatchesRevision(workbook, pivot, cachedResult, options.formula) || staleResult || blockResultReady ? cachedResult : undefined;
  const candidate = buildPivotGridProjectionCandidate(workbook, pivot, effectiveResult, options);
  const cache = lastValidPivotProjections.get(workbook);
  const last = cache?.get(pivot.id);
  const candidateTree = effectiveResult;

  if (options.refreshError && last && candidate.collision.status === 'clear') {
    return projectionWithStatus(workbook, pivot, last, candidate.collision, 'error', options.refreshError, options.formula);
  }

  if (staleResult && !blockResultReady && candidate.collision.status === 'clear') {
    candidate.refresh = refreshState(workbook, pivot, candidate.collision, 'stale', undefined, options.formula);
    return candidate;
  }

  if (candidate.collision.status === 'clear' && candidateTree && candidate.refresh.status === 'ready') {
    const nextCache = cache ?? new Map<string, LastValidPivotProjection>();
    const current = nextCache.get(pivot.id);
    const currentMatches = current
      && current.result.sourceRevision === candidateTree.sourceRevision
      && current.result.layoutRevision === candidateTree.layoutRevision
      && current.result.filterRevision === candidateTree.filterRevision;
    if (!currentMatches) nextCache.set(pivot.id, { projection: structuredClone(candidate), result: structuredClone(candidateTree) });
    if (!cache) lastValidPivotProjections.set(workbook, nextCache);
    return candidate;
  }

  if (last && candidate.collision.status === 'collision') {
    return projectionWithStatus(
      workbook,
      pivot,
      last,
      candidate.collision,
      'collision',
      `Pivot target collision: ${candidate.collision.reasons.join(', ')}`,
      options.formula,
    );
  }

  if (last && (candidate.refresh.status === 'error' || candidate.refresh.status === 'refreshing')) {
    const retainedCollision = detectPivotCollision(workbook, pivot, last.projection.occupiedRange);
    if (retainedCollision.status === 'collision') {
      return projectionWithStatus(
        workbook,
        pivot,
        last,
        retainedCollision,
        'collision',
        `Pivot target collision: ${retainedCollision.reasons.join(', ')}`,
        options.formula,
      );
    }
    return projectionWithStatus(workbook, pivot, last, retainedCollision, candidate.refresh.status, candidate.refresh.error, options.formula);
  }

  return candidate;
}

export function hitTestPivotProjection(projection: PivotGridProjection, row: number, column: number): PivotHitTest {
  const cell = findPivotProjectionCellAt(projection, row, column);
  if (!cell) return { kind: 'none', pivotId: projection.pivotId, row, column };
  return {
    kind: cell.kind === 'expand-toggle' ? 'expand-toggle' : cell.kind === 'filter' ? 'filter' : cell.kind.includes('header') ? 'header' : 'cell',
    pivotId: projection.pivotId,
    cellId: cell.id,
    row,
    column,
    nodeId: cell.nodeId,
    sourceRowPaths: cell.sourceRowPaths,
  };
}

/**
 * Pivot projection cells are emitted in row-major order. Canvas rendering and
 * pointer hit-testing must query that canonical order logarithmically instead
 * of rescanning the complete derived grid for every visible worksheet cell.
 */
export function findPivotProjectionCellAt(projection: PivotGridProjection, row: number, column: number): PivotProjectionCell | undefined {
  let low = 0;
  let high = projection.cells.length - 1;
  while (low <= high) {
    const middle = (low + high) >>> 1;
    const candidate = projection.cells[middle]!;
    const order = candidate.row - row || candidate.column - column;
    if (order === 0) return candidate;
    if (order < 0) low = middle + 1;
    else high = middle - 1;
  }
  return undefined;
}

export function resolvePivotContextHit(projection: PivotGridProjection, row: number, column: number): ContextHit {
  return { ...hitTestPivotProjection(projection, row, column), context: 'pivot', priority: 30 };
}

export function computePivotTable(pivot: PivotModel, tree: PivotResultTree): PivotResultTable {
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  const rows = tree.rows.map((node) => ({ keys: [node.label], values: node.values.flatMap((cell) => cell.values) }));
  const values = tree.valueFields ?? definition.layout.values.map((field) => ({ ...field, sourceFieldId: field.fieldId }));
  const headers = [
    ...definition.layout.rows.map((field) => fieldName(field.fieldId, definition.fieldCatalog)),
    ...tree.columnPaths.flatMap((path) => values.map((field) => path.length ? `${path.map(display).join(' / ')} ${field.displayName ?? fieldName(field.sourceFieldId, definition.fieldCatalog)}` : field.displayName ?? fieldName(field.sourceFieldId, definition.fieldCatalog))),
  ];
  return { headers, rows, grandTotal: tree.grandTotal?.values ?? [], tree };
}

function pivotSourceRangesForExport(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): RangeRef[] {
  return sourceRanges(workbook, pivot, formula);
}

export function getPivotSourceRanges(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): RangeRef[] {
  return structuredClone(pivotSourceRangesForExport(workbook, pivot, formula));
}
