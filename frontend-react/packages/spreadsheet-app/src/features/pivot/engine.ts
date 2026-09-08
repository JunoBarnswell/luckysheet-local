import type {
  PivotDefinition,
  PivotFieldCatalog,
  PivotFieldDataType,
  PivotErrorValue,
  PivotFieldPlacement,
  PivotFilter,
  PivotGroup,
  PivotDateGroupUnit,
  PivotGridProjection,
  PivotHitTest,
  PivotLayout,
  PivotMemberKey,
  PivotModel,
  PivotProjectionCell,
  PivotRefreshState,
  PivotReportFilterSummary,
  PivotReportFilterSummaryEntry,
  PivotResultNode,
  PivotResultTree,
  PivotScalar,
  PivotSort,
  PivotShowAsBaseItem,
  PivotTopBottomMode,
  PivotSource,
  PivotTarget,
  PivotSlicerDrawingPayload,
  PivotTimelineDrawingPayload,
  PivotValueField,
  ContextHit,
  RangeRef,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import {
  PIVOT_GRID_PROJECTION_SCHEMA,
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
  pivotMemberKey,
  normalizePivotRefreshPolicy,
  normalizePivotDisplayOptions,
  normalizePivotNumberFormat,
  pivotNumericValue,
  PIVOT_MAX_MEMBER_COUNT,
  pivotMemberKeyEquals,
  pivotScalarFromMemberKey,
} from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import { readFormulaSpillPages } from '../../formula-spill-sync';

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

/**
 * Render state is ephemeral and belongs to a workbook session. It is not part
 * of PivotDefinition, WorkbookSnapshot, or collaborative operations. A
 * collision/load failure must never destroy the last successful projection.
 */
const lastValidPivotProjections = new WeakMap<WorkbookModel, Map<string, LastValidPivotProjection>>();

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
  const spills = formula ? ranges.map((range) => [...readFormulaSpillPages(formula, range.sheetId, { limit: 256 })]
    .flatMap((page) => page.spills)
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
  return [resolveNamedRange(workbook, source.name, source.sheetId, formula)];
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
    const spill = [...readFormulaSpillPages(calculator, sheet.id, { limit: 256 })]
      .flatMap((page) => page.spills)
      .find((candidate) => candidate.anchor.row === startRow && candidate.anchor.column === startColumn);
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

function scalarCellValue(value: unknown): PivotScalar | null {
  if (isPivotError(value)) return value;
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' ? value : null;
}

function inferFieldDataType(values: readonly PivotScalar[]): PivotFieldDataType {
  const present = values.filter((value) => value !== null && value !== '');
  if (!present.length) return 'mixed';
  if (present.every(isPivotError)) return 'error';
  if (present.some(isPivotError)) return 'mixed';
  if (present.every((value) => typeof value === 'boolean')) return 'boolean';
  if (present.every((value) => typeof value === 'number' && Number.isFinite(value))) return 'number';
  if (present.every((value) => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}(?:[T ].*)?$/.test(value) && !Number.isNaN(Date.parse(value)))) return 'date';
  if (present.every((value) => typeof value === 'string')) return 'text';
  return 'mixed';
}

/**
 * Build only field metadata for a new worksheet Pivot. This bounded metadata
 * read supports the create/recommend flow; it never constructs a source row
 * index and never feeds an analytics calculation. Existing definitions keep
 * their persisted catalog, while the server prepare phase owns execution.
 */
export function getPivotFieldCatalog(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotFieldCatalog {
  if (pivot.fieldCatalog.fields.length) return structuredClone(pivot.fieldCatalog);
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') {
    const manifest = workbook.getDataSource(source.dataSourceId);
    return {
      schema: 'PivotFieldCatalog',
      fields: manifest.fields.map((field) => ({ fieldId: field.id, name: field.name, dataType: field.type, ordinal: field.ordinal, values: [] })),
    };
  }
  const ranges = sourceRanges(workbook, pivot, formula);
  const fields: PivotFieldCatalog['fields'] = [];
  for (let rangeIndex = 0; rangeIndex < ranges.length; rangeIndex += 1) {
    const range = ranges[rangeIndex]!;
    const sheet = workbook.getSheet(range.sheetId);
    const tableFields = source.kind === 'table' ? resolvePivotTable(workbook, source.tableId).fields : [];
    for (let offset = 0; offset <= range.endColumn - range.startColumn; offset += 1) {
      const column = range.startColumn + offset;
      const header = sheet.cells.get(range.startRow, column);
      const name = typeof header?.value === 'string' && header.value.trim() ? header.value : tableFields[offset]?.name ?? `Column ${offset + 1}`;
      const values: PivotScalar[] = [];
      for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
        const cell = sheet.cells.get(row, column);
        const raw = formula && cell?.formula !== undefined ? formula.getCellValue({ sheetId: sheet.id, row, column }) : cell?.formulaValue ?? cell?.value ?? null;
        values.push(scalarCellValue(raw));
      }
      fields.push({ fieldId: sourceIdentity(source, range, offset, rangeIndex), name, dataType: inferFieldDataType(values), ordinal: fields.length, values: canonicalPivotMembers(values) });
    }
  }
  return { schema: 'PivotFieldCatalog', fields };
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

/** Canonicalize a command/task definition from its already validated revision-owned field catalog. */
export function normalizePivotDefinitionFromCatalog(pivot: PivotModel): PivotDefinition {
  const source = getPivotSource(pivot);
  if ((pivot.layout.calculatedFields?.length ?? 0) > 0 || (pivot.layout.calculatedItems?.length ?? 0) > 0) {
    throw new Error('UNSUPPORTED_FEATURE: Pivot calculated fields and items are owned by Rust analytics');
  }
  const fieldCatalog = structuredClone(pivot.fieldCatalog);
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
  // Definition normalization is metadata-only. Source acquisition and
  // calculation belong to the revision-pinned analytics task port.
  return normalizePivotDefinitionFromCatalog(pivot);
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

/**
 * Metadata-only task declaration. Source registration and every analytics
 * phase are owned by ServerPivotTaskPort; this descriptor never contains
 * worksheet rows, a source index, or a client evaluator.
 */
export interface PivotTaskDescriptor {
  definition: PivotDefinition;
  controls: PivotTaskControl[];
  revisions: PivotRevisionKey;
  targetBounds: { rowCount: number; columnCount: number };
}

export function preparePivotTaskDescriptor(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotTaskDescriptor {
  const definition = normalizePivotDefinitionFromCatalog(pivot);
  return {
    definition,
    controls: collectPivotTaskControls(workbook, pivot),
    revisions: getPivotRevisionKey(workbook, pivot, formula),
    targetBounds: pivotTargetBounds(workbook, definition),
  };
}

function pivotTargetBounds(workbook: WorkbookModel, definition: PivotDefinition): { rowCount: number; columnCount: number } {
  const target = workbook.sheets.get(definition.target.sheetId);
  return target
    ? { rowCount: target.rowCount, columnCount: target.columnCount }
    : { rowCount: DEFAULT_SHEET_ROW_COUNT, columnCount: DEFAULT_SHEET_COLUMN_COUNT };
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

function pivotSourceRangesForExport(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): RangeRef[] {
  return sourceRanges(workbook, pivot, formula);
}

export function getPivotSourceRanges(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): RangeRef[] {
  return structuredClone(pivotSourceRangesForExport(workbook, pivot, formula));
}
