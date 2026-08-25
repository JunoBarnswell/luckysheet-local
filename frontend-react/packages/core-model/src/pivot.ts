import type { Column, RangeRef, Row, SheetId } from './index';
import type { FormulaErrorCode, PivotTimelineLevel } from './domain';

/**
 * Pivot is a derived view over workbook data. The definition below is the
 * persistence contract; result trees/projections are intentionally separate
 * runtime values and must never be copied into WorksheetModel.cells.
 */
export const PIVOT_DEFINITION_SCHEMA = 'PivotDefinition' as const;
export const PIVOT_RESULT_TREE_SCHEMA = 'PivotResultTree' as const;
export const PIVOT_GRID_PROJECTION_SCHEMA = 'PivotGridProjection' as const;

/** A formula error remains a first-class Pivot member/value and never becomes blank. */
export interface PivotErrorValue {
  kind: 'error';
  code: FormulaErrorCode;
  message?: string;
}

export type PivotScalar = string | number | boolean | null | PivotErrorValue;
export type PivotScalarType = 'text' | 'number' | 'boolean' | 'blank' | 'error';
export const PIVOT_BLANK_LABEL = '(blank)' as const;

export function isPivotError(value: unknown): value is PivotErrorValue {
  return typeof value === 'object' && value !== null && (value as { kind?: unknown }).kind === 'error'
    && typeof (value as { code?: unknown }).code === 'string';
}

/** One civil day in the canonical UTC calendar used by Pivot timelines. */
export const PIVOT_DAY_MS = 86_400_000;

export interface PivotTimelinePeriodBounds {
  /** Inclusive start instant at 00:00:00 UTC for the selected start day. */
  start?: number;
  /** Exclusive instant at 00:00:00 UTC immediately after the selected end day. */
  endExclusive?: number;
}

const PIVOT_TIMELINE_DATE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

function parsePivotTimelineString(value: string): number | undefined {
  const match = PIVOT_TIMELINE_DATE_PATTERN.exec(value.trim());
  if (!match) return undefined;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4] ?? 0);
  const minute = Number(match[5] ?? 0);
  const second = Number(match[6] ?? 0);
  const millisecond = match[7] ? Number(match[7].padEnd(3, '0')) : 0;
  if (month < 1 || month > 12 || hour > 23 || minute > 59 || second > 59 || millisecond > 999) return undefined;

  // Date.UTC maps years 0..99 to 1900..1999. Constructing through setters
  // keeps the full four-digit year and lets the round-trip validate leap days.
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(hour, minute, second, millisecond);
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day
    || date.getUTCHours() !== hour || date.getUTCMinutes() !== minute || date.getUTCSeconds() !== second
    || date.getUTCMilliseconds() !== millisecond) return undefined;

  const zone = match[8];
  if (!zone || zone === 'Z') return date.getTime();
  const offsetHours = Number(zone.slice(1, 3));
  const offsetMinutes = Number(zone.slice(-2));
  if (offsetHours > 23 || offsetMinutes > 59) return undefined;
  const offset = (offsetHours * 60 + offsetMinutes) * 60_000;
  return date.getTime() + (zone[0] === '+' ? -offset : offset);
}

/** Parse a Pivot date value without relying on browser-local Date parsing. */
export function pivotTimelineInstant(value: PivotScalar): number | undefined {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return undefined;
    return Date.UTC(1899, 11, 30) + value * PIVOT_DAY_MS;
  }
  return typeof value === 'string' ? parsePivotTimelineString(value) : undefined;
}

function timelineDayStart(value: string, label: 'start' | 'end'): number | undefined {
  const instant = pivotTimelineInstant(value);
  if (instant === undefined || !Number.isFinite(instant)) {
    throw new Error(`Invalid Pivot timeline ${label} date: ${value}`);
  }
  return Math.floor(instant / PIVOT_DAY_MS) * PIVOT_DAY_MS;
}

/** Normalize an inclusive date-only period to one deterministic half-open interval. */
export function normalizePivotTimelinePeriod(period: { start?: string; end?: string }): PivotTimelinePeriodBounds {
  const start = period.start === undefined ? undefined : timelineDayStart(period.start, 'start');
  const endDay = period.end === undefined ? undefined : timelineDayStart(period.end, 'end');
  if (start !== undefined && endDay !== undefined && start > endDay) {
    throw new Error(`Pivot timeline period start must not be after end: ${period.start} > ${period.end}`);
  }
  return {
    ...(start === undefined ? {} : { start }),
    ...(endDay === undefined ? {} : { endExclusive: endDay + PIVOT_DAY_MS }),
  };
}

export interface PivotTimelineTile {
  key: string;
  start: string;
  end: string;
  label: string;
  hasData: boolean;
}

function timelineTileStart(instant: number, level: PivotTimelineLevel): Date {
  const date = new Date(Math.floor(instant / PIVOT_DAY_MS) * PIVOT_DAY_MS);
  if (level === 'years') date.setUTCMonth(0, 1);
  else if (level === 'quarters') date.setUTCMonth(Math.floor(date.getUTCMonth() / 3) * 3, 1);
  else if (level === 'months') date.setUTCDate(1);
  else date.setUTCHours(0, 0, 0, 0);
  return date;
}

function nextTimelineTileStart(start: Date, level: PivotTimelineLevel): Date {
  const next = new Date(start.getTime());
  if (level === 'years') next.setUTCFullYear(next.getUTCFullYear() + 1);
  else if (level === 'quarters') next.setUTCMonth(next.getUTCMonth() + 3);
  else if (level === 'months') next.setUTCMonth(next.getUTCMonth() + 1);
  else next.setUTCDate(next.getUTCDate() + 1);
  return next;
}

function timelineTileLabel(start: Date, level: PivotTimelineLevel): string {
  const year = start.getUTCFullYear();
  if (level === 'years') return String(year);
  if (level === 'quarters') return `${year} Q${Math.floor(start.getUTCMonth() / 3) + 1}`;
  const month = String(start.getUTCMonth() + 1).padStart(2, '0');
  if (level === 'months') return `${year}-${month}`;
  return `${year}-${month}-${String(start.getUTCDate()).padStart(2, '0')}`;
}

/** Build contiguous, data-aware period tiles from the canonical Pivot date values. */
export function buildPivotTimelineTiles(values: readonly PivotScalar[], level: PivotTimelineLevel): PivotTimelineTile[] {
  if (!['years', 'quarters', 'months', 'days'].includes(level)) throw new Error(`Invalid Pivot timeline level: ${String(level)}`);
  const instants = values.map(pivotTimelineInstant).filter((value): value is number => value !== undefined && Number.isFinite(value));
  if (instants.length === 0) return [];
  const first = timelineTileStart(Math.min(...instants), level);
  const last = timelineTileStart(Math.max(...instants), level);
  const dataKeys = new Set(instants.map((instant) => timelineTileStart(instant, level).getTime()));
  const tiles: PivotTimelineTile[] = [];
  let cursor = first;
  let guard = 0;
  while (cursor.getTime() <= last.getTime() && guard < 10_000) {
    const next = nextTimelineTileStart(cursor, level);
    tiles.push({
      key: `${level}:${cursor.toISOString().slice(0, 10)}`,
      start: cursor.toISOString().slice(0, 10),
      end: new Date(next.getTime() - PIVOT_DAY_MS).toISOString().slice(0, 10),
      label: timelineTileLabel(cursor, level),
      hasData: dataKeys.has(cursor.getTime()),
    });
    cursor = next;
    guard += 1;
  }
  if (guard >= 10_000) throw new Error('Pivot timeline period range exceeds the supported bound');
  return tiles;
}

/** One presentation rule for Pivot members across grid, filters, and controls. */
export function formatPivotMember(value: PivotScalar): string {
  if (isPivotError(value)) return value.code;
  return value === null || value === '' ? PIVOT_BLANK_LABEL : String(value);
}

/** A member key keeps `1`, `"1"`, `true`, and blank members distinct. */
export interface PivotMemberKey {
  type: PivotScalarType;
  value: string | number | boolean | null;
}

export function createPivotMemberKey(value: PivotScalar): PivotMemberKey {
  if (value === null || value === '') return { type: 'blank', value: null };
  if (isPivotError(value)) return { type: 'error', value: value.code };
  if (typeof value === 'number') return { type: 'number', value };
  if (typeof value === 'boolean') return { type: 'boolean', value };
  return { type: 'text', value };
}

export function pivotMemberKey(value: PivotMemberKey): string {
  return `${value.type}:${JSON.stringify(value.value)}`;
}

export function pivotMemberKeyEquals(left: PivotMemberKey, right: PivotMemberKey): boolean {
  return left.type === right.type && left.value === right.value;
}

export function pivotScalarFromMemberKey(value: PivotMemberKey): PivotScalar {
  if (value.type === 'blank') return null;
  if (value.type === 'error') return { kind: 'error', code: value.value as FormulaErrorCode };
  return value.value as Exclude<PivotScalar, null | PivotErrorValue>;
}

export type PivotAggregateFunction =
  | 'sum'
  | 'count'
  | 'count-numbers'
  | 'average'
  | 'min'
  | 'max'
  | 'product'
  | 'stdev'
  | 'stdevp'
  | 'var'
  | 'varp'
  | 'distinct-count';

export type PivotSubtotalDefinition =
  | { mode: 'automatic' }
  | { mode: 'none' }
  | { mode: 'custom'; functions: PivotAggregateFunction[] };

export type PivotSubtotalLocation = 'top' | 'bottom' | 'off';

/**
 * Canonical Excel PivotTable report layout.  This is intentionally one
 * semantic value: compact, outline, and tabular are not independent flags.
 */
export type PivotReportLayout = 'compact' | 'outline' | 'tabular';

export interface PivotSourceRelationship {
  id: string;
  left: { sourceId: string; fieldId: string };
  right: { sourceId: string; fieldId: string };
  join: 'inner' | 'left';
}

/** A worksheet range is a named logical source node, not merely a position in an array. */
export interface PivotWorksheetSourceRange {
  sourceId: string;
  range: RangeRef;
}

/** Canonical Pivot source. */
export type PivotWorksheetDataSource =
  | { kind: 'worksheet-range'; range: RangeRef }
  | { kind: 'worksheet-ranges'; ranges: PivotWorksheetSourceRange[]; relationships: PivotSourceRelationship[] };

export type PivotSource = PivotWorksheetDataSource
  | { kind: 'table'; tableId: string }
  | { kind: 'named-range'; name: string; sheetId?: SheetId }
  | { kind: 'data-source'; dataSourceId: string };

export type PivotFieldDataType = 'text' | 'number' | 'date' | 'boolean' | 'error' | 'mixed';

export interface PivotFieldDefinition {
  /** Stable identity. It is derived from the source column, never from a row value. */
  fieldId: string;
  name: string;
  dataType: PivotFieldDataType;
  ordinal: number;
  values?: PivotScalar[];
}

export interface PivotFieldCatalog {
  schema?: 'PivotFieldCatalog';
  fields: PivotFieldDefinition[];
}

export interface PivotManualGroup {
  groupId: string;
  name: string;
  items: PivotMemberKey[];
}

export type PivotDateGroupUnit = 'year' | 'quarter' | 'month' | 'week' | 'day';

export type PivotGroup =
  | { kind: 'date'; unit: PivotDateGroupUnit; units?: PivotDateGroupUnit[]; startOfWeek?: 0 | 1 | 2 | 3 | 4 | 5 | 6; start?: PivotScalar; end?: PivotScalar; autoStart?: boolean; autoEnd?: boolean }
  | { kind: 'number'; interval: number; start?: number; end?: number; autoStart?: boolean; autoEnd?: boolean }
  | { kind: 'manual'; groups: PivotManualGroup[] };

export type PivotSort = {
  direction: 'ascending' | 'descending';
  by?: 'label' | 'value';
  valueFieldId?: string;
};

export interface PivotFieldPlacement {
  fieldId: string;
  sort?: PivotSort;
  group?: PivotGroup;
  subtotal?: PivotSubtotalDefinition;
}

/** The independent predicate family used to identify a Pivot filter slot. */
export type PivotFilterFamily = 'manual' | 'label' | 'date' | 'value' | 'top-items';

export type PivotLabelFilterOperator = 'equals' | 'not-equals' | 'begins-with' | 'not-begins-with' | 'ends-with' | 'not-ends-with' | 'contains' | 'not-contains' | 'between' | 'not-between' | 'greater-than' | 'greater-or-equal' | 'less-than' | 'less-or-equal';
export type PivotDateFilterOperator = 'equals' | 'not-equals' | 'before' | 'after' | 'between' | 'not-between';
export type PivotValueFilterOperator = 'equals' | 'not-equals' | 'greater-than' | 'greater-or-equal' | 'less-than' | 'less-or-equal' | 'between' | 'not-between';
export type PivotDynamicDateFilter = 'today' | 'yesterday' | 'tomorrow' | 'this-week' | 'last-week' | 'next-week' | 'this-month' | 'last-month' | 'next-month' | 'this-quarter' | 'last-quarter' | 'next-quarter' | 'this-year' | 'last-year' | 'next-year' | 'year-to-date';

/** Persisted document-owned text collation. It must never be inferred from a host locale. */
export interface PivotCollation {
  locale: string;
  sensitivity: 'base' | 'accent' | 'case' | 'variant';
  numeric: boolean;
  caseFirst: 'upper' | 'lower' | 'false';
}

export const DEFAULT_PIVOT_COLLATION: PivotCollation = {
  locale: 'en-US',
  sensitivity: 'variant',
  numeric: false,
  caseFirst: 'false',
};

export type PivotManualFilter = {
  kind: 'manual';
  family: 'manual';
  fieldId: string;
  /** Report filters occupy the Filters area; field filters stay attached to a row/column field. */
  scope?: 'report' | 'field';
  mode: 'all' | 'include' | 'exclude';
  memberKeys: PivotMemberKey[];
};

export interface PivotLabelFilter {
  kind: 'condition';
  family: 'label';
  fieldId: string;
  valueFieldId?: string;
  scope?: 'report' | 'field';
  operator: PivotLabelFilterOperator;
  value: PivotScalar;
  value2?: PivotScalar;
  dynamic?: PivotDynamicDateFilter;
}

export interface PivotDateFilter {
  kind: 'condition';
  family: 'date';
  fieldId: string;
  valueFieldId?: string;
  scope?: 'report' | 'field';
  operator: PivotDateFilterOperator;
  value: PivotScalar;
  value2?: PivotScalar;
  dynamic?: PivotDynamicDateFilter;
  wholeDay?: boolean;
}

export interface PivotValueFilter {
  kind: 'condition';
  family: 'value';
  fieldId: string;
  /** Optional measure identity for native value filters. */
  valueFieldId?: string;
  scope?: 'report' | 'field';
  operator: PivotValueFilterOperator;
  value: PivotScalar;
  value2?: PivotScalar;
  dynamic?: PivotDynamicDateFilter;
}

export type PivotFilter =
  | PivotManualFilter
  | PivotLabelFilter
  | PivotDateFilter
  | PivotValueFilter
  | {
    kind: 'top-items';
    family: 'top-items';
    fieldId: string;
    scope?: 'report' | 'field';
    count: number;
    valueFieldId: string;
    direction: 'top' | 'bottom';
  };

/**
 * The presentation-neutral state of one report-filter family.  This is
 * deliberately derived from the persisted filter, rather than storing a
 * localized caption or an internal field id in the projection.
 */
export type PivotReportFilterSummaryEntry =
  | {
    kind: 'manual';
    family: 'manual';
    active: boolean;
    mode: 'all' | 'include' | 'exclude';
    count: number;
    memberValues: PivotScalar[];
  }
  | ({
    kind: 'condition';
    active: true;
    value: PivotScalar;
    value2?: PivotScalar;
    dynamic?: PivotDynamicDateFilter;
    valueFieldName?: string;
  } & ({ family: 'label'; operator: PivotLabelFilterOperator }
    | { family: 'date'; operator: PivotDateFilterOperator }
    | { family: 'value'; operator: PivotValueFilterOperator }))
  | {
    kind: 'top-items';
    family: 'top-items';
    active: true;
    count: number;
    direction: 'top' | 'bottom';
    valueFieldName: string;
  };

/** A grouped report-filter summary; one field can expose multiple families. */
export interface PivotReportFilterSummary {
  fieldName: string;
  active: boolean;
  entries: PivotReportFilterSummaryEntry[];
}

/** Derived Slicer item state; never authored or persisted as filter state. */
export interface PivotSlicerItemProjection {
  key: PivotMemberKey;
  value: PivotScalar;
  label: string;
  selected: boolean;
  hasData: boolean;
}

export type PivotShowAs =
  | { kind: 'normal' }
  | { kind: 'grand-percentage' }
  | { kind: 'row-percentage' }
  | { kind: 'column-percentage' }
  | { kind: 'parent-percentage' }
  | { kind: 'difference'; base: 'grand' | 'row' | 'column' | 'parent' }
  | { kind: 'percentage-difference'; base: 'grand' | 'row' | 'column' | 'parent' }
  | { kind: 'running-total'; axis: 'row' | 'column' }
  | { kind: 'rank'; axis: 'row' | 'column'; direction: 'ascending' | 'descending' }
  | { kind: 'index' };

export interface PivotValueField {
  fieldId: string;
  summarizeBy: PivotAggregateFunction;
  displayName?: string;
  numberFormat?: string;
  baseFieldId?: string;
  baseItem?: PivotMemberKey | string;
  showAs?: PivotShowAs;
}

/**
 * Normalize the one persisted number-format contract for Pivot value fields.
 *
 * The format is an Excel format code, not display text.  Empty, malformed or
 * control-character-containing codes are rejected at the model boundary so a
 * renderer or OOXML writer cannot silently fall back to General.
 */
export function normalizePivotNumberFormat(input: unknown): string | undefined {
  if (input === undefined) return undefined;
  if (typeof input !== 'string') throw new Error('Pivot value field numberFormat must be a string');
  const format = input.trim();
  if (!format) throw new Error('Pivot value field numberFormat must not be empty');
  if (format.length > 255 || /[\u0000-\u001f\u007f]/.test(format)) throw new Error('Pivot value field numberFormat is invalid');

  let quoted = false;
  let bracketed = false;
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index]!;
    if (quoted) {
      if (character === '"') quoted = false;
      continue;
    }
    if (bracketed) {
      if (character === ']') bracketed = false;
      continue;
    }
    if (character === '"') {
      quoted = true;
      continue;
    }
    if (character === '[') {
      bracketed = true;
      continue;
    }
    if (character === ']') throw new Error('Pivot value field numberFormat has an unmatched bracket');
    if (character === '\\') {
      if (index + 1 >= format.length) throw new Error('Pivot value field numberFormat has a dangling escape');
      index += 1;
    }
  }
  if (quoted || bracketed) throw new Error('Pivot value field numberFormat is unterminated');
  if (format.split(';').some((section) => section.trim().length === 0)) throw new Error('Pivot value field numberFormat has an empty section');
  return format;
}

export interface PivotCalculatedField {
  fieldId: string;
  name: string;
  formula: string;
}

export interface PivotCalculatedItem {
  fieldId: string;
  targetFieldId: string;
  name: string;
  formula: string;
}

export interface PivotExpansionState {
  /** Stable result-node paths. This is intentionally not a field-level flag. */
  expandedNodeIds: string[];
  collapsedNodeIds: string[];
  showButtons: boolean;
}

export interface PivotLayout {
  rows: PivotFieldPlacement[];
  columns: PivotFieldPlacement[];
  filters: PivotFilter[];
  /** Explicitly controls whether more than one predicate may target a field. */
  allowMultipleFiltersPerField: boolean;
  collation: PivotCollation;
  values: PivotValueField[];
  calculatedFields?: PivotCalculatedField[];
  calculatedItems?: PivotCalculatedItem[];
  subtotalLocation: PivotSubtotalLocation;
  /** Controls the grand-total column (totals across each row's columns). */
  showRowGrandTotals: boolean;
  /** Controls the grand-total row (totals across each column's rows). */
  showColumnGrandTotals: boolean;
  /** Controls the row-field presentation and label repetition semantics. */
  reportLayout: PivotReportLayout;
  expansion?: PivotExpansionState;
}

export function pivotFilterScope(filter: Pick<PivotFilter, 'scope'>): 'report' | 'field' {
  return filter.scope ?? 'report';
}

export function pivotFilterIdentity(filter: Pick<PivotFilter, 'fieldId' | 'family' | 'scope'>): string {
  return `${filter.fieldId}|${pivotFilterScope(filter)}|${filter.family}`;
}

export function allowsMultiplePivotFilters(layout: Pick<PivotLayout, 'allowMultipleFiltersPerField'>): boolean {
  return layout.allowMultipleFiltersPerField;
}

export function createPivotCollator(collation: PivotCollation): Intl.Collator {
  if (!collation || typeof collation.locale !== 'string' || !collation.locale.trim()
    || !['base', 'accent', 'case', 'variant'].includes(collation.sensitivity)
    || typeof collation.numeric !== 'boolean' || !['upper', 'lower', 'false'].includes(collation.caseFirst)) {
    throw new Error('Pivot collation is invalid');
  }
  try {
    return new Intl.Collator(collation.locale, {
      sensitivity: collation.sensitivity,
      numeric: collation.numeric,
      caseFirst: collation.caseFirst,
    });
  } catch (error) {
    throw new Error(`Pivot collation locale is unsupported: ${collation.locale}`, { cause: error });
  }
}

export interface PivotStyleOptions {
  showRowHeaders: boolean;
  showColumnHeaders: boolean;
  showRowStripes: boolean;
  showColumnStripes: boolean;
  showLastColumn: boolean;
}

/** Persistent Layout & Format / Display options owned by one Pivot definition. */
export interface PivotDisplayOptions {
  fillEmptyCells: boolean;
  emptyCellText: string;
  showErrorValues: boolean;
  errorCellText: string;
  showFieldHeaders: boolean;
  autoFitColumnsOnUpdate: boolean;
}

export interface PivotPresentation {
  styleName?: string;
  styleOptions: PivotStyleOptions;
  displayOptions?: PivotDisplayOptions;
}

export const DEFAULT_PIVOT_STYLE_OPTIONS: PivotStyleOptions = {
  showRowHeaders: true,
  showColumnHeaders: true,
  showRowStripes: false,
  showColumnStripes: false,
  showLastColumn: false,
};

export const DEFAULT_PIVOT_DISPLAY_OPTIONS: PivotDisplayOptions = {
  fillEmptyCells: false,
  emptyCellText: '',
  showErrorValues: true,
  errorCellText: '',
  showFieldHeaders: true,
  autoFitColumnsOnUpdate: true,
};

export function normalizePivotDisplayOptions(input?: Partial<PivotDisplayOptions>): PivotDisplayOptions {
  const allowed = new Set<keyof PivotDisplayOptions>(['fillEmptyCells', 'emptyCellText', 'showErrorValues', 'errorCellText', 'showFieldHeaders', 'autoFitColumnsOnUpdate']);
  if (input && Object.keys(input).some((key) => !allowed.has(key as keyof PivotDisplayOptions))) {
    throw new Error('Pivot display options contain unsupported fields');
  }
  const options = { ...DEFAULT_PIVOT_DISPLAY_OPTIONS, ...(input ?? {}) };
  if (typeof options.fillEmptyCells !== 'boolean' || typeof options.emptyCellText !== 'string'
    || typeof options.showErrorValues !== 'boolean' || typeof options.errorCellText !== 'string'
    || typeof options.showFieldHeaders !== 'boolean' || typeof options.autoFitColumnsOnUpdate !== 'boolean') {
    throw new Error('Pivot display options are invalid');
  }
  return {
    fillEmptyCells: options.fillEmptyCells,
    emptyCellText: options.emptyCellText,
    showErrorValues: options.showErrorValues,
    errorCellText: options.errorCellText,
    showFieldHeaders: options.showFieldHeaders,
    autoFitColumnsOnUpdate: options.autoFitColumnsOnUpdate,
  };
}

export interface PivotRefreshPolicy {
  mode: 'manual' | 'on-open' | 'on-change';
  preserveFormatting: boolean;
  /**
   * Derived wire projection of `mode`.  It remains on the public contract for
   * the current protocol revision, but canonicalization rejects a value that
   * disagrees with the mode, so there is only one refresh decision.
   */
  refreshOnLoad: boolean;
}

export interface PivotNativeCacheFlags {
  /** Native attributes belong to the shared PivotCache, not one PivotTable. */
  refreshOnLoad?: boolean;
  refreshOnSave?: boolean;
  saveData?: boolean;
  enableRefresh?: boolean;
}

/**
 * Native Pivot filters which cannot be represented by the canonical filter
 * algebra remain explicit boundary-owned metadata. The original attributes
 * are retained so an import/export round trip cannot silently broaden or
 * remove a filter that the runtime does not understand.
 */
export interface PivotNativeFilterMetadata {
  fieldIndex: number;
  type: string;
  attributes: Record<string, string>;
}

export interface PivotNativeAutoSortMetadata {
  fieldIndex: number;
  sortType?: 'manual' | 'ascending' | 'descending';
  nonAutoSortDefault?: boolean;
  attributes: Record<string, string>;
  references: Array<{
    field: number;
    selected?: boolean;
    itemIndexes?: number[];
  }>;
}

export interface PivotNativeMetadata {
  /** Stable native/cache identity used when several PivotTables share a cache. */
  cacheKey?: string;
  cacheId?: number;
  cacheDefinitionPart?: string;
  cacheRecordsPart?: string;
  pivotTablePart?: string;
  /**
   * The original cache-level flags are preserved at the OOXML boundary. They
   * are not a second runtime refresh policy; edited canonical mode owns the
   * refreshOnLoad/refreshOnSave values when the package is regenerated.
   */
  cacheFlags?: PivotNativeCacheFlags;
  fieldBindings?: Record<string, { cacheFieldIndex: number; sourceName?: string }>;
  /** Native filters retained when their exact semantics exceed PivotFilter. */
  preservedPivotFilters?: PivotNativeFilterMetadata[];
  /** Native auto-sort scopes retained when their exact semantics exceed PivotSort. */
  preservedAutoSortScopes?: PivotNativeAutoSortMetadata[];
  /** Only canonical identities, presentation, and explicit native-preservation records cross the model boundary. */
  preservedFeatures?: Array<'external-connection' | 'olap' | 'consolidation' | 'macro' | 'custom-xml' | 'slicer' | 'timeline'>;
}

export interface PivotTarget {
  sheetId: SheetId;
  anchor: { row: Row; column: Column };
}

/** Canonical persisted Pivot definition. */
export interface PivotDefinition {
  schema: typeof PIVOT_DEFINITION_SCHEMA;
  id: string;
  source: PivotSource;
  target: PivotTarget;
  fieldCatalog: PivotFieldCatalog;
  layout: PivotLayout;
  refreshPolicy: PivotRefreshPolicy;
  presentation?: PivotPresentation;
  nativeMetadata?: PivotNativeMetadata;
}

/** Public model name used by workbook collections; there is exactly one Pivot shape. */
export type PivotModel = PivotDefinition;

export interface PivotSourceRowPath {
  /** Logical source node identity; required for multi-range joins and optional for legacy single-source detail payloads. */
  sourceId?: string;
  /** Joined-record identity; allows Show Details to group provenance without relying on source-array order. */
  recordId?: string;
  sheetId: SheetId;
  row: Row;
}

export interface PivotResultCell {
  id?: string;
  nodePath?: string[];
  kind?: 'detail' | 'subtotal' | 'grand-total';
  columnPath: PivotScalar[];
  values: PivotScalar[];
  sourceRowPaths: PivotSourceRowPath[];
}

export interface PivotResultValueField extends PivotValueField {
  sourceFieldId: string;
  subtotalFunction?: PivotAggregateFunction;
  subtotalFieldId?: string;
}

export interface PivotResultNode {
  nodeId?: string;
  path?: string[];
  kind: 'leaf' | 'subtotal';
  fieldId?: string;
  memberKey?: PivotMemberKey;
  key: PivotScalar;
  label: string;
  depth: number;
  children: PivotResultNode[];
  values: PivotResultCell[];
  /** Aggregate across every column path for this row node. */
  rowGrandTotal?: PivotResultCell;
  subtotal: boolean;
  sourceRowPaths: PivotSourceRowPath[];
}

export interface PivotResultTree {
  schema: typeof PIVOT_RESULT_TREE_SCHEMA;
  pivotId: string;
  fields: PivotFieldCatalog;
  columnPaths: PivotScalar[][];
  valueFields?: PivotResultValueField[];
  rows: PivotResultNode[];
  /** Per-column aggregates used by the grand-total row. */
  columnGrandTotals?: PivotResultCell[];
  grandTotal: PivotResultCell | null;
  sourceRowPaths: PivotSourceRowPath[];
  sourceRevision?: string;
  layoutRevision?: string;
  filterRevision?: string;
  /** Derived item availability for floating Slicer controls, keyed by drawing id. */
  slicerItems?: Record<string, PivotSlicerItemProjection[]>;
}

export type PivotProjectionCellKind =
  | 'title'
  | 'filter'
  | 'row-header'
  | 'column-header'
  | 'value'
  | 'subtotal'
  | 'grand-total'
  | 'expand-toggle'
  | 'loading'
  | 'error';

export interface PivotProjectionCell {
  id: string;
  pivotId: string;
  row: number;
  column: number;
  kind: PivotProjectionCellKind;
  value: PivotScalar;
  text: string;
  /** Canonical Excel format used for value/subtotal/grand-total cells. */
  numberFormat?: string;
  /** Canonical field identity for header/filter interactions. */
  fieldId?: string;
  /** Locale-independent caption owned by the presentation layer. */
  captionKey?: 'row-labels' | 'grand-total' | 'loading';
  filterSummary?: PivotReportFilterSummary;
  nodeId?: string;
  resultCellId?: string;
  columnPath?: PivotScalar[];
  isLastColumn?: boolean;
  sourceRowPaths?: PivotSourceRowPath[];
  expandable?: boolean;
  expanded?: boolean;
}

export interface PivotCollision {
  status: 'clear' | 'collision';
  reasons: Array<'cell-data' | 'merge' | 'pivot' | 'worksheet-bounds'>;
  conflictingRanges: RangeRef[];
}

export type PivotRefreshStatus = 'idle' | 'refreshing' | 'ready' | 'stale' | 'error' | 'collision';

export interface PivotRefreshState {
  status: PivotRefreshStatus;
  revision: number;
  sourceRevision: string;
  requestedAt?: string;
  completedAt?: string;
  error?: string;
}

export interface PivotGridProjection {
  schema: typeof PIVOT_GRID_PROJECTION_SCHEMA;
  pivotId: string;
  sheetId: SheetId;
  target: PivotTarget;
  presentation?: PivotPresentation;
  occupiedRange: RangeRef;
  cells: PivotProjectionCell[];
  collision: PivotCollision;
  refresh: PivotRefreshState;
}

export type PivotHitKind = 'cell' | 'expand-toggle' | 'filter' | 'header' | 'none';

export interface PivotHitTest {
  kind: PivotHitKind;
  pivotId: string;
  cellId?: string;
  row?: number;
  column?: number;
  nodeId?: string;
  sourceRowPaths?: PivotSourceRowPath[];
}

/** Generic context-hit shape used by the UI resolver for Pivot cells. */
export interface ContextHit extends PivotHitTest {
  context: 'pivot';
  priority: 30;
}

/**
 * Pure snapshot migration. It only translates shape/identity fields; live
 * source inspection and field catalog inference belong to the spreadsheet-app
 * engine where a WorkbookModel is available.
 */
export function canonicalizePivotDefinition(input: PivotDefinition): PivotDefinition {
  if (input.schema !== PIVOT_DEFINITION_SCHEMA) throw new Error(`Pivot ${input.id} is not a canonical definition`);
  if (!input.source || !input.target || !input.fieldCatalog || !input.refreshPolicy) throw new Error(`Pivot ${input.id} is missing canonical fields`);
  const refreshPolicy = normalizePivotRefreshPolicy(input.refreshPolicy);
  if (input.layout.rows.some((entry) => !entry.fieldId)
    || input.layout.columns.some((entry) => !entry.fieldId)
    || input.layout.values.some((entry) => !entry.fieldId)
    || input.layout.filters.some((entry) => !entry.fieldId)) {
    throw new Error(`Pivot ${input.id} has non-canonical field references`);
  }
  const canonical = structuredClone(input);
  if (typeof canonical.layout.allowMultipleFiltersPerField !== 'boolean') throw new Error(`Pivot ${input.id} is missing allowMultipleFiltersPerField`);
  canonical.layout.values = canonical.layout.values.map((field) => {
    const numberFormat = normalizePivotNumberFormat(field.numberFormat);
    return { ...field, ...(numberFormat === undefined ? {} : { numberFormat }) };
  });
  createPivotCollator(canonical.layout.collation);
  const axisFields = new Set([...canonical.layout.rows, ...canonical.layout.columns].map((entry) => entry.fieldId));
  const identities = new Set<string>();
  for (const filter of canonical.layout.filters) {
    const expectedFamily = filter.kind === 'manual' ? 'manual' : filter.kind === 'top-items' ? 'top-items' : filter.family;
    if (filter.family !== expectedFamily) throw new Error(`Pivot ${input.id} has an invalid filter family`);
    if (filter.scope === 'field' && !axisFields.has(filter.fieldId)) {
      throw new Error(`Pivot ${input.id} field filter must target a row or column field`);
    }
    const identity = pivotFilterIdentity(filter);
    if (identities.has(identity)) throw new Error(`Pivot ${input.id} contains duplicate filter family ${identity}`);
    identities.add(identity);
  }
  canonical.refreshPolicy = refreshPolicy;
  canonical.presentation = {
    ...(canonical.presentation?.styleName ? { styleName: canonical.presentation.styleName } : {}),
    styleOptions: { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(canonical.presentation?.styleOptions ?? {}) },
    ...(canonical.presentation?.displayOptions ? { displayOptions: normalizePivotDisplayOptions(canonical.presentation.displayOptions) } : {}),
  };
  canonical.layout.filters = canonical.layout.filters.map((filter) => ({
    ...filter,
    scope: filter.scope ?? (axisFields.has(filter.fieldId) ? 'field' : 'report'),
  }));
  return canonical;
}

/**
 * Normalize the one canonical refresh decision. `on-change` includes opening
 * refresh because a cache that is refreshed from source changes must also be
 * current when the workbook is opened; manual is the only non-opening mode.
 */
export function normalizePivotRefreshPolicy(input: PivotRefreshPolicy): PivotRefreshPolicy {
  if (!input || !['manual', 'on-open', 'on-change'].includes(input.mode) || typeof input.preserveFormatting !== 'boolean' || typeof input.refreshOnLoad !== 'boolean') {
    throw new Error('Pivot refresh policy is invalid');
  }
  const expectedRefreshOnLoad = input.mode !== 'manual';
  if (input.refreshOnLoad !== expectedRefreshOnLoad) {
    throw new Error(`Pivot refresh policy is contradictory for mode ${input.mode}`);
  }
  return { mode: input.mode, preserveFormatting: input.preserveFormatting, refreshOnLoad: expectedRefreshOnLoad };
}

export function refreshOnSaveForPivotMode(mode: PivotRefreshPolicy['mode']): boolean {
  return mode === 'on-change';
}
