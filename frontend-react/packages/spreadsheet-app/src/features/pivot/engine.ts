import type {
  PivotAggregateFunction,
  PivotDefinition,
  PivotFieldCatalog,
  PivotFieldDataType,
  PivotFieldDefinition,
  PivotFieldPlacement,
  PivotFilter,
  PivotGroup,
  PivotGridProjection,
  PivotHitTest,
  PivotLayout,
  PivotMemberKey,
  PivotModel,
  PivotProjectionCell,
  PivotRefreshState,
  PivotResultCell,
  PivotResultNode,
  PivotResultTree,
  PivotScalar,
  PivotSource,
  PivotSourceRowPath,
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
  PIVOT_RESULT_TREE_SCHEMA,
  createPivotMemberKey,
  pivotMemberKey,
  pivotMemberKeyEquals,
  pivotScalarFromMemberKey,
} from '@react-sheets/core-model';
import { FormulaEngine, type FormulaValue } from '@react-sheets/formula-engine';

export interface PivotSourceRowInput {
  values: Record<string, PivotScalar>;
  paths: PivotSourceRowPath[];
}

export interface PivotSourceFieldInput {
  fieldId: string;
  name: string;
  ordinal: number;
  dataType?: PivotFieldDataType;
}

export interface PivotSourceTableInput {
  fields: PivotSourceFieldInput[];
  rows: PivotSourceRowInput[];
}

type SourceTable = PivotSourceTableInput;
type SourceRow = PivotSourceRowInput;
type SourceField = PivotSourceFieldInput;

interface AxisGroup {
  values: PivotScalar[];
  rows: SourceRow[];
}

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

const same = (left: PivotScalar, right: PivotScalar): boolean => {
  if ((left == null || left === '') && (right == null || right === '')) return true;
  return left === right;
};

const display = (value: PivotScalar): string => value == null || value === '' ? '(blank)' : String(value);

function stableSerialize(value: unknown): string {
  if (value === undefined) return 'undefined';
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(stableSerialize).join(',')}]`;
  const record = value as Record<string, unknown>;
  return `{${Object.keys(record).sort().map((key) => `${JSON.stringify(key)}:${stableSerialize(record[key])}`).join(',')}}`;
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

function sourceRevision(workbook: WorkbookModel, pivot: PivotModel): string {
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') {
    const manifest = workbook.getDataSource(source.dataSourceId);
    return fingerprint({
      source,
      revision: manifest.revision,
      blocks: manifest.blocks.map((block) => ({ id: block.id, checksum: block.checksum, revision: block.revision })),
    });
  }
  const ranges = sourceRanges(workbook, pivot);
  const revisions = ranges.map((range) => {
    const sheet = workbook.getSheet(range.sheetId);
    // CellMatrix revision is supplied by the block/data-source implementation
    // when available. Do not scan a whole range merely to build a cache key.
    const revision = (sheet.cells as unknown as { revision?: number }).revision;
    return `${range.sheetId}:${revision ?? 'live'}:${sheet.cells.count()}`;
  });
  return fingerprint({ source, revisions });
}

function linkedFilterDefinitions(workbook: WorkbookModel, pivot: PivotModel): unknown[] {
  return workbook.getSheets().flatMap((sheet) => sheet.drawings.map((drawing) => {
    const payload = sheet.drawingPayloads.get(drawing.payloadId);
    if (!payload || (payload.kind !== 'slicer' && payload.kind !== 'timeline')) return undefined;
    const linked = [payload.pivotId, ...(payload.connectedPivotIds ?? [])];
    return linked.includes(pivot.id) ? { drawingId: drawing.id, payload } : undefined;
  })).filter((entry): entry is NonNullable<typeof entry> => Boolean(entry));
}

export function getPivotRevisionKey(workbook: WorkbookModel, pivot: PivotModel): PivotRevisionKey {
  const definition = normalizePivotDefinition(workbook, pivot);
  return {
    pivotId: definition.id,
    sourceRevision: sourceRevision(workbook, pivot),
    layoutRevision: fingerprint({ source: definition.source, fieldCatalog: definition.fieldCatalog, layout: definition.layout }),
    filterRevision: fingerprint({ filters: definition.layout.filters, linked: linkedFilterDefinitions(workbook, pivot) }),
  };
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
  return `sheet:${range.sheetId}:column:${range.startColumn + ordinal}:range:${rangeIndex}`;
}

/** Stable field identity used by the catalog and all layout references. */
export function getStablePivotFieldId(source: PivotSource, range: RangeRef, ordinal: number, rangeIndex = 0): string {
  return sourceIdentity(source, range, ordinal, rangeIndex);
}

function sourceRanges(workbook: WorkbookModel, pivot: PivotModel): RangeRef[] {
  const source = getPivotSource(pivot);
  if (source.kind === 'worksheet-range') return [source.range];
  if (source.kind === 'worksheet-ranges') return source.ranges;
  if (source.kind === 'table') {
    return [resolvePivotTable(workbook, source.tableId).range];
  }
  if (source.kind === 'data-source') {
    const manifest = workbook.getDataSource(source.dataSourceId);
    if (!manifest.sourceRange) throw new Error(`Pivot data source ${source.dataSourceId} has no worksheet range`);
    return [manifest.sourceRange];
  }
  return [resolveNamedRange(workbook, source.name, source.sheetId ?? pivot.target.sheetId)];
}

function resolvePivotTable(workbook: WorkbookModel, tableId: string): {
  range: RangeRef;
  fields: Array<{ id: string; name: string }>;
} {
  const workbookTable = workbook.tables.get(tableId);
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
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || value === null ? value : null;
}

function inferType(values: PivotScalar[]): PivotFieldDataType {
  const present = values.filter((value) => value != null && value !== '');
  if (!present.length) return 'mixed';
  if (present.every((value) => typeof value === 'boolean')) return 'boolean';
  if (present.every((value) => typeof value === 'number' && Number.isFinite(value))) return 'number';
  const dateLike = present.every((value) => typeof value === 'string' && /^\d{4}-\d{2}-\d{2}(?:[T ].*)?$/.test(value) && !Number.isNaN(Date.parse(value)));
  if (dateLike) return 'date';
  if (present.every((value) => typeof value === 'string')) return 'text';
  return 'mixed';
}

function parseColumnLabel(value: string): number {
  let column = 0;
  for (const character of value.toUpperCase()) {
    if (character < 'A' || character > 'Z') throw new Error(`Invalid named range column: ${value}`);
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column - 1;
}

function parseA1Range(formula: string, workbook: WorkbookModel, fallbackSheetId: string): RangeRef {
  const cleaned = formula.trim().replace(/^=/, '').replace(/^\+/, '');
  const match = cleaned.match(/^(?:'((?:[^']|'')+)'|([A-Za-z0-9_-]+))?!?\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/);
  if (!match) throw new Error(`Named range is not a worksheet range: ${formula}`);
  const sheetName = (match[1] ?? match[2])?.replace(/''/g, "'");
  const sheet = sheetName ? workbook.getSheetByName(sheetName) : workbook.getSheet(fallbackSheetId);
  if (!sheet) throw new Error(`Named range references unknown worksheet: ${sheetName ?? fallbackSheetId}`);
  const startColumn = parseColumnLabel(match[3]!);
  const startRow = Number(match[4]) - 1;
  const endColumn = match[5] ? parseColumnLabel(match[5]) : startColumn;
  const endRow = match[6] ? Number(match[6]) - 1 : startRow;
  if (startRow < 0 || endRow < startRow || startColumn < 0 || endColumn < startColumn) throw new Error(`Invalid named range: ${formula}`);
  return { sheetId: sheet.id, startRow, endRow, startColumn, endColumn };
}

function resolveNamedRange(workbook: WorkbookModel, name: string, sheetId?: string): RangeRef {
  const formula = workbook.getDefinedName(name, sheetId)?.formula ?? '';
  if (!formula) throw new Error(`Unknown named range: ${name}`);
  return parseA1Range(formula, workbook, workbook.primarySheetId);
}

function readRange(sheet: WorksheetModel, range: RangeRef, source: PivotSource, rangeIndex: number, persisted?: PivotFieldCatalog): SourceTable {
  const fields: SourceField[] = [];
  const persistedFields = persisted?.fields ?? [];
  for (let ordinal = 0; ordinal <= range.endColumn - range.startColumn; ordinal += 1) {
    const column = range.startColumn + ordinal;
    const raw = sheet.cells.get(range.startRow, column)?.value;
    const name = raw == null || raw === '' ? `Column ${ordinal + 1}` : String(raw);
    // Ordinal/source-column identity survives a header rename. A changed
    // physical column is a new field, while a changed caption is not.
    const persistedField = persistedFields.find((field) => field.ordinal === ordinal);
    fields.push({ fieldId: persistedField?.fieldId ?? sourceIdentity(source, range, ordinal, rangeIndex), name, ordinal });
  }
  const rows: SourceRow[] = [];
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const values: Record<string, PivotScalar> = {};
    fields.forEach((field, ordinal) => {
      const cell = sheet.cells.get(row, range.startColumn + ordinal);
      values[field.fieldId] = cellScalar(cell?.formulaValue ?? cell?.value ?? null);
    });
    rows.push({ values, paths: [{ sheetId: range.sheetId, row }] });
  }
  return { fields, rows };
}

function sourceTable(workbook: WorkbookModel, pivot: PivotModel, catalog?: PivotFieldCatalog): SourceTable {
  const source = getPivotSource(pivot);
  if (source.kind === 'data-source') {
    throw new Error(`Block-backed data source ${source.dataSourceId} requires asynchronous Pivot computation`);
  }
  const ranges = sourceRanges(workbook, pivot);
  if (source.kind === 'worksheet-ranges') {
    const tables = ranges.map((range, index) => readRange(workbook.getSheet(range.sheetId), range, source, index, catalog));
    let current = tables[0] ?? { fields: [], rows: [] };
    for (let index = 1; index < tables.length; index += 1) {
      const right = tables[index]!;
      const relationship = source.relationships.find((candidate) => candidate.left.sheetId === ranges[index - 1]!.sheetId && candidate.right.sheetId === ranges[index]!.sheetId);
      if (!relationship) throw new Error('Every local worksheet range must have an adjacent typed relationship');
      const leftField = relationship.left.fieldId;
      const rightField = relationship.right.fieldId;
      const leftResolved = current.fields.find((field) => field.fieldId === leftField || field.name === leftField)?.fieldId;
      const rightResolved = right.fields.find((field) => field.fieldId === rightField || field.name === rightField)?.fieldId;
      if (!leftResolved || !rightResolved) throw new Error('Pivot relationship references an unknown field');
      const rows: SourceRow[] = [];
      for (const left of current.rows) {
        const matches = right.rows.filter((candidate) => same(left.values[leftResolved] ?? null, candidate.values[rightResolved] ?? null));
        if (!matches.length && relationship.join === 'left') rows.push(left);
        for (const match of matches) rows.push({ values: { ...left.values, ...match.values }, paths: [...left.paths, ...match.paths] });
      }
      current = { fields: [...current.fields, ...right.fields], rows };
    }
    return current;
  }
  const range = ranges[0]!;
  // Table field IDs come from the table model, so read the source columns with
  // their physical identities first and remap once below.
  const table = readRange(workbook.getSheet(range.sheetId), range, source, 0, source.kind === 'table' ? undefined : catalog);
  if (source.kind === 'table') {
    const stored = resolvePivotTable(workbook, source.tableId).fields;
    table.fields.forEach((field, index) => {
      const declared = stored[index];
      if (declared?.id) field.fieldId = declared.id;
      if (declared?.name) field.name = declared.name;
    });
    const remapped: SourceRow[] = table.rows.map((row) => {
      const values: Record<string, PivotScalar> = {};
      table.fields.forEach((field, index) => {
        const oldId = sourceIdentity(source, range, index);
        values[field.fieldId] = row.values[oldId] ?? null;
      });
      return { values, paths: row.paths };
    });
    table.rows = remapped;
  }
  return table;
}

function normalizeFieldCatalog(sourceTableValue: SourceTable, persisted?: PivotFieldCatalog): PivotFieldCatalog {
  const fields = sourceTableValue.fields.map((field, ordinal) => {
    const values = sourceTableValue.rows.map((row) => row.values[field.fieldId] ?? null);
    const persistedField = persisted?.fields.find((candidate) => candidate.ordinal === ordinal);
    const fieldId = persistedField?.fieldId ?? field.fieldId ?? `field:${ordinal}`;
    return { fieldId, name: field.name, dataType: inferType(values), ordinal, values: [...new Map(values.filter((value) => value != null).map((value) => [pivotMemberKey(createPivotMemberKey(value)), value])).values()].slice(0, 10_000) };
  });
  return { schema: 'PivotFieldCatalog', fields };
}

function resolveFieldId(reference: string | undefined, catalog: PivotFieldCatalog): string | undefined {
  if (!reference) return undefined;
  return catalog.fields.find((field) => field.fieldId === reference || field.name === reference)?.fieldId;
}

function fieldName(fieldId: string, catalog: PivotFieldCatalog): string {
  return catalog.fields.find((field) => field.fieldId === fieldId)?.name ?? fieldId;
}

function normalizePlacement(placement: PivotFieldPlacement, catalog: PivotFieldCatalog): PivotFieldPlacement {
  const fieldId = resolveFieldId(placement.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot field: ${placement.fieldId}`);
  const sort = placement.sort ? {
    ...placement.sort,
    ...(placement.sort.valueFieldId ? { valueFieldId: resolveFieldId(placement.sort.valueFieldId, catalog) } : {}),
  } : undefined;
  return { fieldId, sort, group: placement.group };
}

function normalizeFilter(filter: PivotFilter, catalog: PivotFieldCatalog): PivotFilter {
  const fieldId = resolveFieldId(filter.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot field: ${filter.fieldId}`);
  if (filter.kind === 'manual') {
    return { kind: 'manual', fieldId, mode: filter.mode, memberKeys: structuredClone(filter.memberKeys) };
  }
  if (filter.kind === 'top-items') {
    const valueFieldId = resolveFieldId(filter.valueFieldId, catalog);
    if (!valueFieldId) throw new Error(`Unknown pivot value field: ${filter.valueFieldId}`);
    return { ...filter, fieldId, valueFieldId };
  }
  return { ...filter, fieldId };
}

function normalizeValueField(field: PivotValueField, catalog: PivotFieldCatalog): PivotValueField {
  const fieldId = resolveFieldId(field.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot value field: ${field.fieldId}`);
  return { ...field, fieldId, ...(field.baseFieldId ? { baseFieldId: resolveFieldId(field.baseFieldId, catalog) } : {}) };
}

function normalizeLayout(layout: PivotLayout, catalog: PivotFieldCatalog): PivotLayout {
  return {
    ...structuredClone(layout),
    rows: layout.rows.map((entry) => normalizePlacement(entry, catalog)),
    columns: layout.columns.map((entry) => normalizePlacement(entry, catalog)),
    filters: layout.filters.map((entry) => normalizeFilter(entry, catalog)),
    values: layout.values.map((entry) => normalizeValueField(entry, catalog)),
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

/** Canonicalize field catalog values against the live source. Calculation has one model shape. */
export function normalizePivotDefinition(workbook: WorkbookModel, pivot: PivotModel): PivotDefinition {
  const source = getPivotSource(pivot);
  const fieldCatalog = source.kind === 'data-source'
    ? getPivotFieldCatalog(workbook, pivot)
    : normalizeFieldCatalog(sourceTable(workbook, pivot, pivot.fieldCatalog), pivot.fieldCatalog);
  const calculatedFields = [
    ...(pivot.layout.calculatedFields ?? []).map((field) => ({ fieldId: field.fieldId, name: field.name })),
    ...(pivot.layout.calculatedItems ?? []).map((field) => ({ fieldId: field.fieldId, name: field.name })),
  ];
  for (const calculated of calculatedFields) {
    if (!fieldCatalog.fields.some((field) => field.fieldId === calculated.fieldId || field.name === calculated.name)) {
      fieldCatalog.fields.push({ fieldId: calculated.fieldId, name: calculated.name, dataType: 'mixed', ordinal: fieldCatalog.fields.length, values: [] });
    }
  }
  return {
    schema: 'PivotDefinition',
    id: pivot.id,
    source,
    target: getPivotTarget(pivot),
    fieldCatalog,
    layout: normalizeLayout(pivot.layout, fieldCatalog),
    refreshPolicy: structuredClone(pivot.refreshPolicy),
    ...(pivot.nativeMetadata ? { nativeMetadata: structuredClone(pivot.nativeMetadata) } : {}),
  };
}

export function getPivotFieldCatalog(workbook: WorkbookModel, pivot: PivotModel): PivotFieldCatalog {
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
  return normalizeFieldCatalog(sourceTable(workbook, { ...pivot, source }, pivot.fieldCatalog), pivot.fieldCatalog);
}

function formulaScalar(value: FormulaValue): PivotScalar | null {
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

function rewriteCalculatedFormula(formula: string, fields: string[]): string {
  let rewritten = formula.trim().replace(/^=/, '');
  fields.map((field, index) => ({ field, index })).sort((left, right) => right.field.length - left.field.length).forEach(({ field, index }) => {
    const escaped = field.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const reference = `${columnLabel(index)}1`;
    rewritten = rewritten.replace(new RegExp(`\\[${escaped}\\]`, 'g'), reference);
    if (!formulaFunctions.has(field.toUpperCase())) rewritten = rewritten.replace(new RegExp(`(?<![A-Za-z0-9_])${escaped}(?![A-Za-z0-9_])`, 'g'), reference);
  });
  return `=${rewritten}`;
}

function calculateRowFormula(row: SourceRow, formula: string, fields: SourceField[]): PivotScalar | null {
  const engine = new FormulaEngine({ defaultSheetId: 'pivot' });
  fields.forEach((field, index) => engine.setValue({ sheetId: 'pivot', row: 0, column: index }, row.values[field.fieldId] ?? null));
  engine.setFormula({ sheetId: 'pivot', row: 1, column: 0 }, rewriteCalculatedFormula(formula, fields.map((field) => field.name)));
  return formulaScalar(engine.getCellValue({ sheetId: 'pivot', row: 1, column: 0 }));
}

function applyCalculatedData(rows: SourceRow[], fields: PivotFieldDefinition[], calculatedFields: PivotLayout['calculatedFields'] = [], calculatedItems: PivotLayout['calculatedItems'] = []): SourceRow[] {
  if (!calculatedFields.length && !calculatedItems.length) return rows;
  const currentFields: SourceField[] = fields.map((field) => ({ fieldId: field.fieldId, name: field.name, ordinal: field.ordinal, dataType: field.dataType }));
  return rows.map((row) => {
    const values = { ...row.values };
    for (const calculated of calculatedFields) {
      const fieldId = calculated.fieldId;
      values[fieldId] = calculateRowFormula({ ...row, values }, calculated.formula, currentFields);
      currentFields.push({ fieldId, name: calculated.name, ordinal: currentFields.length, dataType: 'mixed' });
    }
    for (const calculated of calculatedItems) {
      const fieldId = calculated.fieldId;
      const targetFieldId = resolveFieldId(calculated.targetFieldId, { fields: currentFields.map((field) => ({ fieldId: field.fieldId, name: field.name, ordinal: field.ordinal, dataType: field.dataType ?? 'mixed' })) });
      if (targetFieldId && same(values[targetFieldId] ?? null, calculated.name)) values[fieldId] = calculateRowFormula({ ...row, values }, calculated.formula, currentFields);
      if (!currentFields.some((field) => field.fieldId === fieldId)) currentFields.push({ fieldId, name: calculated.name, ordinal: currentFields.length, dataType: 'mixed' });
    }
    return { ...row, values };
  });
}

function toNumber(value: PivotScalar): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string' || value.trim() === '') return null;
  const parsed = Number(value.replace(/[$,%]/g, ''));
  return Number.isFinite(parsed) ? parsed : null;
}

function compare(left: PivotScalar, right: PivotScalar): number {
  if (same(left, right)) return 0;
  if (left == null || left === '') return -1;
  if (right == null || right === '') return 1;
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (leftNumber != null && rightNumber != null) return leftNumber - rightNumber;
  return String(left).localeCompare(String(right));
}

/** Every aggregate has its own semantics; no operation falls through to sum. */
export function aggregatePivotValues(rows: ReadonlyArray<{ values: Record<string, PivotScalar> }>, fieldId: string, operation: PivotAggregateFunction): number | null {
  const numbers: number[] = [];
  const members = new Set<string>();
  let nonBlank = 0;
  for (const row of rows) {
    const raw = row.values[fieldId] ?? null;
    if (raw != null && raw !== '') nonBlank += 1;
    if (raw != null && raw !== '') members.add(pivotMemberKey(createPivotMemberKey(raw)));
    const number = toNumber(raw);
    if (number != null) numbers.push(number);
  }
  switch (operation) {
    case 'count': return nonBlank;
    case 'count-numbers': return numbers.length;
    case 'sum': return numbers.reduce((sum, value) => sum + value, 0);
    case 'average': return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null;
    case 'min': return numbers.length ? Math.min(...numbers) : null;
    case 'max': return numbers.length ? Math.max(...numbers) : null;
    case 'product': return numbers.length ? numbers.reduce((product, value) => product * value, 1) : null;
    case 'distinct-count': return members.size;
    case 'stdev': {
      if (numbers.length < 2) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return Math.sqrt(numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - 1));
    }
    case 'stdevp': {
      if (!numbers.length) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return Math.sqrt(numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / numbers.length);
    }
    case 'var': {
      if (numbers.length < 2) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - 1);
    }
    case 'varp': {
      if (!numbers.length) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / numbers.length;
    }
    default: return assertNever(operation);
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
    const number = toNumber(value);
    if (number == null) return value;
    if (!Number.isFinite(group.interval) || group.interval <= 0) throw new Error('Pivot number grouping interval must be positive');
    const start = group.start ?? 0;
    const result = start + Math.floor((number - start) / group.interval) * group.interval;
    return group.end !== undefined && result > group.end ? group.end : result;
  }
  const date = new Date(String(value));
  if (Number.isNaN(date.getTime())) return value;
  if (group.unit === 'year') return date.getFullYear();
  if (group.unit === 'quarter') return `${date.getFullYear()} Q${Math.floor(date.getMonth() / 3) + 1}`;
  if (group.unit === 'month') return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  if (group.unit === 'week') return Math.ceil((((date.getTime() - new Date(date.getFullYear(), 0, 1).getTime()) / 86400000) + date.getDay() + 1) / 7);
  return date.toISOString().slice(0, 10);
}

function axisGroups(rows: SourceRow[], placements: PivotFieldPlacement[]): AxisGroup[] {
  const map = new Map<string, AxisGroup>();
  for (const row of rows) {
    const values = placements.map((placement) => grouped(row.values[placement.fieldId] ?? null, placement.group));
    const key = JSON.stringify(values.map(createPivotMemberKey));
    const group = map.get(key) ?? { values, rows: [] };
    group.rows.push(row);
    map.set(key, group);
  }
  const placement = placements[placements.length - 1];
  const result = [...map.values()].sort((left, right) => {
    if (placement?.sort?.by === 'value' && placement.sort.valueFieldId) {
      return (aggregatePivotValues(left.rows, placement.sort.valueFieldId, 'sum') ?? 0) - (aggregatePivotValues(right.rows, placement.sort.valueFieldId, 'sum') ?? 0);
    }
    for (let index = 0; index < left.values.length; index += 1) {
      const order = compare(left.values[index] ?? null, right.values[index] ?? null);
      if (order) return order;
    }
    return 0;
  });
  if (placement?.sort?.direction === 'descending') result.reverse();
  return result;
}

function manualFilterMatches(value: PivotScalar, filter: Extract<PivotFilter, { kind: 'manual' }>): boolean {
  if (filter.mode === 'all') return true;
  const key = createPivotMemberKey(value);
  const included = (filter.memberKeys ?? []).some((candidate) => pivotMemberKeyEquals(candidate, key));
  return filter.mode === 'include' ? included : !included;
}

function matchesFilter(row: SourceRow, filter: PivotFilter): boolean {
  const fieldId = filter.fieldId;
  const value = row.values[fieldId] ?? null;
  if (filter.kind === 'top-items') return true;
  if (filter.kind === 'manual') return manualFilterMatches(value, filter);
  const leftNumber = toNumber(value);
  const rightNumber = toNumber(filter.value);
  const order = leftNumber != null && rightNumber != null ? leftNumber - rightNumber : compare(value, filter.value);
  switch (filter.operator) {
    case 'equals': return same(value, filter.value);
    case 'not-equals': return !same(value, filter.value);
    case 'contains': return String(value ?? '').includes(String(filter.value ?? ''));
    case 'greater-than': return order > 0;
    case 'greater-or-equal': return order >= 0;
    case 'less-than': return order < 0;
    case 'less-or-equal': return order <= 0;
    default: return assertNever(filter.operator);
  }
}

function topItems(rows: SourceRow[], filters: PivotFilter[]): SourceRow[] {
  let result = rows;
  for (const filter of filters) {
    if (filter.kind !== 'top-items') continue;
    const fieldId = filter.fieldId;
    const valueFieldId = filter.valueFieldId;
    if (!Number.isInteger(filter.count) || filter.count < 1) throw new Error('Pivot top-items count must be a positive integer');
    const buckets = new Map<string, SourceRow[]>();
    for (const row of result) {
      const key = pivotMemberKey(createPivotMemberKey(row.values[fieldId] ?? null));
      const bucket = buckets.get(key) ?? [];
      bucket.push(row);
      buckets.set(key, bucket);
    }
    const ranked = [...buckets.values()].sort((left, right) => (aggregatePivotValues(left, valueFieldId, 'sum') ?? 0) - (aggregatePivotValues(right, valueFieldId, 'sum') ?? 0));
    if (filter.direction === 'top') ranked.reverse();
    result = ranked.slice(0, filter.count).flat();
  }
  return result;
}

function matchesSlicer(row: SourceRow, slicer: PivotSlicerDrawingPayload): boolean {
  const { fieldId, filter } = slicer;
  if (filter.mode === 'all') return true;
  const included = filter.memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, createPivotMemberKey(row.values[fieldId] ?? null)));
  return filter.mode === 'include' ? included : !included;
}

function matchesTimeline(row: SourceRow, timeline: PivotTimelineDrawingPayload): boolean {
  const fieldId = timeline.fieldId;
  const raw = row.values[fieldId];
  if (raw == null || raw === '') return false;
  const date = new Date(String(raw));
  if (Number.isNaN(date.getTime())) return false;
  const start = timeline.period.start ? new Date(timeline.period.start).getTime() : Number.NEGATIVE_INFINITY;
  const end = timeline.period.end ? new Date(timeline.period.end).getTime() : Number.POSITIVE_INFINITY;
  return date.getTime() >= start && date.getTime() <= end;
}

function matchesControls(workbook: WorkbookModel, rows: SourceRow[], pivot: PivotModel): SourceRow[] {
  const slicers: PivotSlicerDrawingPayload[] = [];
  const timelines: PivotTimelineDrawingPayload[] = [];
  for (const sheet of workbook.getSheets()) {
    for (const drawing of sheet.drawings) {
      if (drawing.kind !== 'slicer' && drawing.kind !== 'timeline') continue;
      const payload = sheet.drawingPayloads.get(drawing.payloadId);
      if (!payload || (payload.kind !== 'slicer' && payload.kind !== 'timeline')) continue;
      if (![payload.pivotId, ...(payload.connectedPivotIds ?? [])].includes(pivot.id)) continue;
      if (payload.kind === 'slicer') slicers.push(payload);
      else if (payload.kind === 'timeline') timelines.push(payload);
    }
  }
  return rows.filter((row) => slicers.every((slicer) => matchesSlicer(row, slicer)) && timelines.every((timeline) => matchesTimeline(row, timeline)));
}

function resultCells(rows: SourceRow[], columns: AxisGroup[], values: PivotValueField[], nodePath: string[], kind: PivotResultCell['kind'] = 'detail'): PivotResultCell[] {
  return columns.map((column, columnIndex) => {
    const nodeRows = new Set(rows);
    const columnRows = column.rows.filter((candidate) => nodeRows.has(candidate));
    return {
      id: `${nodePath.join('/') || 'root'}|column:${columnIndex}`,
      nodePath,
      kind,
      columnPath: column.values,
      sourceRowPaths: columnRows.flatMap((row) => row.paths),
      values: values.map((value) => aggregatePivotValues(columnRows, value.fieldId, value.summarizeBy)),
    };
  });
}

function resultNodes(rows: SourceRow[], placements: PivotFieldPlacement[], depth: number, columns: AxisGroup[], values: PivotValueField[], showSubtotals: boolean, prefix: string[] = []): PivotResultNode[] {
  if (depth >= placements.length) return [];
  const placement = placements[depth]!;
  return axisGroups(rows, [placement]).map((group) => {
    const fieldId = placement.fieldId;
    const member = createPivotMemberKey(group.values[0] ?? null);
    const path = [...prefix, `${fieldId}=${pivotMemberKey(member)}`];
    const children = resultNodes(group.rows, placements, depth + 1, columns, values, showSubtotals, path);
    const leaf = children.length === 0;
    return {
      nodeId: path.join('/'),
      path,
      kind: leaf ? 'leaf' : 'subtotal',
      fieldId,
      memberKey: member,
      key: group.values[0] ?? null,
      label: display(group.values[0] ?? null),
      depth,
      children,
      values: resultCells(group.rows, columns, values, path, showSubtotals && !leaf ? 'subtotal' : 'detail'),
      subtotal: showSubtotals && !leaf,
      sourceRowPaths: group.rows.flatMap((row) => row.paths),
    };
  });
}

function applyShowAs(tree: PivotResultTree, fields: PivotValueField[]): void {
  const leaves: PivotResultNode[] = [];
  const collectLeaves = (nodes: PivotResultNode[]) => nodes.forEach((node) => node.children.length ? collectLeaves(node.children) : leaves.push(node));
  collectLeaves(tree.rows);
  const raw = new Map<PivotResultCell, PivotScalar[]>();
  const snapshot = (nodes: PivotResultNode[]) => nodes.forEach((node) => { node.values.forEach((cell) => raw.set(cell, [...cell.values])); snapshot(node.children); });
  snapshot(tree.rows);
  const rawValue = (cell: PivotResultCell | undefined, index: number): PivotScalar | null => cell ? raw.get(cell)?.[index] ?? null : null;
  const grandValues = tree.grandTotal?.values.map((value) => value) ?? [];
  const visit = (nodes: PivotResultNode[], parent?: PivotResultNode) => nodes.forEach((node) => {
    node.values.forEach((cell, columnIndex) => fields.forEach((field, valueIndex) => {
      const spec = field.showAs ?? { kind: 'normal' as const };
      const current = toNumber(rawValue(cell, valueIndex));
      if (current == null || spec.kind === 'normal') return;
      const grand = toNumber(grandValues[valueIndex] ?? null);
      const rowTotal = node.values.reduce((sum, item) => sum + (toNumber(rawValue(item, valueIndex)) ?? 0), 0);
      const columnTotal = leaves.reduce((sum, item) => sum + (toNumber(rawValue(item.values[columnIndex], valueIndex)) ?? 0), 0);
      const parentTotal = parent ? toNumber(rawValue(parent.values[columnIndex], valueIndex)) : null;
      if (spec.kind === 'grand-percentage') cell.values[valueIndex] = grand ? current / grand : null;
      else if (spec.kind === 'row-percentage') cell.values[valueIndex] = rowTotal ? current / rowTotal : null;
      else if (spec.kind === 'column-percentage') cell.values[valueIndex] = columnTotal ? current / columnTotal : null;
      else if (spec.kind === 'parent-percentage') cell.values[valueIndex] = parentTotal ? current / parentTotal : null;
      else if (spec.kind === 'difference' || spec.kind === 'percentage-difference') {
        const base = spec.base === 'grand' ? grand : spec.base === 'row' ? rowTotal : spec.base === 'column' ? columnTotal : parentTotal;
        cell.values[valueIndex] = base == null ? null : spec.kind === 'difference' ? current - base : base ? (current - base) / base : null;
      } else if (spec.kind === 'running-total') {
        if (spec.axis === 'row') {
          const end = leaves.indexOf(node);
          cell.values[valueIndex] = leaves.slice(0, end + 1).reduce((sum, item) => sum + (toNumber(rawValue(item.values[columnIndex], valueIndex)) ?? 0), 0);
        } else {
          const end = columnIndex;
          cell.values[valueIndex] = node.values.slice(0, end + 1).reduce((sum, item) => sum + (toNumber(rawValue(item, valueIndex)) ?? 0), 0);
        }
      } else if (spec.kind === 'rank') {
        const series = spec.axis === 'row' ? leaves.map((item) => toNumber(rawValue(item.values[columnIndex], valueIndex))) : node.values.map((item) => toNumber(rawValue(item, valueIndex)));
        const ranked = series.filter((value): value is number => value != null).sort((left, right) => spec.direction === 'ascending' ? left - right : right - left);
        cell.values[valueIndex] = ranked.indexOf(current) + 1;
      } else if (spec.kind === 'index') {
        cell.values[valueIndex] = grand != null && rowTotal && columnTotal ? current * grand / rowTotal / columnTotal : null;
      }
    }));
    visit(node.children, node);
  });
  visit(tree.rows);
}

function computePivotResultFromTable(
  workbook: WorkbookModel,
  pivot: PivotModel,
  definition: PivotDefinition,
  rawTable: PivotSourceTableInput,
  sourceRevisionOverride?: string,
): PivotResultTree {
  const rows = applyCalculatedData(rawTable.rows, definition.fieldCatalog.fields, definition.layout.calculatedFields, definition.layout.calculatedItems);
  const references = [
    ...definition.layout.rows.map((entry) => entry.fieldId),
    ...definition.layout.columns.map((entry) => entry.fieldId),
    ...definition.layout.filters.flatMap((filter) => filter.kind === 'top-items' ? [filter.fieldId, filter.valueFieldId] : [filter.fieldId]),
    ...definition.layout.values.map((entry) => entry.fieldId),
  ];
  const known = new Set([...definition.fieldCatalog.fields.map((field) => field.fieldId), ...(definition.layout.calculatedFields ?? []).map((field) => field.fieldId), ...(definition.layout.calculatedItems ?? []).map((field) => field.fieldId)]);
  const unknown = references.find((field) => field && !known.has(field));
  if (unknown && rawTable.fields.length) throw new Error(`Unknown pivot field: ${unknown}`);
  let filtered = matchesControls(workbook, rows, pivot);
  filtered = filtered.filter((row) => definition.layout.filters.filter((filter) => filter.kind !== 'top-items').every((filter) => matchesFilter(row, filter)));
  filtered = topItems(filtered, definition.layout.filters);
  const columns = definition.layout.columns.length ? axisGroups(filtered, definition.layout.columns) : [{ values: [], rows: filtered }];
  const grandTotal: PivotResultCell | null = definition.layout.showGrandTotals ? {
    id: `${definition.id}|grand-total`,
    kind: 'grand-total',
    columnPath: [],
    values: definition.layout.values.map((field) => aggregatePivotValues(filtered, field.fieldId, field.summarizeBy)),
    sourceRowPaths: filtered.flatMap((row) => row.paths),
  } : null;
  const tree: PivotResultTree = {
    schema: PIVOT_RESULT_TREE_SCHEMA,
    pivotId: definition.id,
    fields: definition.fieldCatalog,
    columnPaths: columns.map((column) => column.values),
    rows: resultNodes(filtered, definition.layout.rows, 0, columns, definition.layout.values, definition.layout.showSubtotals),
    grandTotal,
    sourceRowPaths: filtered.flatMap((row) => row.paths),
  };
  applyShowAs(tree, definition.layout.values);
  const revisions = getPivotRevisionKey(workbook, pivot);
  tree.sourceRevision = sourceRevisionOverride ?? revisions.sourceRevision;
  tree.layoutRevision = revisions.layoutRevision;
  tree.filterRevision = revisions.filterRevision;
  return tree;
}

function computePivotResultUncached(workbook: WorkbookModel, pivot: PivotModel): PivotResultTree {
  const definition = normalizePivotDefinition(workbook, pivot);
  const rawTable = sourceTable(workbook, pivot, definition.fieldCatalog);
  return computePivotResultFromTable(workbook, pivot, definition, rawTable);
}

export function computePivotResult(workbook: WorkbookModel, pivot: PivotModel): PivotResultTree {
  return structuredClone(computePivotResultUncached(workbook, pivot));
}

/** Apply the normal Pivot calculation pipeline to asynchronously loaded block data. */
export function computePivotResultFromBlockSource(
  workbook: WorkbookModel,
  pivot: PivotModel,
  source: PivotSourceTableInput,
  sourceRevision: string,
): PivotResultTree {
  const definition = normalizePivotDefinition(workbook, pivot);
  if (definition.source.kind !== 'data-source') throw new Error('Block source calculation requires a data-source Pivot');
  return structuredClone(computePivotResultFromTable(workbook, pivot, definition, source, sourceRevision));
}

function nodeVisible(node: PivotResultNode, layout: PivotLayout, ancestorsVisible: boolean): boolean {
  if (!ancestorsVisible) return false;
  const expansion = layout.expansion;
  if (!expansion || !node.children.length) return true;
  if (expansion.collapsedNodeIds.includes(node.nodeId ?? '')) return false;
  // An empty expanded set means the default Excel state (all nodes expanded);
  // once the user explicitly records paths, only those paths remain open.
  return expansion.expandedNodeIds.length === 0 || expansion.expandedNodeIds.includes(node.nodeId ?? '') || node.depth === 0;
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
    const visible = nodeVisible(node, layout, parentVisible);
    output.push({ node, labels: currentLabels, visible });
    if (visible && node.children.length) output.push(...flattenNodes(node.children, layout, currentLabels, true));
  }
  return output;
}

function textForValue(value: PivotScalar): string {
  return value == null ? '' : String(value);
}

function projectionCell(pivotId: string, row: number, column: number, kind: PivotProjectionCell['kind'], value: PivotScalar, text: string, extra: Partial<PivotProjectionCell> = {}): PivotProjectionCell {
  return { id: `${pivotId}|r${row}|c${column}`, pivotId, row, column, kind, value, text, ...extra };
}

function projectionRange(target: PivotTarget, rowCount: number, columnCount: number): RangeRef {
  return { sheetId: target.sheetId, startRow: target.anchor.row, endRow: target.anchor.row + Math.max(rowCount - 1, 0), startColumn: target.anchor.column, endColumn: target.anchor.column + Math.max(columnCount - 1, 0) };
}

export function detectPivotCollision(workbook: WorkbookModel, pivot: PivotModel, range: RangeRef): import('@react-sheets/core-model').PivotCollision {
  const sheet = workbook.getSheet(range.sheetId);
  const reasons = new Set<import('@react-sheets/core-model').PivotCollision['reasons'][number]>();
  const conflictingRanges: RangeRef[] = [];
  if (range.endRow >= sheet.rowCount || range.endColumn >= sheet.columnCount) reasons.add('worksheet-bounds');
  sheet.cells.forEach((_cell, row, column) => {
    if (row >= range.startRow && row <= range.endRow && column >= range.startColumn && column <= range.endColumn) reasons.add('cell-data');
  });
  for (const merge of sheet.merges) {
    if (rangesIntersect(range, merge.range)) {
      reasons.add('merge');
      conflictingRanges.push(structuredClone(merge.range));
    }
  }
  for (const candidate of sheet.pivots) {
    if (candidate.id === pivot.id) continue;
    const target = getPivotTarget(candidate);
    if (target.sheetId !== range.sheetId) continue;
    const candidateRange = { sheetId: target.sheetId, startRow: target.anchor.row, endRow: target.anchor.row, startColumn: target.anchor.column, endColumn: target.anchor.column };
    if (rangesIntersect(range, candidateRange)) {
      reasons.add('pivot');
      conflictingRanges.push(candidateRange);
    }
  }
  return { status: reasons.size ? 'collision' : 'clear', reasons: [...reasons], conflictingRanges };
}

function rangesIntersect(left: RangeRef, right: RangeRef): boolean {
  return left.sheetId === right.sheetId && left.startRow <= right.endRow && right.startRow <= left.endRow && left.startColumn <= right.endColumn && right.startColumn <= left.endColumn;
}

function refreshState(workbook: WorkbookModel, pivot: PivotModel, collision: import('@react-sheets/core-model').PivotCollision, status: PivotRefreshState['status'] = 'ready', error?: string): PivotRefreshState {
  const revisions = getPivotRevisionKey(workbook, pivot);
  return {
    status: collision.status === 'collision' ? 'collision' : status,
    revision: Number.parseInt(revisions.sourceRevision.slice(-6), 16) || 0,
    sourceRevision: revisions.sourceRevision,
    completedAt: new Date().toISOString(),
    ...(error ? { error } : {}),
  };
}

export function getPivotRefreshState(workbook: WorkbookModel, pivot: PivotModel, collision?: import('@react-sheets/core-model').PivotCollision, status: PivotRefreshState['status'] = 'ready', error?: string): PivotRefreshState {
  const effectiveCollision = collision ?? { status: 'clear' as const, reasons: [], conflictingRanges: [] };
  return refreshState(workbook, pivot, effectiveCollision, status, error);
}

/** Build one candidate worksheet overlay. It returns cells only; no workbook cell is mutated. */
function buildPivotGridProjectionCandidate(
  workbook: WorkbookModel,
  pivot: PivotModel,
  cachedResult?: PivotResultTree,
  options: PivotProjectionOptions = {},
): PivotGridProjection {
  const definition = normalizePivotDefinition(workbook, pivot);
  const target = definition.target;
  let tree: PivotResultTree | undefined = cachedResult;
  let error: string | undefined;
  let loading = false;
  const sourceState = options.sourceState;
  if (!tree && definition.source.kind === 'data-source') {
    if (sourceState?.availability === 'error' || sourceState?.availability === 'missing') {
      error = sourceState.error ?? `PivotTable source ${sourceState.availability}`;
    } else {
      // Block-backed sources are asynchronous by contract. Never turn the
      // intentional sync boundary error into a red error projection.
      loading = true;
    }
  } else if (!tree) {
    try {
      tree = computePivotResult(workbook, pivot);
    } catch (cause) {
      error = cause instanceof Error ? cause.message : String(cause);
    }
  }
  if (tree && (sourceState?.availability === 'error' || sourceState?.availability === 'missing')) {
    error = sourceState.error ?? `PivotTable source ${sourceState.availability}`;
  } else if (tree && sourceState?.availability === 'loading') {
    loading = true;
  }
  const cells: PivotProjectionCell[] = [];
  const rowHeaderCount = Math.max(definition.layout.rows.length, 1);
  const valueColumnCount = Math.max(tree ? tree.columnPaths.length * definition.layout.values.length : definition.layout.values.length, 1);
  let row = 0;
  cells.push(projectionCell(definition.id, row, 0, 'title', definition.id, definition.id));
  row += 1;
  for (const filter of definition.layout.filters) {
    cells.push(projectionCell(definition.id, row, 0, 'filter', filter.fieldId, `${fieldName(filter.fieldId, definition.fieldCatalog)}: All`));
    row += 1;
  }
  for (let index = 0; index < rowHeaderCount; index += 1) cells.push(projectionCell(definition.id, row, index, 'column-header', null, index === 0 ? 'Row Labels' : ''));
  const columnPaths = tree?.columnPaths ?? [];
  const values = definition.layout.values;
  for (let columnIndex = 0; columnIndex < Math.max(columnPaths.length, 1); columnIndex += 1) {
    const path = columnPaths[columnIndex] ?? [];
    for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
      const column = rowHeaderCount + columnIndex * Math.max(values.length, 1) + valueIndex;
      const valueField = values[valueIndex];
      const label = path.length ? `${path.map(display).join(' / ')} ${valueField ? (valueField.displayName ?? valueField.fieldId) : ''}`.trim() : valueField ? (valueField.displayName ?? valueField.fieldId) : '';
      cells.push(projectionCell(definition.id, row, column, 'column-header', path[0] ?? null, label, { columnPath: path }));
    }
  }
  row += 1;
  if (tree) {
    const flat = flattenNodes(tree.rows, definition.layout);
    for (const item of flat) {
      if (!item.visible) continue;
      const node = item.node;
      for (let axis = 0; axis < rowHeaderCount; axis += 1) {
        const label = item.labels[axis] ?? (axis === 0 ? node.label : '');
        const kind: PivotProjectionCell['kind'] = axis === 0 && node.children.length && definition.layout.expansion?.showButtons ? 'expand-toggle' : node.subtotal ? 'subtotal' : 'row-header';
        cells.push(projectionCell(definition.id, row, axis, kind, axis === 0 ? node.key : null, label, { nodeId: node.nodeId, expandable: node.children.length > 0, expanded: node.children.length > 0 && !definition.layout.expansion?.collapsedNodeIds.includes(node.nodeId ?? '') }));
      }
      for (let columnIndex = 0; columnIndex < Math.max(columnPaths.length, 1); columnIndex += 1) {
        const resultCell = node.values[columnIndex];
        for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
          const column = rowHeaderCount + columnIndex * Math.max(values.length, 1) + valueIndex;
          const value = resultCell?.values[valueIndex] ?? null;
          cells.push(projectionCell(definition.id, row, column, node.subtotal ? 'subtotal' : 'value', value, textForValue(value), { nodeId: node.nodeId, resultCellId: resultCell?.id, columnPath: resultCell?.columnPath, sourceRowPaths: resultCell?.sourceRowPaths }));
        }
      }
      row += 1;
    }
    if (tree.grandTotal) {
      cells.push(projectionCell(definition.id, row, 0, 'grand-total', null, 'Grand Total', { resultCellId: tree.grandTotal.id, sourceRowPaths: tree.grandTotal.sourceRowPaths }));
      tree.grandTotal.values.forEach((value, index) => cells.push(projectionCell(definition.id, row, rowHeaderCount + index, 'grand-total', value, textForValue(value), { resultCellId: tree.grandTotal?.id, sourceRowPaths: tree.grandTotal?.sourceRowPaths })));
      row += 1;
    }
  } else {
    cells.push(projectionCell(definition.id, row, 0, error ? 'error' : 'loading', null, error ?? 'Loading PivotTable'));
    row += 1;
  }
  const occupiedRange = projectionRange(target, Math.max(row, 1), Math.max(valueColumnCount + rowHeaderCount, 1));
  const collision = detectPivotCollision(workbook, pivot, occupiedRange);
  return {
    schema: PIVOT_GRID_PROJECTION_SCHEMA,
    pivotId: definition.id,
    sheetId: target.sheetId,
    target,
    occupiedRange,
    cells,
    collision,
    refresh: refreshState(workbook, pivot, collision, error ? 'error' : loading ? 'refreshing' : tree ? 'ready' : 'refreshing', error),
  };
}

function projectionWithStatus(
  workbook: WorkbookModel,
  pivot: PivotModel,
  entry: LastValidPivotProjection,
  collision: import('@react-sheets/core-model').PivotCollision,
  status: PivotRefreshState['status'],
  error?: string,
): PivotGridProjection {
  const projection = structuredClone(entry.projection);
  projection.collision = structuredClone(collision);
  projection.refresh = refreshState(workbook, pivot, collision, status, error);
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
  let effectiveResult = cachedResult;
  if (!effectiveResult && pivot.source.kind !== 'data-source') {
    try {
      effectiveResult = computePivotResult(workbook, pivot);
    } catch {
      // The candidate builder creates the explicit synchronous error state.
    }
  }
  const candidate = buildPivotGridProjectionCandidate(workbook, pivot, effectiveResult, options);
  const cache = lastValidPivotProjections.get(workbook);
  const last = cache?.get(pivot.id);
  const candidateTree = effectiveResult;

  if (candidate.collision.status === 'clear' && candidateTree && candidate.refresh.status === 'ready') {
    const nextCache = cache ?? new Map<string, LastValidPivotProjection>();
    nextCache.set(pivot.id, { projection: structuredClone(candidate), result: structuredClone(candidateTree) });
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
      );
    }
    return projectionWithStatus(workbook, pivot, last, retainedCollision, candidate.refresh.status, candidate.refresh.error);
  }

  return candidate;
}

export function hitTestPivotProjection(projection: PivotGridProjection, row: number, column: number): PivotHitTest {
  const cell = projection.cells.find((candidate) => candidate.row === row && candidate.column === column);
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

export function resolvePivotContextHit(projection: PivotGridProjection, row: number, column: number): ContextHit {
  return { ...hitTestPivotProjection(projection, row, column), context: 'pivot', priority: 30 };
}

export function computePivotTable(workbook: WorkbookModel, pivot: PivotModel): PivotResultTable {
  const tree = computePivotResult(workbook, pivot);
  const definition = normalizePivotDefinition(workbook, pivot);
  const rows = tree.rows.map((node) => ({ keys: [node.label], values: node.values.flatMap((cell) => cell.values) }));
  const headers = [
    ...definition.layout.rows.map((field) => fieldName(field.fieldId, definition.fieldCatalog)),
    ...tree.columnPaths.flatMap((path) => definition.layout.values.map((field) => path.length ? `${path.map(display).join(' / ')} ${field.displayName ?? fieldName(field.fieldId, definition.fieldCatalog)}` : field.displayName ?? fieldName(field.fieldId, definition.fieldCatalog))),
  ];
  return { headers, rows, grandTotal: tree.grandTotal?.values ?? [], tree };
}

function pivotSourceRangesForExport(workbook: WorkbookModel, pivot: PivotModel): RangeRef[] {
  return sourceRanges(workbook, pivot);
}

export function getPivotSourceRanges(workbook: WorkbookModel, pivot: PivotModel): RangeRef[] {
  return structuredClone(pivotSourceRangesForExport(workbook, pivot));
}
