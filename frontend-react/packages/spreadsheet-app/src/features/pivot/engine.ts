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
  createPivotCollator,
  createPivotMemberKey,
  formatPivotMember,
  isPivotError,
  normalizePivotTimelinePeriod,
  pivotMemberKey,
  normalizePivotRefreshPolicy,
  normalizePivotDisplayOptions,
  normalizePivotNumberFormat,
  pivotTimelineInstant,
  pivotMemberKeyEquals,
  pivotScalarFromMemberKey,
} from '@react-sheets/core-model';
import type { PivotTimelinePeriodBounds } from '@react-sheets/core-model';
import { FormulaEngine, isFormulaError, type FormulaValue } from '@react-sheets/formula-engine';
import { formatValue as formatNumberValue } from '@react-sheets/number-format';
import { configureWorkbookSpillEnvironments, syncWorkbookSheetTables } from '../../formula-spill-sync';

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
  /** The session's canonical FormulaEngine; required for live spill values. */
  formula?: FormulaEngine;
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
  return fingerprint({
    source: canonicalPivotSource(source),
    revisions,
    ...(formula ? {
      formulaGeneration: formula.getCalculationGeneration(),
      spills: ranges.map((range) => formula.getSpillsForSheet(range.sheetId)
        .map((spill) => ({ anchor: spill.anchor, range: spill.range, state: spill.state }))),
    } : {}),
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
    const linked = [payload.pivotId, ...(payload.connectedPivotIds ?? [])];
    return linked.includes(pivot.id) ? { drawingId: drawing.id, payload } : undefined;
  })).filter((entry): entry is NonNullable<typeof entry> => Boolean(entry));
}

export function getPivotRevisionKey(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotRevisionKey {
  return {
    pivotId: pivot.id,
    sourceRevision: sourceRevision(workbook, pivot, formula),
    // Live member values belong exclusively to sourceRevision. Including them
    // here made an ordinary source edit look like a layout mutation and caused
    // manual-refresh PivotTables to discard their last refreshed result.
    layoutRevision: fingerprint({
      source: canonicalPivotSource(pivot.source),
      fieldCatalog: pivot.fieldCatalog.fields.map(({ fieldId, name, dataType, ordinal }) => ({ fieldId, name, dataType, ordinal })),
      layout: pivot.layout,
      presentation: pivot.presentation,
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
    if (!manifest.sourceRange) throw new Error(`Pivot data source ${source.dataSourceId} has no worksheet range`);
    return [manifest.sourceRange];
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

function inferType(values: PivotScalar[]): PivotFieldDataType {
  const present = values.filter((value) => value != null && value !== '');
  if (!present.length) return 'mixed';
  if (present.every(isPivotError)) return 'error';
  if (present.some(isPivotError)) return 'mixed';
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
  const rows: SourceRow[] = [];
  const sourceId = source.kind === 'worksheet-ranges' ? source.ranges[rangeIndex]?.sourceId : undefined;
  for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
    const values: Record<string, PivotScalar> = {};
    fields.forEach((field, ordinal) => {
      const cell = sheet.cells.get(row, range.startColumn + ordinal);
      const raw = formula
        ? formulaCellValue(formula, { sheetId: sheet.id, row, column: range.startColumn + ordinal }, cell?.formulaValue ?? cell?.value ?? null)
        : cell?.formulaValue ?? cell?.value ?? null;
      values[field.fieldId] = cellScalar(raw);
    });
    rows.push({ values, paths: [{ ...(sourceId ? { sourceId } : {}), recordId: `${sourceId ?? range.sheetId}:${row}`, sheetId: range.sheetId, row }] });
  }
  return { fields, rows };
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
  return inferType(table.rows.map((row) => row.values[fieldId] ?? null));
}

function joinKey(value: PivotScalar): string {
  return pivotMemberKey(createPivotMemberKey(value));
}

function assertUniqueLookupKeys(table: SourceTable, fieldId: string, sourceId: string): void {
  const keys = new Set<string>();
  for (const row of table.rows) {
    const key = joinKey(row.values[fieldId] ?? null);
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
  for (const row of attached.rows) lookup.set(joinKey(row.values[attachedFieldId] ?? null), row);
  const rows: SourceRow[] = [];
  for (const left of current.rows) {
    const match = lookup.get(joinKey(left.values[currentFieldId] ?? null));
    if (!match) {
      if (join === 'left') rows.push(left);
      continue;
    }
    const recordId = left.paths[0]?.recordId ?? match.paths[0]?.recordId;
    rows.push({
      values: { ...left.values, ...match.values },
      paths: [...left.paths, ...match.paths].map((path) => ({ ...path, ...(recordId ? { recordId } : {}) })),
    });
  }
  return { fields: [...current.fields, ...attached.fields], rows };
}

function sourceTable(workbook: WorkbookModel, pivot: PivotModel, catalog?: PivotFieldCatalog, formula?: FormulaEngine): SourceTable {
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
    const persistedField = persisted?.fields.find((candidate) => candidate.fieldId === field.fieldId);
    const fieldId = persistedField?.fieldId ?? field.fieldId ?? `field:${ordinal}`;
    const members = [...new Map(values.map((value) => {
      // Empty text and null are one semantic blank member. Keep that identity
      // in the catalogue even when the UI preview reaches its display bound.
      const canonical = value === '' ? null : value;
      return [pivotMemberKey(createPivotMemberKey(canonical)), canonical] as const;
    })).values()];
    const preview = members.slice(0, 10_000);
    const blank = members.find((value) => createPivotMemberKey(value).type === 'blank');
    if (blank !== undefined && !preview.some((value) => pivotMemberKey(createPivotMemberKey(value)) === pivotMemberKey(createPivotMemberKey(blank)))) preview.push(blank);
    return { fieldId, name: field.name, dataType: inferType(values), ordinal, values: preview };
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

/**
 * Build the canonical, locale-independent summary for one report field.
 * Report filters are grouped by field so allowing multiple filter families
 * cannot silently select one family and render the others as `All`.
 */
export function summarizePivotReportFilters(
  filters: readonly PivotFilter[],
  catalog: PivotFieldCatalog,
  fieldId: string,
): PivotReportFilterSummary {
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
        count: filter.count,
        direction: filter.direction,
        valueFieldName: fieldName(filter.valueFieldId, catalog),
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
      ...(filter.valueFieldId === undefined ? {} : { valueFieldName: fieldName(filter.valueFieldId, catalog) }),
    } as PivotReportFilterSummaryEntry;
  });
  return {
    fieldName: fieldName(fieldId, catalog),
    active: entries.some((entry) => entry.active),
    entries,
  };
}

function normalizePlacement(placement: PivotFieldPlacement, catalog: PivotFieldCatalog, valueFieldIds: ReadonlySet<string>): PivotFieldPlacement {
  const fieldId = resolveFieldId(placement.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot field: ${placement.fieldId}`);
  let sort: PivotSort | undefined;
  if (placement.sort) {
    if (placement.sort.by === 'value') {
      const valueFieldId = resolveFieldId(placement.sort.valueFieldId, catalog);
      if (!valueFieldId) throw new Error(`Pivot value sort requires a valueFieldId for ${fieldId}`);
      if (!valueFieldIds.has(valueFieldId)) throw new Error(`Pivot value sort field is not in Values: ${valueFieldId}`);
      sort = { ...placement.sort, valueFieldId };
    } else {
      sort = { direction: placement.sort.direction, by: 'label' };
    }
  }
  return { fieldId, sort, group: placement.group, subtotal: placement.subtotal ? structuredClone(placement.subtotal) : undefined };
}

function normalizeFilter(filter: PivotFilter, catalog: PivotFieldCatalog): PivotFilter {
  const fieldId = resolveFieldId(filter.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot field: ${filter.fieldId}`);
  if (filter.kind === 'manual') {
    if (filter.family !== 'manual') throw new Error(`Pivot manual filter family is invalid: ${fieldId}`);
    return { kind: 'manual', family: 'manual', fieldId, scope: filter.scope ?? 'report', mode: filter.mode, memberKeys: structuredClone(filter.memberKeys) };
  }
  if (filter.kind === 'top-items') {
    if (filter.family !== 'top-items') throw new Error(`Pivot top-items filter family is invalid: ${fieldId}`);
    const valueFieldId = resolveFieldId(filter.valueFieldId, catalog);
    if (!valueFieldId) throw new Error(`Unknown pivot value field: ${filter.valueFieldId}`);
    return { ...filter, fieldId, valueFieldId };
  }
  if (!['label', 'date', 'value'].includes(filter.family)) throw new Error(`Pivot condition filter family is invalid: ${fieldId}`);
  if (filter.valueFieldId !== undefined) {
    const valueFieldId = resolveFieldId(filter.valueFieldId, catalog);
    if (!valueFieldId) throw new Error(`Unknown pivot value filter field: ${filter.valueFieldId}`);
    return { ...filter, fieldId, valueFieldId };
  }
  return { ...filter, fieldId };
}

function normalizeValueField(field: PivotValueField, catalog: PivotFieldCatalog): PivotValueField {
  const fieldId = resolveFieldId(field.fieldId, catalog);
  if (!fieldId) throw new Error(`Unknown pivot value field: ${field.fieldId}`);
  const numberFormat = normalizePivotNumberFormat(field.numberFormat);
  return { ...field, fieldId, ...(field.baseFieldId ? { baseFieldId: resolveFieldId(field.baseFieldId, catalog) } : {}), ...(numberFormat === undefined ? {} : { numberFormat }) };
}

function normalizeLayout(layout: PivotLayout, catalog: PivotFieldCatalog): PivotLayout {
  if (!['compact', 'outline', 'tabular'].includes(layout.reportLayout)) throw new Error('Pivot report layout is invalid');
  const values = layout.values.map((entry) => normalizeValueField(entry, catalog));
  const valueFieldIds = new Set(values.map((entry) => entry.fieldId));
  const filters = layout.filters.map((entry) => normalizeFilter(entry, catalog));
  const normalizedRows = layout.rows.map((entry) => normalizePlacement(entry, catalog, valueFieldIds));
  const normalizedColumns = layout.columns.map((entry) => normalizePlacement(entry, catalog, valueFieldIds));
  for (const filter of filters) {
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
  for (const filter of filters) {
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
    filters,
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

/** Canonicalize field catalog values against the live source. Calculation has one model shape. */
export function normalizePivotDefinition(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotDefinition {
  const source = getPivotSource(pivot);
  const calculator = source.kind === 'data-source' ? formula : (formula ?? createPivotFormulaEngine(workbook));
  const fieldCatalog = source.kind === 'data-source'
    ? getPivotFieldCatalog(workbook, pivot)
    : normalizeFieldCatalog(sourceTable(workbook, pivot, pivot.fieldCatalog, calculator), pivot.fieldCatalog);
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
    refreshPolicy: normalizePivotRefreshPolicy(pivot.refreshPolicy),
    presentation: {
      ...(pivot.presentation?.styleName ? { styleName: pivot.presentation.styleName } : {}),
      styleOptions: { ...DEFAULT_PIVOT_STYLE_OPTIONS, ...(pivot.presentation?.styleOptions ?? {}) },
    },
    ...(pivot.nativeMetadata ? { nativeMetadata: structuredClone(pivot.nativeMetadata) } : {}),
  };
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
  const calculator = formula ?? createPivotFormulaEngine(workbook);
  return normalizeFieldCatalog(sourceTable(workbook, { ...pivot, source }, pivot.fieldCatalog, calculator), pivot.fieldCatalog);
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
  const sourceError = Object.values(row.values).find(isPivotError);
  if (sourceError) return sourceError;
  const engine = new FormulaEngine({ defaultSheetId: 'pivot' });
  fields.forEach((field, index) => {
    const value = row.values[field.fieldId] ?? null;
    engine.setValue({ sheetId: 'pivot', row: 0, column: index }, isPivotError(value) ? null : value);
  });
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
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (leftNumber != null && rightNumber != null) return leftNumber - rightNumber;
  return collator.compare(String(left), String(right));
}

/** Every aggregate has its own semantics; no operation falls through to sum. */
export function aggregatePivotValues(rows: ReadonlyArray<{ values: Record<string, PivotScalar> }>, fieldId: string, operation: PivotAggregateFunction): PivotScalar {
  const numbers: number[] = [];
  const members = new Set<string>();
  const errors: Extract<PivotScalar, { kind: 'error' }>[] = [];
  let nonBlank = 0;
  for (const row of rows) {
    const raw = row.values[fieldId] ?? null;
    if (raw != null && raw !== '') nonBlank += 1;
    if (isPivotError(raw)) errors.push(raw);
    if (raw != null && raw !== '') members.add(pivotMemberKey(createPivotMemberKey(raw)));
    const number = toNumber(raw);
    if (number != null) numbers.push(number);
  }
  const firstError = errors[0];
  switch (operation) {
    case 'count': return nonBlank;
    case 'count-numbers': return numbers.length;
    case 'sum': return firstError ?? numbers.reduce((sum, value) => sum + value, 0);
    case 'average': return firstError ?? (numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : null);
    case 'min': return firstError ?? (numbers.length ? Math.min(...numbers) : null);
    case 'max': return firstError ?? (numbers.length ? Math.max(...numbers) : null);
    case 'product': return firstError ?? (numbers.length ? numbers.reduce((product, value) => product * value, 1) : null);
    case 'distinct-count': return members.size;
    case 'stdev': {
      if (firstError) return firstError;
      if (numbers.length < 2) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return Math.sqrt(numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - 1));
    }
    case 'stdevp': {
      if (firstError) return firstError;
      if (!numbers.length) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return Math.sqrt(numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / numbers.length);
    }
    case 'var': {
      if (firstError) return firstError;
      if (numbers.length < 2) return null;
      const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
      return numbers.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (numbers.length - 1);
    }
    case 'varp': {
      if (firstError) return firstError;
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

function axisGroups(rows: SourceRow[], placements: PivotFieldPlacement[], fieldCatalog: PivotFieldCatalog, collator: Intl.Collator): AxisGroup[] {
  const map = new Map<string, AxisGroup>();
  for (const row of rows) {
    const values = placements.map((placement) => grouped(row.values[placement.fieldId] ?? null, placement.group));
    const key = JSON.stringify(values.map(createPivotMemberKey));
    const group = map.get(key) ?? { values, rows: [] };
    group.rows.push(row);
    map.set(key, group);
  }
  const placement = placements[placements.length - 1];
  const dataType = placement ? fieldCatalog.fields.find((field) => field.fieldId === placement.fieldId)?.dataType : undefined;
  const result = [...map.values()].sort((left, right) => {
    if (placement?.sort?.by === 'value' && placement.sort.valueFieldId) {
      return (toNumber(aggregatePivotValues(left.rows, placement.sort.valueFieldId, 'sum')) ?? 0) - (toNumber(aggregatePivotValues(right.rows, placement.sort.valueFieldId, 'sum')) ?? 0);
    }
    for (let index = 0; index < left.values.length; index += 1) {
      const fieldType = fieldCatalog.fields.find((field) => field.fieldId === placements[index]?.fieldId)?.dataType ?? dataType;
      const order = compare(left.values[index] ?? null, right.values[index] ?? null, fieldType, collator);
      if (order) return order;
    }
    return 0;
  });
  if (placement?.sort?.direction === 'descending') result.reverse();
  return result;
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
  if (filter.kind === 'condition' && filter.valueFieldId !== undefined) return undefined;
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

function matchesFilter(row: SourceRow, filter: PivotFilter, collator: Intl.Collator, definition?: PivotDefinition): boolean {
  const fieldId = filter.fieldId;
  const rawValue = filter.kind === 'condition' && filter.valueFieldId
    ? row.values[filter.valueFieldId] ?? null
    : row.values[fieldId] ?? null;
  const placement = definition ? groupedPlacementForFilter(definition, filter) : undefined;
  const value = placement?.group ? grouped(rawValue, placement.group) : rawValue;
  if (filter.kind === 'top-items') return true;
  if (filter.kind === 'manual') return manualFilterMatches(rawValue, filter, placement?.group);
  if (filter.family === 'date') return placement?.group ? groupedDateFilterMatches(rawValue, filter, placement.group, collator) : dateFilterMatches(value, filter);
  if (filter.family === 'label') return labelFilterMatches(value, filter, collator);
  const leftNumber = toNumber(value);
  const rightNumber = toNumber(filter.value);
  const order = leftNumber != null && rightNumber != null ? leftNumber - rightNumber : compare(value, filter.value, undefined, collator);
  switch (filter.operator) {
    case 'equals': return same(value, filter.value);
    case 'not-equals': return !same(value, filter.value);
    case 'greater-than': return order > 0;
    case 'greater-or-equal': return order >= 0;
    case 'less-than': return order < 0;
    case 'less-or-equal': return order <= 0;
    case 'between': return filter.value2 !== undefined && order >= 0 && (toNumber(value) != null && toNumber(filter.value2) != null ? toNumber(value)! <= toNumber(filter.value2)! : compare(value, filter.value2, undefined, collator) <= 0);
    case 'not-between': return filter.value2 !== undefined && !(order >= 0 && (toNumber(value) != null && toNumber(filter.value2) != null ? toNumber(value)! <= toNumber(filter.value2)! : compare(value, filter.value2, undefined, collator) <= 0));
    default: return false;
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
    const ranked = [...buckets.values()].sort((left, right) => (toNumber(aggregatePivotValues(left, valueFieldId, 'sum')) ?? 0) - (toNumber(aggregatePivotValues(right, valueFieldId, 'sum')) ?? 0));
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

interface LinkedPivotControl {
  drawingId: string;
  payload: PivotSlicerDrawingPayload | PivotTimelineDrawingPayload;
}

function linkedPivotControls(workbook: WorkbookModel, pivot: PivotModel): LinkedPivotControl[] {
  return workbook.getSheets().flatMap((sheet) => sheet.drawings.flatMap((drawing) => {
    if (drawing.kind !== 'slicer' && drawing.kind !== 'timeline') return [];
    const payload = sheet.drawingPayloads.get(drawing.payloadId);
    if (!payload || (payload.kind !== 'slicer' && payload.kind !== 'timeline')) return [];
    const linked = [payload.pivotId, ...(payload.connectedPivotIds ?? [])];
    return linked.includes(pivot.id) ? [{ drawingId: drawing.id, payload }] : [];
  }));
}

function matchesTimeline(row: SourceRow, timeline: PivotTimelineDrawingPayload, bounds: PivotTimelinePeriodBounds): boolean {
  const fieldId = timeline.fieldId;
  const raw = row.values[fieldId];
  if (raw == null || raw === '') return false;
  const instant = pivotTimelineInstant(raw);
  if (instant === undefined) return false;
  return instant >= (bounds.start ?? Number.NEGATIVE_INFINITY)
    && instant < (bounds.endExclusive ?? Number.POSITIVE_INFINITY);
}

function matchesControls(workbook: WorkbookModel, rows: SourceRow[], pivot: PivotModel, excludedSlicerDrawingId?: string): SourceRow[] {
  const controls = linkedPivotControls(workbook, pivot).filter((entry) => entry.drawingId !== excludedSlicerDrawingId);
  const slicers = controls.filter((entry): entry is LinkedPivotControl & { payload: PivotSlicerDrawingPayload } => entry.payload.kind === 'slicer');
  const timelines = controls.filter((entry): entry is LinkedPivotControl & { payload: PivotTimelineDrawingPayload } => entry.payload.kind === 'timeline');
  const timelineBounds = timelines.map((entry) => normalizePivotTimelinePeriod(entry.payload.period));
  return rows.filter((row) => slicers.every((entry) => matchesSlicer(row, entry.payload))
    && timelines.every((entry, index) => matchesTimeline(row, entry.payload, timelineBounds[index]!)));
}

function slicerItemProjection(
  workbook: WorkbookModel,
  pivot: PivotModel,
  definition: PivotDefinition,
  rows: SourceRow[],
  drawingId: string,
  payload: PivotSlicerDrawingPayload,
  collator: Intl.Collator,
): PivotSlicerItemProjection[] {
  const fieldValues = rows.map((row) => row.values[payload.fieldId] ?? null);
  const members = new Map<string, PivotSlicerItemProjection>();
  for (const value of fieldValues) {
    const key = createPivotMemberKey(value);
    const identity = pivotMemberKey(key);
    if (!members.has(identity)) members.set(identity, { key, value, label: formatPivotMember(value), selected: false, hasData: false });
  }
  const filteredRows = matchesControls(workbook, rows, pivot, drawingId)
    .filter((row) => definition.layout.filters.filter((filter) => filter.kind !== 'top-items').every((filter) => matchesFilter(row, filter, collator, definition)));
  const availableRows = topItems(filteredRows, definition.layout.filters);
  const available = new Set(availableRows.map((row) => pivotMemberKey(createPivotMemberKey(row.values[payload.fieldId] ?? null))));
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
      sourceFieldId: field.fieldId,
      subtotalFunction: fn,
      subtotalFieldId: fieldId,
      displayName: `${field.displayName ?? field.fieldId} (${fn})`,
    }));
    return [base, ...extras];
  });
}

function resultCells(rows: SourceRow[], columns: AxisGroup[], values: PivotResultValueField[], nodePath: string[], kind: PivotResultCell['kind'] = 'detail', subtotalFieldId?: string): PivotResultCell[] {
  return columns.map((column, columnIndex) => {
    const nodeRows = new Set(rows);
    const columnRows = column.rows.filter((candidate) => nodeRows.has(candidate));
    return {
      id: `${nodePath.join('/') || 'root'}|column:${columnIndex}`,
      nodePath,
      kind,
      columnPath: column.values,
      sourceRowPaths: columnRows.flatMap((row) => row.paths),
      values: values.map((value) => aggregatePivotValues(columnRows, value.sourceFieldId, kind === 'subtotal' && value.subtotalFieldId === subtotalFieldId
        ? value.subtotalFunction ?? value.summarizeBy
        : value.summarizeBy)),
    };
  });
}

function resultGrandTotalCell(rows: SourceRow[], values: PivotResultValueField[], nodePath: string[], subtotalFieldId?: string): PivotResultCell {
  return {
    id: `${nodePath.join('/') || 'root'}|grand-total:row`,
    nodePath,
    kind: 'grand-total',
    columnPath: [],
    sourceRowPaths: rows.flatMap((row) => row.paths),
    values: values.map((value) => aggregatePivotValues(rows, value.sourceFieldId, subtotalFieldId === value.subtotalFieldId
      ? value.subtotalFunction ?? value.summarizeBy
      : value.summarizeBy)),
  };
}

function resultNodes(rows: SourceRow[], placements: PivotFieldPlacement[], depth: number, columns: AxisGroup[], values: PivotResultValueField[], subtotalLocation: PivotLayout['subtotalLocation'], showRowGrandTotals: boolean, fieldCatalog: PivotFieldCatalog, collator: Intl.Collator, prefix: string[] = []): PivotResultNode[] {
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
      values: resultCells(rows, columns, values, path),
      ...(showRowGrandTotals ? { rowGrandTotal: resultGrandTotalCell(rows, values, path) } : {}),
      subtotal: false,
      sourceRowPaths: rows.flatMap((row) => row.paths),
    }];
  }
  const placement = placements[depth]!;
  return axisGroups(rows, [placement], fieldCatalog, collator).map((group) => {
    const fieldId = placement.fieldId;
    const member = createPivotMemberKey(group.values[0] ?? null);
    const path = [...prefix, `${fieldId}=${pivotMemberKey(member)}`];
    const children = resultNodes(group.rows, placements, depth + 1, columns, values, subtotalLocation, showRowGrandTotals, fieldCatalog, collator, path);
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
      values: resultCells(group.rows, columns, values, path, subtotal ? 'subtotal' : 'detail', subtotal ? placement.fieldId : undefined),
      ...(showRowGrandTotals ? { rowGrandTotal: resultGrandTotalCell(group.rows, values, path, subtotal ? placement.fieldId : undefined) } : {}),
      subtotal,
      sourceRowPaths: group.rows.flatMap((row) => row.paths),
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

/**
 * Apply Show Values As from one immutable result matrix.
 *
 * The result tree contains three different calculation domains: detail rows,
 * subtotal rows and the grand-total row.  Keeping those contexts explicit is
 * important because subtotal values are valid Pivot members in their own
 * right; they must never be looked up in a leaf-only sequence.
 */
function applyShowAs(tree: PivotResultTree, fields: PivotValueField[]): void {
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
  const subtotalContexts = rowContexts.filter((context) => context.kind === 'subtotal');

  const parentKey = (context: PivotShowAsCellContext): string => context.parent?.path?.join('\u001f') ?? '';
  const rowSeries = (context: PivotShowAsCellContext): PivotShowAsCellContext[] => {
    if (context.kind === 'subtotal' && context.node) {
      // Subtotals are ranked/accumulated against subtotal peers of the same
      // hierarchy field and parent context, never against leaf aggregates.
      return subtotalContexts.filter((candidate) => candidate.node?.depth === context.node?.depth && parentKey(candidate) === parentKey(context));
    }
    return leafContexts;
  };

  const numericSum = (cells: PivotShowAsCellContext[], valueIndex: number, columnIndex: number): number => cells.reduce((sum, context) => {
    const cell = context.node?.values[columnIndex];
    return sum + (toNumber(rawValue(cell, valueIndex)) ?? 0);
  }, 0);

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
      if (spec.kind === 'difference') return spec.base === 'grand' ? current - (grand ?? current) : 0;
      if (spec.kind === 'percentage-difference') {
        const base = spec.base === 'grand' ? grand : rowTotal || columnTotal || parentTotal;
        return base ? (current - base) / base : null;
      }
      if (spec.kind === 'running-total') return current;
      if (spec.kind === 'rank') return 1;
      if (spec.kind === 'index') return grand != null && rowTotal && columnTotal ? current * grand / rowTotal / columnTotal : null;
    }
    if (spec.kind === 'grand-percentage') return grand ? current / grand : null;
    if (spec.kind === 'row-percentage') return rowTotal ? current / rowTotal : null;
    if (spec.kind === 'column-percentage') return columnTotal ? current / columnTotal : null;
    if (spec.kind === 'parent-percentage') return parentTotal == null ? null : parentTotal ? current / parentTotal : null;
    if (spec.kind === 'difference' || spec.kind === 'percentage-difference') {
      const base = spec.base === 'grand' ? grand : spec.base === 'row' ? rowTotal : spec.base === 'column' ? columnTotal : parentTotal;
      return base == null ? null : spec.kind === 'difference' ? current - base : base ? (current - base) / base : null;
    }
    if (spec.kind === 'running-total') {
      if (spec.axis === 'column') {
        const values = context.node?.values.slice(0, context.columnIndex + 1) ?? [];
        return values.reduce((sum, cell) => sum + (toNumber(rawValue(cell, valueIndex)) ?? 0), 0);
      }
      const series = rowSeries(context);
      const end = series.indexOf(context);
      return end < 0 ? null : series.slice(0, end + 1).reduce((sum, candidate) => sum + (toNumber(rawValue(candidate.node?.values[candidate.columnIndex], valueIndex)) ?? 0), 0);
    }
    if (spec.kind === 'rank') {
      const series = spec.axis === 'column'
        ? (context.node?.values ?? []).map((cell) => toNumber(rawValue(cell, valueIndex)))
        : rowSeries(context).map((candidate) => toNumber(rawValue(candidate.node?.values[candidate.columnIndex], valueIndex)));
      const ranked = series.filter((value): value is number => value != null).sort((left, right) => spec.direction === 'ascending' ? left - right : right - left);
      const rank = ranked.findIndex((value) => value === current);
      return rank < 0 ? null : rank + 1;
    }
    if (spec.kind === 'index') return grand != null && rowTotal && columnTotal ? current * grand / rowTotal / columnTotal : null;
    return null;
  };

  for (const context of contexts) {
    for (const [valueIndex, field] of fields.entries()) {
      const spec = field.showAs ?? { kind: 'normal' as const };
      const current = toNumber(rawValue(context.cell, valueIndex));
      if (current == null || spec.kind === 'normal') continue;
      const grand = toNumber(grandValues[valueIndex] ?? null);
      const rowTotal = context.kind === 'grand-total'
        ? (grand ?? current)
        : context.node?.values.reduce((sum, cell) => sum + (toNumber(rawValue(cell, valueIndex)) ?? 0), 0) ?? 0;
      const columnTotal = context.kind === 'grand-total'
        ? (grand ?? current)
        : numericSum(leafContexts, valueIndex, context.columnIndex);
      // Top-level members have the grand total as their parent context. This
      // is the only deterministic parent for a Pivot root member.
      const parentTotal = context.kind === 'grand-total'
        ? grand
        : context.parent ? toNumber(rawValue(context.parent.values[context.columnIndex], valueIndex)) : grand;
      context.cell.values[valueIndex] = transform(spec, current, grand, rowTotal, columnTotal, parentTotal, context, valueIndex);
    }
  }
}

function computePivotResultFromTable(
  workbook: WorkbookModel,
  pivot: PivotModel,
  definition: PivotDefinition,
  rawTable: PivotSourceTableInput,
  sourceRevisionOverride?: string,
  formula?: FormulaEngine,
): PivotResultTree {
  const collator = createPivotCollator(definition.layout.collation);
  const rows = applyCalculatedData(rawTable.rows, definition.fieldCatalog.fields, definition.layout.calculatedFields, definition.layout.calculatedItems);
  const references = [
    ...definition.layout.rows.map((entry) => entry.fieldId),
    ...definition.layout.columns.map((entry) => entry.fieldId),
    ...definition.layout.filters.flatMap((filter) => filter.kind === 'top-items' || (filter.kind === 'condition' && filter.valueFieldId) ? [filter.fieldId, filter.valueFieldId] : [filter.fieldId]),
    ...definition.layout.values.map((entry) => entry.fieldId),
  ];
  const known = new Set([...definition.fieldCatalog.fields.map((field) => field.fieldId), ...(definition.layout.calculatedFields ?? []).map((field) => field.fieldId), ...(definition.layout.calculatedItems ?? []).map((field) => field.fieldId)]);
  const unknown = references.find((field) => field && !known.has(field));
  if (unknown && rawTable.fields.length) throw new Error(`Unknown pivot field: ${unknown}`);
  let filtered = matchesControls(workbook, rows, pivot);
  filtered = filtered.filter((row) => definition.layout.filters.filter((filter) => filter.kind !== 'top-items').every((filter) => matchesFilter(row, filter, collator, definition)));
  filtered = topItems(filtered, definition.layout.filters);
  const columns = definition.layout.columns.length ? axisGroups(filtered, definition.layout.columns, definition.fieldCatalog, collator) : [{ values: [], rows: filtered }];
  const resultFields = resultValueFields(definition.layout);
  const grandTotal: PivotResultCell = {
    id: `${definition.id}|grand-total`,
    kind: 'grand-total',
    columnPath: [],
    values: resultFields.map((field) => aggregatePivotValues(filtered, field.sourceFieldId, field.summarizeBy)),
    sourceRowPaths: filtered.flatMap((row) => row.paths),
  };
  const tree: PivotResultTree = {
    schema: PIVOT_RESULT_TREE_SCHEMA,
    pivotId: definition.id,
    fields: definition.fieldCatalog,
    columnPaths: columns.map((column) => column.values),
    valueFields: resultFields,
    rows: resultNodes(filtered, definition.layout.rows, 0, columns, resultFields, definition.layout.subtotalLocation, definition.layout.showRowGrandTotals, definition.fieldCatalog, collator),
    columnGrandTotals: resultCells(filtered, columns, resultFields, [`${definition.id}|grand-total`], 'grand-total'),
    grandTotal,
    sourceRowPaths: filtered.flatMap((row) => row.paths),
  };
  const controls = linkedPivotControls(workbook, pivot);
  const slicerItems: Record<string, PivotSlicerItemProjection[]> = {};
  for (const control of controls) {
    if (control.payload.kind !== 'slicer') continue;
    slicerItems[control.drawingId] = slicerItemProjection(workbook, pivot, definition, rows, control.drawingId, control.payload, collator);
  }
  if (Object.keys(slicerItems).length > 0) tree.slicerItems = slicerItems;
  applyShowAs(tree, resultFields);
  const revisions = getPivotRevisionKey(workbook, pivot, formula);
  tree.sourceRevision = sourceRevisionOverride ?? revisions.sourceRevision;
  tree.layoutRevision = revisions.layoutRevision;
  tree.filterRevision = revisions.filterRevision;
  return tree;
}

function computePivotResultUncached(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotResultTree {
  const calculator = pivot.source.kind === 'data-source' ? formula : (formula ?? createPivotFormulaEngine(workbook));
  const definition = normalizePivotDefinition(workbook, pivot, calculator);
  const rawTable = sourceTable(workbook, pivot, definition.fieldCatalog, calculator);
  return computePivotResultFromTable(workbook, pivot, definition, rawTable, undefined, calculator);
}

export function computePivotResult(workbook: WorkbookModel, pivot: PivotModel, formula?: FormulaEngine): PivotResultTree {
  return structuredClone(computePivotResultUncached(workbook, pivot, formula));
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
  const definition = normalizePivotDefinition(workbook, pivot, options.formula);
  const displayOptions = normalizePivotDisplayOptions(definition.presentation?.displayOptions);
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
      tree = computePivotResult(workbook, pivot, options.formula);
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
    const filterSummary = summarizePivotReportFilters(definition.layout.filters, definition.fieldCatalog, fieldId);
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
      if (displayOptions.showFieldHeaders) cells.push(projectionCell(definition.id, row, column, 'column-header', path[0] ?? null, label, { columnPath: path, fieldId: definition.layout.columns[definition.layout.columns.length - 1]?.fieldId, isLastColumn: !definition.layout.showRowGrandTotals && columnIndex === columnPathCount - 1 }));
    }
  }
  if (definition.layout.showRowGrandTotals && displayOptions.showFieldHeaders) {
    for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
      const valueField = values[valueIndex];
      const column = rowHeaderCount + columnPathCount * Math.max(values.length, 1) + valueIndex;
      cells.push(projectionCell(definition.id, row, column, 'column-header', null, valueField ? `Grand Total ${valueField.displayName ?? fieldName(valueField.fieldId, definition.fieldCatalog)}` : 'Grand Total', { captionKey: 'grand-total', isLastColumn: valueIndex === Math.max(values.length, 1) - 1 }));
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
          cells.push(projectionCell(definition.id, row, column, node.subtotal ? 'subtotal' : 'value', value, textForValue(value, displayOptions, valueField?.numberFormat), { nodeId: node.nodeId, resultCellId: resultCell?.id, columnPath: resultCell?.columnPath, sourceRowPaths: resultCell?.sourceRowPaths, isLastColumn: !definition.layout.showRowGrandTotals && columnIndex === columnPathCount - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
        }
      }
      if (definition.layout.showRowGrandTotals) {
        const resultCell = node.rowGrandTotal;
        for (let valueIndex = 0; valueIndex < Math.max(values.length, 1); valueIndex += 1) {
          const column = rowHeaderCount + columnPathCount * Math.max(values.length, 1) + valueIndex;
          const value = resultCell?.values[valueIndex] ?? null;
          const valueField = values[valueIndex];
          cells.push(projectionCell(definition.id, row, column, 'grand-total', value, textForValue(value, displayOptions, valueField?.numberFormat), { nodeId: node.nodeId, resultCellId: resultCell?.id, columnPath: resultCell?.columnPath, sourceRowPaths: resultCell?.sourceRowPaths, isLastColumn: valueIndex === Math.max(values.length, 1) - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
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
        cells.push(projectionCell(definition.id, row, column, 'grand-total', value, textForValue(value, displayOptions, valueField?.numberFormat), { resultCellId: resultCell.id, columnPath: resultCell.columnPath, sourceRowPaths: resultCell.sourceRowPaths, isLastColumn: !definition.layout.showRowGrandTotals && columnIndex === columnGrandTotals.length - 1 && valueIndex === resultCell.values.length - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
      }));
      if (definition.layout.showRowGrandTotals) {
        tree.grandTotal.values.forEach((value, valueIndex) => {
          const column = rowHeaderCount + columnPathCount * Math.max(values.length, 1) + valueIndex;
          const valueField = values[valueIndex];
          cells.push(projectionCell(definition.id, row, column, 'grand-total', value, textForValue(value, displayOptions, valueField?.numberFormat), { resultCellId: tree.grandTotal?.id, sourceRowPaths: tree.grandTotal?.sourceRowPaths, isLastColumn: valueIndex === tree.grandTotal!.values.length - 1, ...(valueField?.numberFormat ? { numberFormat: valueField.numberFormat } : {}) }));
        });
      }
      row += 1;
    }
  } else {
    cells.push(projectionCell(definition.id, row, 0, error ? 'error' : 'loading', null, error ?? 'Loading PivotTable', error ? {} : { captionKey: 'loading' }));
    row += 1;
  }
  const occupiedRange = projectionRange(target, Math.max(row, 1), Math.max(valueColumnCount + rowHeaderCount, 1));
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
  const blockResultReady = pivot.source.kind === 'data-source'
    && options.sourceState?.availability === 'ready'
    && pivotResultMatchesLayoutAndFilter(workbook, pivot, cachedResult, options.formula);
  const staleManualResult = pivot.refreshPolicy.mode === 'manual'
    && pivotResultMatchesLayoutAndFilter(workbook, pivot, cachedResult, options.formula)
    && cachedResult.sourceRevision !== revision.sourceRevision;
  let effectiveResult = pivotResultMatchesRevision(workbook, pivot, cachedResult, options.formula) || staleManualResult || blockResultReady ? cachedResult : undefined;
  if (!effectiveResult && pivot.source.kind !== 'data-source') {
    try {
      effectiveResult = computePivotResult(workbook, pivot, options.formula);
    } catch {
      // The candidate builder creates the explicit synchronous error state.
    }
  }
  const candidate = buildPivotGridProjectionCandidate(workbook, pivot, effectiveResult, options);
  const cache = lastValidPivotProjections.get(workbook);
  const last = cache?.get(pivot.id);
  const candidateTree = effectiveResult;

  if (staleManualResult && candidate.collision.status === 'clear') {
    candidate.refresh = refreshState(workbook, pivot, candidate.collision, 'stale', undefined, options.formula);
    return candidate;
  }

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
