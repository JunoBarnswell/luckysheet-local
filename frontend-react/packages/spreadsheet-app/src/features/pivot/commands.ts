import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import type {
  PivotAggregateFunction,
  PivotDefinition,
  PivotSource,
  PivotGroup,
  PivotLayout,
  PivotModel,
  PivotPresentation,
  PivotShowAs,
  PivotSourceRowPath,
  ChartDrawingPayload,
  DrawingObject,
  RangeRef,
} from '@react-sheets/core-model';
import { assertPivotDefinition, assertPivotField, createPivotDrillDownSheetName, setPivotAggregate, setPivotGroup, setPivotShowAs } from './panel-state';
import { buildPivotGridProjection, computePivotResult, getPivotSourceRanges, normalizePivotDefinition } from './engine';

function sheetRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

function removeById<T extends { id: string }>(items: T[], id: string): T | undefined {
  const index = items.findIndex((item) => item.id === id);
  if (index < 0) return undefined;
  return items.splice(index, 1)[0];
}

export interface PivotUpdateParams {
  sheetId: string;
  pivotId: string;
  source?: PivotSource;
  target?: PivotDefinition['target'];
  fieldCatalog?: PivotDefinition['fieldCatalog'];
  refreshPolicy?: PivotDefinition['refreshPolicy'];
  nativeMetadata?: PivotDefinition['nativeMetadata'];
  presentation?: PivotPresentation;
  layout?: PivotLayout;
}

export type PivotCreateDestination =
  | {
    kind: 'new-sheet';
    sheetId: string;
    name: string;
    rowCount?: number;
    columnCount?: number;
  }
  | {
    kind: 'existing-sheet';
    sheetId: string;
  };

/**
 * The only command contract for creating a PivotTable.  A new destination
 * worksheet is part of this command rather than a preceding sheet command so
 * the runtime can record, publish, and roll back the complete operation as one
 * root transaction.
 */
export interface PivotCreateParams {
  pivot: PivotModel;
  destination: PivotCreateDestination;
}

export interface PivotLayoutCommandParams {
  sheetId: string;
  pivotId: string;
  layout: PivotLayout;
}

export interface PivotExpansionToggleParams {
  sheetId: string;
  pivotId: string;
  nodeId: string;
}

export interface PivotExpansionFieldParams {
  sheetId: string;
  pivotId: string;
  fieldId: string;
  expanded: boolean;
}

export interface PivotExpansionButtonsParams {
  sheetId: string;
  pivotId: string;
  showButtons: boolean;
}

export interface PivotAggregateParams {
  sheetId: string;
  pivotId: string;
  fieldId: string;
  summarizeBy: PivotAggregateFunction;
}

export interface PivotShowAsParams {
  sheetId: string;
  pivotId: string;
  fieldId: string;
  showAs: PivotShowAs;
}

export interface PivotGroupParams {
  sheetId: string;
  pivotId: string;
  axis: 'rows' | 'columns';
  fieldId: string;
  group: PivotGroup;
}

export interface PivotDrillDownParams {
  sheetId: string;
  pivotId: string;
  label: string;
  sourceRowPaths: PivotSourceRowPath[];
  targetSheetId: string;
  target: { row: number; column: number };
}

export interface PivotDrillDownRemoveParams {
  targetSheetId: string;
}


/** A PivotChart is one command and one reversible drawing mutation. */
export interface PivotChartCreateParams {
  sheetId: string;
  pivotId: string;
  drawing: DrawingObject;
  payload: ChartDrawingPayload;
}

function pivotFor(context: CommandContext, sheetId: string, pivotId: string): PivotModel | undefined {
  const direct = context.workbook.getSheet(sheetId).pivots.find((entry) => entry.id === pivotId);
  if (direct) return direct;
  return context.workbook.getSheets().flatMap((sheet) => sheet.pivots).find((entry) => entry.id === pivotId);
}

function pivotDisplaySheetId(pivot: PivotModel, fallback = ''): string {
  return pivot.target.sheetId || fallback;
}

function pivotSourceRanges(pivot: PivotModel, workbook?: import('@react-sheets/core-model').WorkbookModel): RangeRef[] {
  if (workbook) {
    try { return getPivotSourceRanges(workbook, pivot); } catch { /* validator/migration will report the source error */ }
  }
  if (pivot.source.kind === 'worksheet-ranges') return pivot.source.ranges.map((sourceRange) => structuredClone(sourceRange.range));
  if (pivot.source.kind === 'worksheet-range') return [structuredClone(pivot.source.range)];
  return [];
}

function pivotSourceNodes(pivot: PivotModel, workbook?: import('@react-sheets/core-model').WorkbookModel): Array<{ sourceId?: string; range: RangeRef }> {
  if (pivot.source.kind === 'worksheet-ranges') return pivot.source.ranges.map((sourceRange) => ({ sourceId: sourceRange.sourceId, range: structuredClone(sourceRange.range) }));
  return pivotSourceRanges(pivot, workbook).map((range) => ({ range }));
}

interface DrillDownColumn {
  sourceId?: string;
  range: RangeRef;
  column: number;
  label: string;
}

function sourceCellValue(cell: { value: unknown; formulaValue?: unknown } | undefined): string | number | boolean | null {
  const value = cell?.formulaValue ?? cell?.value ?? null;
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || value === null ? value : null;
}

interface PivotDrillDownRecord {
  paths: Map<string, PivotSourceRowPath>;
}

interface PivotDrillDownPlan {
  columns: DrillDownColumn[];
  records: PivotDrillDownRecord[];
}

function drillDownRootSourceId(pivot: PivotModel, nodes: Array<{ sourceId?: string; range: RangeRef }>): string | undefined {
  if (pivot.source.kind !== 'worksheet-ranges') return nodes[0]?.sourceId;
  const leftJoins = pivot.source.relationships.filter((relationship) => relationship.join === 'left');
  const candidates = leftJoins.length
    ? nodes.filter((node) => node.sourceId && !leftJoins.some((relationship) => relationship.right.sourceId === node.sourceId))
    : [...nodes].sort((left, right) => (left.sourceId ?? '').localeCompare(right.sourceId ?? '')).slice(0, 1);
  if (candidates.length !== 1 || !candidates[0]?.sourceId) throw new Error('Pivot drill-down source graph has no deterministic root');
  return candidates[0].sourceId;
}

function optionalDrillDownSourceIds(pivot: PivotModel): Set<string> {
  if (pivot.source.kind !== 'worksheet-ranges') return new Set();
  const optional = new Set<string>();
  for (const relationship of pivot.source.relationships) {
    if (relationship.join === 'left') optional.add(relationship.right.sourceId);
  }
  let changed = true;
  while (changed) {
    changed = false;
    for (const relationship of pivot.source.relationships) {
      if (relationship.join === 'left' && optional.has(relationship.left.sourceId)) {
        if (!optional.has(relationship.right.sourceId)) {
          optional.add(relationship.right.sourceId);
          changed = true;
        }
      }
    }
  }
  return optional;
}

function planPivotDrillDown(context: CommandContext, params: PivotDrillDownParams): PivotDrillDownPlan {
  const sourceSheet = context.workbook.getSheet(params.sheetId);
  const pivot = sourceSheet.pivots.find((entry) => entry.id === params.pivotId);
  if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
  if (context.workbook.sheets.has(params.targetSheetId)) throw new Error(`Drill-down target already exists: ${params.targetSheetId}`);
  const nodes = pivotSourceNodes(pivot, context.workbook);
  const sourceIds = new Set<string>();
  for (const node of nodes) {
    const sheet = context.workbook.getSheet(node.range.sheetId);
    if (node.range.startRow < 0 || node.range.endRow >= sheet.rowCount || node.range.startColumn < 0 || node.range.endColumn >= sheet.columnCount) {
      throw new Error('Pivot source range exceeds worksheet bounds');
    }
    if (node.sourceId !== undefined && (!isNonEmptyString(node.sourceId) || sourceIds.has(node.sourceId))) {
      throw new Error('Pivot drill-down source identity is invalid or duplicated');
    }
    if (node.sourceId !== undefined) sourceIds.add(node.sourceId);
  }
  const multiSource = nodes.length > 1;
  const rootSourceId = drillDownRootSourceId(pivot, nodes);
  const optionalSourceIds = optionalDrillDownSourceIds(pivot);
  const requiredSourceIds = multiSource
    ? [...sourceIds].filter((sourceId) => !optionalSourceIds.has(sourceId))
    : [];
  const records = new Map<string, PivotDrillDownRecord>();
  for (const path of params.sourceRowPaths) {
    const sheet = context.workbook.getSheet(path.sheetId);
    if (path.row < 0 || path.row >= sheet.rowCount) throw new Error(`Pivot drill-down source row is invalid: ${path.row}`);
    const node = multiSource
      ? nodes.find((candidate) => candidate.sourceId === path.sourceId)
      : nodes.length === 1
        ? nodes[0]
        : nodes.find((candidate) => candidate.sourceId === path.sourceId || (candidate.sourceId === undefined && candidate.range.sheetId === path.sheetId));
    if (!node) throw new Error(`Pivot drill-down provenance references an unknown source: ${path.sourceId ?? path.sheetId}`);
    if (path.sheetId !== node.range.sheetId || path.row <= node.range.startRow || path.row > node.range.endRow) {
      throw new Error(`Pivot drill-down provenance is outside source range: ${path.sourceId ?? path.sheetId}:${path.row}`);
    }
    if (multiSource && (!path.sourceId || !path.recordId)) {
      throw new Error('Pivot drill-down provenance requires sourceId and recordId for joined sources');
    }
    const sourceKey = multiSource ? path.sourceId! : (node.sourceId ?? '__single-source__');
    const recordId = path.recordId ?? `${path.sheetId}:${path.row}`;
    const record = records.get(recordId) ?? { paths: new Map<string, PivotSourceRowPath>() };
    if (record.paths.has(sourceKey)) throw new Error(`Pivot drill-down provenance repeats source ${sourceKey} in record ${recordId}`);
    record.paths.set(sourceKey, structuredClone(path));
    records.set(recordId, record);
  }
  if (multiSource) {
    if (!rootSourceId) throw new Error('Pivot drill-down source graph root is missing');
    for (const [recordId, record] of records) {
      for (const sourceId of requiredSourceIds) {
        if (!record.paths.has(sourceId)) throw new Error(`Pivot drill-down provenance is incomplete for record ${recordId}: ${sourceId}`);
      }
      if (!record.paths.has(rootSourceId)) throw new Error(`Pivot drill-down provenance is missing root source for record ${recordId}`);
    }
  }
  return { columns: drillDownColumns(context, pivot), records: [...records.values()] };
}

function drillDownColumns(context: CommandContext, pivot: PivotModel): DrillDownColumn[] {
  const columns: DrillDownColumn[] = [];
  const labels = new Set<string>();
  const nodes = pivotSourceNodes(pivot, context.workbook);
  for (const node of nodes) {
    const { range } = node;
    const sheet = context.workbook.getSheet(range.sheetId);
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const raw = sheet.cells.get(range.startRow, column)?.value;
      const base = raw == null || raw === '' ? `Column ${column - range.startColumn + 1}` : String(raw);
      let label = base;
      if (labels.has(label) && nodes.length > 1) label = `${sheet.name}.${base}`;
      let suffix = 2;
      while (labels.has(label)) label = `${base} (${suffix++})`;
      labels.add(label);
      columns.push({ sourceId: node.sourceId, range, column, label });
    }
  }
  return columns;
}

function writePivotDrillDown(context: CommandContext, params: PivotDrillDownParams): void {
  const sourceSheet = context.workbook.getSheet(params.sheetId);
  const pivot = sourceSheet.pivots.find((entry) => entry.id === params.pivotId);
  if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
  const plan = planPivotDrillDown(context, params);
  const target = context.workbook.addSheet(params.targetSheetId, createPivotDrillDownSheetName(pivot, params.label));
  plan.columns.forEach((column, index) => target.cells.set(params.target.row, params.target.column + index, { value: column.label }));
  plan.records.forEach((record, rowOffset) => {
    plan.columns.forEach((column, columnOffset) => {
      const sourceKey = column.sourceId ?? '__single-source__';
      const path = record.paths.get(sourceKey);
      const source = path ? context.workbook.getSheet(path.sheetId) : undefined;
      const value = source && path ? sourceCellValue(source.cells.get(path.row, column.column)) : null;
      target.cells.set(params.target.row + rowOffset + 1, params.target.column + columnOffset, { value });
    });
  });
}

function applyPivotUpdate(context: CommandContext, params: PivotUpdateParams): void {
  const pivot = pivotFor(context, params.sheetId, params.pivotId);
  if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
  const current = normalizePivotDefinition(context.workbook, pivot);
  const source = params.source ?? current.source;
  const next: PivotModel = {
    schema: 'PivotDefinition',
    id: current.id,
    source: structuredClone(source),
    target: structuredClone(params.target ?? current.target),
    fieldCatalog: structuredClone(params.fieldCatalog ?? current.fieldCatalog),
    layout: structuredClone(params.layout ?? current.layout),
    refreshPolicy: structuredClone(params.refreshPolicy ?? current.refreshPolicy),
    presentation: structuredClone(params.presentation ?? current.presentation),
    ...(params.nativeMetadata ?? current.nativeMetadata ? { nativeMetadata: structuredClone(params.nativeMetadata ?? current.nativeMetadata) } : {}),
  };
  assertPivotDefinition(context.workbook, next);
  const projection = buildPivotGridProjection(context.workbook, next);
  if (projection.collision.status === 'collision') {
    throw new Error(`Pivot target collision: ${projection.collision.reasons.join(', ')}`);
  }
  Object.assign(pivot, next);
}

function previousPivotUpdate(pivot: PivotModel): PivotUpdateParams {
  const definition = pivot as PivotDefinition;
  return {
    sheetId: definition.target.sheetId,
    pivotId: pivot.id,
    source: structuredClone(definition.source),
    target: structuredClone(definition.target),
    fieldCatalog: structuredClone(definition.fieldCatalog),
    refreshPolicy: structuredClone(definition.refreshPolicy),
    ...(definition.nativeMetadata ? { nativeMetadata: structuredClone(definition.nativeMetadata) } : {}),
    ...(definition.presentation ? { presentation: structuredClone(definition.presentation) } : {}),
    layout: structuredClone(pivot.layout),
  };
}

function pivotNodeIds(nodes: readonly import('@react-sheets/core-model').PivotResultNode[], target = new Set<string>()): Set<string> {
  for (const node of nodes) {
    if (node.nodeId) target.add(node.nodeId);
    pivotNodeIds(node.children, target);
  }
  return target;
}

function pivotExpansion(pivot: PivotModel): NonNullable<PivotLayout['expansion']> {
  return structuredClone(pivot.layout.expansion ?? { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true });
}

function applyPivotExpansionMutation(
  context: CommandContext,
  pivot: PivotModel,
  nextLayout: PivotLayout,
): void {
  applyPivotUpdate(context, { sheetId: pivot.target.sheetId, pivotId: pivot.id, layout: nextLayout });
}

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null;
const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.length > 0;
const isRange = (value: unknown): value is RangeRef => isRecord(value)
  && isNonEmptyString(value.sheetId)
  && Number.isSafeInteger(value.startRow) && Number.isSafeInteger(value.endRow)
  && Number.isSafeInteger(value.startColumn) && Number.isSafeInteger(value.endColumn)
  && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
  && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
const isPivotExpansion = (value: unknown): boolean => isRecord(value)
  && Object.keys(value).every((key) => ['expandedNodeIds', 'collapsedNodeIds', 'showButtons'].includes(key))
  && Array.isArray(value.expandedNodeIds) && value.expandedNodeIds.every(isNonEmptyString)
  && Array.isArray(value.collapsedNodeIds) && value.collapsedNodeIds.every(isNonEmptyString)
  && typeof value.showButtons === 'boolean';
const isPivotLayout = (value: unknown): value is PivotLayout => isRecord(value)
  && Array.isArray(value.rows) && Array.isArray(value.columns) && Array.isArray(value.filters)
  && Array.isArray(value.values)
  && ['top', 'bottom', 'off'].includes(String(value.subtotalLocation)) && typeof value.showGrandTotals === 'boolean'
  && typeof value.compact === 'boolean' && typeof value.repeatLabels === 'boolean'
  && (value.expansion === undefined || isPivotExpansion(value.expansion));
const isPivotPresentation = (value: unknown): value is PivotPresentation => isRecord(value)
  && Object.keys(value).every((key) => key === 'styleName' || key === 'styleOptions')
  && (value.styleName === undefined || typeof value.styleName === 'string')
  && isRecord(value.styleOptions)
  && Object.keys(value.styleOptions).every((key) => ['showRowHeaders', 'showColumnHeaders', 'showRowStripes', 'showColumnStripes', 'showLastColumn'].includes(key))
  && typeof value.styleOptions.showRowHeaders === 'boolean'
  && typeof value.styleOptions.showColumnHeaders === 'boolean'
  && typeof value.styleOptions.showRowStripes === 'boolean'
  && typeof value.styleOptions.showColumnStripes === 'boolean'
  && typeof value.styleOptions.showLastColumn === 'boolean';
const isPivotSourceRange = (value: unknown): boolean => isRecord(value)
  && isNonEmptyString(value.sourceId) && isRange(value.range);
const isPivotWorksheetSource = (value: unknown): boolean => {
  if (!isRecord(value)) return false;
  if (value.kind === 'worksheet-range') return isRange(value.range);
  if (value.kind === 'worksheet-ranges') return Array.isArray(value.ranges) && value.ranges.length > 0 && value.ranges.every(isPivotSourceRange)
    && Array.isArray(value.relationships);
  return false;
};
const isPivotSource = (value: unknown): value is PivotSource => isPivotWorksheetSource(value)
  || (isRecord(value) && value.kind === 'table' && isNonEmptyString(value.tableId))
  || (isRecord(value) && value.kind === 'named-range' && isNonEmptyString(value.name)
    && (value.sheetId === undefined || isNonEmptyString(value.sheetId)))
  || (isRecord(value) && value.kind === 'data-source' && isNonEmptyString(value.dataSourceId));
const isPivotModel = (value: unknown): value is PivotModel => isRecord(value)
  && isNonEmptyString(value.id)
  && value.schema === 'PivotDefinition'
  && isPivotLayout(value.layout)
  && isPivotSource(value.source)
  && isRecord(value.target) && isNonEmptyString(value.target.sheetId) && isRecord(value.target.anchor)
    && Number.isSafeInteger(value.target.anchor.row) && Number(value.target.anchor.row) >= 0
    && Number.isSafeInteger(value.target.anchor.column) && Number(value.target.anchor.column) >= 0
  && isRecord(value.fieldCatalog) && Array.isArray(value.fieldCatalog.fields)
  && isRecord(value.refreshPolicy) && ['manual', 'on-open', 'on-change'].includes(String(value.refreshPolicy.mode))
  && (value.presentation === undefined || isPivotPresentation(value.presentation));
const isPivotCreateDestination = (value: unknown): value is PivotCreateDestination => {
  if (!isRecord(value) || !isNonEmptyString(value.kind) || !isNonEmptyString(value.sheetId)) return false;
  if (value.kind === 'existing-sheet') return Object.keys(value).every((key) => key === 'kind' || key === 'sheetId');
  if (value.kind !== 'new-sheet' || !isNonEmptyString(value.name)) return false;
  return (value.rowCount === undefined || (Number.isSafeInteger(value.rowCount) && Number(value.rowCount) > 0))
    && (value.columnCount === undefined || (Number.isSafeInteger(value.columnCount) && Number(value.columnCount) > 0));
};
const isPivotCreate = (value: unknown): value is PivotCreateParams => isRecord(value)
  && isPivotModel(value.pivot)
  && isPivotCreateDestination(value.destination);
const isPivotUpdate = (value: unknown): value is PivotUpdateParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && (value.source === undefined || isPivotSource(value.source))
  && (value.target === undefined || (isRecord(value.target) && isNonEmptyString(value.target.sheetId) && isRecord(value.target.anchor)))
  && (value.layout === undefined || isPivotLayout(value.layout))
  && (value.presentation === undefined || isPivotPresentation(value.presentation))
  && ['source', 'target', 'fieldCatalog', 'refreshPolicy', 'nativeMetadata', 'presentation', 'layout'].some((key) => value[key] !== undefined);
const isPivotDrillDown = (value: unknown): value is PivotDrillDownParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && typeof value.label === 'string' && isNonEmptyString(value.targetSheetId)
  && Array.isArray(value.sourceRowPaths)
  && value.sourceRowPaths.every((entry) => isRecord(entry) && (entry.sourceId === undefined || isNonEmptyString(entry.sourceId)) && (entry.recordId === undefined || isNonEmptyString(entry.recordId)) && isNonEmptyString(entry.sheetId) && Number.isSafeInteger(entry.row) && Number(entry.row) >= 0)
  && isRecord(value.target)
  && Number.isSafeInteger(value.target.row) && Number.isSafeInteger(value.target.column)
  && Number(value.target.row) >= 0 && Number(value.target.column) >= 0;
const isPivotDrillDownRemove = (value: unknown): value is PivotDrillDownRemoveParams => isRecord(value) && isNonEmptyString(value.targetSheetId);

function pivotMutationRanges(value: unknown): RangeRef[] {
  if (isPivotModel(value)) return pivotSourceRanges(value);
  if (isPivotUpdate(value)) {
    if (value.source?.kind === 'worksheet-ranges') return value.source.ranges.map((sourceRange) => structuredClone(sourceRange.range));
    if (value.source?.kind === 'worksheet-range') return [structuredClone(value.source.range)];
    return sheetRange(value.sheetId);
  }
  if (isRecord(value) && isNonEmptyString(value.sheetId)) return sheetRange(value.sheetId);
  return [];
}

interface PivotCreatePlan {
  pivot: PivotModel;
  destination: PivotCreateDestination;
  affectedRanges: RangeRef[];
}

function assertPivotSourceHeaders(workbook: WorkbookModel, pivot: PivotModel): void {
  const ranges = getPivotSourceRanges(workbook, pivot);
  for (const range of ranges) {
    const sheet = workbook.getSheet(range.sheetId);
    const names = new Set<string>();
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const raw = sheet.cells.get(range.startRow, column)?.formulaValue ?? sheet.cells.get(range.startRow, column)?.value;
      const name = raw == null ? '' : String(raw).trim();
      if (!name) throw new Error(`Pivot source header is blank at ${range.sheetId}:${range.startRow}:${column}`);
      const key = name.toLocaleLowerCase();
      if (names.has(key)) throw new Error(`Pivot source header is duplicated: ${name}`);
      names.add(key);
    }
  }
}

function assertPivotCreateIdentity(workbook: WorkbookModel, params: PivotCreateParams): void {
  if (workbook.getSheets().some((sheet) => sheet.pivots.some((pivot) => pivot.id === params.pivot.id))) {
    throw new Error(`Pivot already exists: ${params.pivot.id}`);
  }
  if (params.destination.kind === 'existing-sheet') {
    if (params.pivot.target.sheetId !== params.destination.sheetId) throw new Error('Pivot destination sheet does not match the target');
    workbook.getSheet(params.destination.sheetId);
    return;
  }
  if (params.pivot.target.sheetId !== params.destination.sheetId) throw new Error('Pivot destination sheet does not match the target');
  if (!params.destination.sheetId.trim()) throw new Error('Pivot destination sheet id is required');
  const name = params.destination.name.trim();
  if (!name) throw new Error('Pivot destination sheet name is required');
  const folded = name.toLocaleLowerCase();
  if (workbook.sheets.has(params.destination.sheetId)) throw new Error(`Sheet already exists: ${params.destination.sheetId}`);
  if (workbook.getSheets().some((sheet) => sheet.name.trim().toLocaleLowerCase() === folded)) {
    throw new Error(`Sheet name already exists: ${name}`);
  }
  const dimensions: Array<readonly [string, number | undefined]> = [
    ['rowCount', params.destination.rowCount],
    ['columnCount', params.destination.columnCount],
  ];
  for (const [key, value] of dimensions) {
    if (value !== undefined && (!Number.isSafeInteger(value) || value <= 0)) throw new Error(`Pivot destination ${key} is invalid`);
  }
}

/** Build and validate the complete create operation without touching the live workbook. */
function planPivotCreate(workbook: WorkbookModel, params: PivotCreateParams): PivotCreatePlan {
  if (!isPivotCreate(params)) throw new Error('Invalid pivot.create payload');
  assertPivotCreateIdentity(workbook, params);

  const preflight = WorkbookModel.fromSnapshot(workbook.snapshot());
  if (params.destination.kind === 'new-sheet') {
    preflight.addSheet(
      params.destination.sheetId,
      params.destination.name.trim(),
      params.destination.rowCount,
      params.destination.columnCount,
    );
  }
  const candidate = structuredClone(params.pivot);
  assertPivotDefinition(preflight, candidate);
  assertPivotSourceHeaders(preflight, candidate);
  const canonical = structuredClone(normalizePivotDefinition(preflight, candidate));
  if (canonical.fieldCatalog.fields.length === 0) throw new Error('PivotTable source does not contain usable fields');
  assertPivotDefinition(preflight, canonical);
  const projection = buildPivotGridProjection(preflight, canonical);
  if (projection.collision.status === 'collision') {
    throw new Error(`Pivot target collision: ${projection.collision.reasons.join(', ')}`);
  }
  return {
    pivot: canonical,
    destination: structuredClone(params.destination),
    affectedRanges: sheetRange(canonical.target.sheetId),
  };
}

function applyPivotAdd(context: CommandContext, params: PivotModel): void {
  if (context.workbook.getSheets().some((candidate) => candidate.pivots.some((entry) => entry.id === params.id))) throw new Error(`Pivot already exists: ${params.id}`);
  const definition = normalizePivotDefinition(context.workbook, params);
  const canonical: PivotModel = structuredClone(definition);
  assertPivotDefinition(context.workbook, canonical);
  const projection = buildPivotGridProjection(context.workbook, canonical);
  if (projection.collision.status === 'collision') throw new Error(`Pivot target collision: ${projection.collision.reasons.join(', ')}`);
  context.workbook.getSheet(definition.target.sheetId).pivots.push(canonical);
}

function applyPivotRemove(context: CommandContext, params: string, sheetId: string): void {
  const sheet = context.workbook.getSheet(sheetId);
  if (removeById(sheet.pivots, params)) return;
  for (const candidate of context.workbook.getSheets()) if (candidate !== sheet && removeById(candidate.pivots, params)) return;
  throw new Error(`Unknown pivot: ${params}`);
}

export function registerPivotCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  runtime.registry.registerMutation<PivotModel>({
      id: 'pivot.add',
      handler: (item, context) => {
    applyPivotAdd(context, item.params);
  },
      metadata: {
    schema: { name: 'PivotModel', validate: isPivotModel },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<string>({
      id: 'pivot.remove',
      handler: (item, context) => {
    applyPivotRemove(context, item.params, item.sheetId);
  },
      metadata: {
    schema: { name: 'PivotId', validate: isNonEmptyString },
    permission: { capability: 'pivot.delete' },
    affectedRanges: { resolve: () => [], mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.add'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<PivotUpdateParams>({
      id: 'pivot.update',
      handler: (item, context) => {
    applyPivotUpdate(context, item.params);
  },
      metadata: {
    schema: { name: 'PivotUpdateParams', validate: isPivotUpdate },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.update'], minCount: 1, maxCount: 1 },
  },
    });

  runtime.registry.registerCommand<PivotModel>({
    id: 'pivot.add',
    execute: (params, context) => {
      const affectedRanges = sheetRange(pivotDisplaySheetId(params));
      context.applyMutation({
        id: 'pivot.add',
        unitId: context.workbook.unitId,
        sheetId: pivotDisplaySheetId(params),
        params: structuredClone(params),
        affectedRanges,
        inverse: [{ id: 'pivot.remove', unitId: context.workbook.unitId, sheetId: pivotDisplaySheetId(params), params: params.id, affectedRanges }],
        apply: () => applyPivotAdd(context, structuredClone(params)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.add');

  runtime.registry.registerCommand<PivotCreateParams>({
    id: 'pivot.create',
    execute: (params, context) => {
      const plan = planPivotCreate(context.workbook, params);
      const { pivot, destination, affectedRanges } = plan;
      if (destination.kind === 'new-sheet') {
        const sheetParams = {
          id: destination.sheetId,
          name: destination.name.trim(),
          ...(destination.rowCount === undefined ? {} : { rowCount: destination.rowCount }),
          ...(destination.columnCount === undefined ? {} : { columnCount: destination.columnCount }),
        };
        context.applyMutation({
          id: 'sheet.add',
          unitId: context.workbook.unitId,
          sheetId: destination.sheetId,
          params: sheetParams,
          affectedRanges: [],
          inverse: [{
            id: 'sheet.remove',
            unitId: context.workbook.unitId,
            sheetId: destination.sheetId,
            params: { id: destination.sheetId },
            affectedRanges: [],
          }],
          apply: () => {
            context.workbook.addSheet(
              sheetParams.id,
              sheetParams.name,
              sheetParams.rowCount,
              sheetParams.columnCount,
            );
          },
        });
      }
      context.applyMutation({
        id: 'pivot.add',
        unitId: context.workbook.unitId,
        sheetId: pivot.target.sheetId,
        params: structuredClone(pivot),
        affectedRanges,
        inverse: [{
          id: 'pivot.remove',
          unitId: context.workbook.unitId,
          sheetId: pivot.target.sheetId,
          params: pivot.id,
          affectedRanges,
        }],
        apply: () => applyPivotAdd(context, structuredClone(pivot)),
      });
      return { operationId: context.operationId, mutationCount: destination.kind === 'new-sheet' ? 2 : 1, affectedRanges };
    },
  });
  commandIds.push('pivot.create');

  runtime.registry.registerCommand<string | { sheetId: string; pivotId: string }>({
    id: 'pivot.remove',
    execute: (input, context) => {
      const sheetId = typeof input === 'string' ? context.workbook.primarySheetId : input.sheetId;
      const pivotId = typeof input === 'string' ? input : input.pivotId;
      const pivot = pivotFor(context, sheetId, pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${pivotId}`);
      const affectedRanges = sheetRange(sheetId);
      context.applyMutation({
        id: 'pivot.remove',
        unitId: context.workbook.unitId,
        sheetId,
        params: pivotId,
        affectedRanges,
        inverse: [{ id: 'pivot.add', unitId: context.workbook.unitId, sheetId, params: structuredClone(pivot), affectedRanges }],
        apply: () => applyPivotRemove(context, pivotId, sheetId),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.remove');

  runtime.registry.registerCommand<PivotUpdateParams>({
    id: 'pivot.update',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      const previous = previousPivotUpdate(pivot);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'pivot.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: structuredClone(params),
        affectedRanges,
        inverse: [{ id: 'pivot.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => applyPivotUpdate(context, structuredClone(params)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.update');

  const registerExpansionCommand = <P extends { sheetId: string; pivotId: string }>(
    commandId: string,
    update: (pivot: PivotModel, params: P, tree: import('@react-sheets/core-model').PivotResultTree | undefined) => PivotLayout,
    requiresTree = true,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const pivot = pivotFor(context, params.sheetId, params.pivotId);
        if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
        const tree = requiresTree ? computePivotResult(context.workbook, pivot) : undefined;
        const nextLayout = update(pivot, params, tree);
        const affectedRanges = sheetRange(params.sheetId);
        const previousLayout = structuredClone(pivot.layout);
        context.applyMutation({
          id: 'pivot.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, pivotId: params.pivotId, layout: structuredClone(nextLayout) },
          affectedRanges,
          inverse: [{ id: 'pivot.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, pivotId: params.pivotId, layout: previousLayout }, affectedRanges }],
          apply: () => applyPivotExpansionMutation(context, pivot, structuredClone(nextLayout)),
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges };
      },
    });
    commandIds.push(commandId);
  };

  registerExpansionCommand<PivotExpansionToggleParams>('pivot.expansion.toggle', (pivot, params, tree) => {
    if (!params.nodeId.trim()) throw new Error('Pivot expansion node id is required');
    if (!tree) throw new Error('Pivot expansion result is unavailable');
    const known = pivotNodeIds(tree.rows);
    if (!known.has(params.nodeId)) throw new Error(`Unknown Pivot expansion node: ${params.nodeId}`);
    const expansion = pivotExpansion(pivot);
    const collapsed = new Set(expansion.collapsedNodeIds);
    const expanded = new Set(expansion.expandedNodeIds);
    if (collapsed.has(params.nodeId)) {
      collapsed.delete(params.nodeId);
      expanded.add(params.nodeId);
    } else {
      collapsed.add(params.nodeId);
      expanded.delete(params.nodeId);
    }
    return { ...structuredClone(pivot.layout), expansion: { ...expansion, expandedNodeIds: [...expanded], collapsedNodeIds: [...collapsed] } };
  });

  registerExpansionCommand<PivotExpansionFieldParams>('pivot.expansion.field', (pivot, params, tree) => {
    if (!params.fieldId.trim()) throw new Error('Pivot expansion field id is required');
    if (!tree) throw new Error('Pivot expansion result is unavailable');
    const nodes: import('@react-sheets/core-model').PivotResultNode[] = [];
    const collect = (entries: readonly import('@react-sheets/core-model').PivotResultNode[]) => entries.forEach((node) => { if (node.fieldId === params.fieldId && node.children.length) nodes.push(node); collect(node.children); });
    collect(tree.rows);
    if (nodes.length === 0) throw new Error(`Unknown Pivot expansion field: ${params.fieldId}`);
    const expansion = pivotExpansion(pivot);
    const collapsed = new Set(expansion.collapsedNodeIds);
    const expanded = new Set(expansion.expandedNodeIds);
    for (const node of nodes) {
      if (params.expanded) { collapsed.delete(node.nodeId ?? ''); expanded.add(node.nodeId ?? ''); }
      else { collapsed.add(node.nodeId ?? ''); expanded.delete(node.nodeId ?? ''); }
    }
    return { ...structuredClone(pivot.layout), expansion: { ...expansion, expandedNodeIds: [...expanded], collapsedNodeIds: [...collapsed] } };
  });

  registerExpansionCommand<PivotExpansionButtonsParams>('pivot.expansion.buttons', (pivot, params) => ({
    ...structuredClone(pivot.layout),
    expansion: { ...pivotExpansion(pivot), showButtons: params.showButtons },
  }), false);

  // Refresh is intentionally a derived operation: it validates and rebuilds
  // the projection without persisting a stale result cache or a fake counter.
  runtime.registry.registerCommand<{ sheetId: string; pivotId: string }>({
    id: 'pivot.refresh',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      const projection = buildPivotGridProjection(context.workbook, pivot);
      if (projection.refresh.status === 'error') throw new Error(projection.refresh.error ?? 'PivotTable refresh failed');
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [projection.occupiedRange] };
    },
  });
  commandIds.push('pivot.refresh');

  runtime.registry.registerMutation<PivotDrillDownParams>({
      id: 'pivot.drilldown.add',
      handler: (item, context) => {
    writePivotDrillDown(context, item.params);
  },
      metadata: {
    schema: { name: 'PivotDrillDownParams', validate: isPivotDrillDown },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.drilldown.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<PivotDrillDownRemoveParams>({
      id: 'pivot.drilldown.remove',
      handler: (item, context) => {
    if (!context.workbook.sheets.has(item.params.targetSheetId)) throw new Error(`Unknown drill-down target: ${item.params.targetSheetId}`);
    context.workbook.removeSheet(item.params.targetSheetId);
  },
      metadata: {
    schema: { name: 'PivotDrillDownRemoveParams', validate: isPivotDrillDownRemove },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: (value) => isPivotDrillDownRemove(value) ? sheetRange(value.targetSheetId) : [], mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.drilldown.add'], minCount: 1, maxCount: 1 },
  },
    });

  const registerLayoutPatch = <P extends { sheetId: string; pivotId: string; fieldId: string }>(
    commandId: string,
    buildLayout: (layout: PivotLayout, params: P) => PivotLayout,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const pivot = pivotFor(context, params.sheetId, params.pivotId);
        if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
        assertPivotField(context.workbook, pivot, params.fieldId);
        const previousLayout = structuredClone(pivot.layout);
        const nextLayout = buildLayout(previousLayout, params);
        const affectedRanges = sheetRange(params.sheetId);
        context.applyMutation({
          id: 'pivot.update',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, pivotId: params.pivotId, layout: nextLayout },
          affectedRanges,
          inverse: [{ id: 'pivot.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, pivotId: params.pivotId, layout: previousLayout }, affectedRanges }],
          apply: () => applyPivotUpdate(context, { sheetId: params.sheetId, pivotId: params.pivotId, layout: nextLayout }),
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges };
      },
    });
    commandIds.push(commandId);
  };

  registerLayoutPatch<PivotAggregateParams>('pivot.setAggregate', (layout, params) => setPivotAggregate(layout, params.fieldId, params.summarizeBy));
  registerLayoutPatch<PivotShowAsParams>('pivot.setShowAs', (layout, params) => setPivotShowAs(layout, params.fieldId, params.showAs));
  registerLayoutPatch<PivotGroupParams>('pivot.setGroup', (layout, params) => setPivotGroup(layout, params.axis, params.fieldId, params.group));

  runtime.registry.registerCommand<PivotChartCreateParams>({
    id: 'pivot.chart.create',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      if (params.drawing.sheetId !== params.sheetId || params.drawing.kind !== 'chart'
        || params.payload.kind !== 'chart' || params.drawing.payloadId !== params.payload.chartId
        || params.payload.pivotId !== pivot.id) {
        throw new Error('PivotChart drawing and payload identity mismatch');
      }
      const sheet = context.workbook.getSheet(params.sheetId);
      if (sheet.drawings.some((entry) => entry.id === params.drawing.id) || sheet.drawingPayloads.has(params.drawing.payloadId)) {
        throw new Error(`PivotChart already exists: ${params.drawing.id}`);
      }
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'drawing.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { sheetId: params.sheetId, drawing: structuredClone(params.drawing), payload: structuredClone(params.payload) },
        affectedRanges,
        inverse: [{
          id: 'drawing.remove',
          unitId: context.workbook.unitId,
          sheetId: params.sheetId,
          params: { sheetId: params.sheetId, drawingId: params.drawing.id },
          affectedRanges,
        }],
        apply: () => {
          const destination = context.workbook.getSheet(params.sheetId);
          destination.drawings.push(structuredClone(params.drawing));
          destination.drawingPayloads.set(params.drawing.payloadId, structuredClone(params.payload));
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.chart.create');

  runtime.registry.registerCommand<PivotDrillDownParams>({
    id: 'pivot.drillDown',
    execute: (params, context) => {
      const plan = planPivotDrillDown(context, params);
      const columns = plan.columns.length;
      const detailRows = plan.records.length;
      const affectedRanges: RangeRef[] = [
        { sheetId: params.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        { sheetId: params.targetSheetId, startRow: params.target.row, endRow: params.target.row + detailRows, startColumn: params.target.column, endColumn: params.target.column + Math.max(columns - 1, 0) },
      ];
      context.applyMutation({
        id: 'pivot.drilldown.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: structuredClone(params),
        affectedRanges,
        inverse: [{
          id: 'pivot.drilldown.remove',
          unitId: context.workbook.unitId,
          sheetId: params.targetSheetId,
          params: { targetSheetId: params.targetSheetId },
          affectedRanges,
        }],
        apply: () => writePivotDrillDown(context, structuredClone(params)),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.drillDown');

  return commandIds;
}

export const PIVOT_MUTATION_IDS = ['pivot.add', 'pivot.remove', 'pivot.update', 'pivot.drilldown.add', 'pivot.drilldown.remove'] as const;
