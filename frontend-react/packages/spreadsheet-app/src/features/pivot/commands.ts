import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type {
  PivotAggregateFunction,
  PivotDefinition,
  PivotSource,
  PivotGroup,
  PivotLayout,
  PivotModel,
  PivotShowAs,
  ChartDrawingPayload,
  DrawingObject,
  RangeRef,
} from '@react-sheets/core-model';
import { assertPivotDefinition, assertPivotField, createPivotDrillDownSheetName, setPivotAggregate, setPivotGroup, setPivotShowAs } from './panel-state';
import { buildPivotGridProjection, getPivotSourceRanges, normalizePivotDefinition } from './engine';

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
  layout?: PivotLayout;
}

export interface PivotLayoutCommandParams {
  sheetId: string;
  pivotId: string;
  layout: PivotLayout;
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
  sourceRowPaths: Array<{ sheetId: string; row: number }>;
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
  if (pivot.source.kind === 'worksheet-ranges') return structuredClone(pivot.source.ranges);
  if (pivot.source.kind === 'worksheet-range') return [structuredClone(pivot.source.range)];
  return [];
}

interface DrillDownColumn {
  range: RangeRef;
  column: number;
  label: string;
}

function sourceCellValue(cell: { value: unknown; formulaValue?: unknown } | undefined): string | number | boolean | null {
  const value = cell?.formulaValue ?? cell?.value ?? null;
  return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || value === null ? value : null;
}

function drillDownColumns(context: CommandContext, pivot: PivotModel): DrillDownColumn[] {
  const columns: DrillDownColumn[] = [];
  const labels = new Set<string>();
  const ranges = pivotSourceRanges(pivot, context.workbook);
  for (const range of ranges) {
    const sheet = context.workbook.getSheet(range.sheetId);
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const raw = sheet.cells.get(range.startRow, column)?.value;
      const base = raw == null || raw === '' ? `Column ${column - range.startColumn + 1}` : String(raw);
      let label = base;
      if (labels.has(label) && ranges.length > 1) label = `${sheet.name}.${base}`;
      let suffix = 2;
      while (labels.has(label)) label = `${base} (${suffix++})`;
      labels.add(label);
      columns.push({ range, column, label });
    }
  }
  return columns;
}

function writePivotDrillDown(context: CommandContext, params: PivotDrillDownParams): void {
  const sourceSheet = context.workbook.getSheet(params.sheetId);
  const pivot = sourceSheet.pivots.find((entry) => entry.id === params.pivotId);
  if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
  if (context.workbook.sheets.has(params.targetSheetId)) throw new Error(`Drill-down target already exists: ${params.targetSheetId}`);
  for (const range of pivotSourceRanges(pivot, context.workbook)) {
    const sheet = context.workbook.getSheet(range.sheetId);
    if (range.startRow < 0 || range.endRow >= sheet.rowCount || range.startColumn < 0 || range.endColumn >= sheet.columnCount) {
      throw new Error('Pivot source range exceeds worksheet bounds');
    }
  }
  for (const path of params.sourceRowPaths) {
    const sheet = context.workbook.getSheet(path.sheetId);
    if (path.row < 0 || path.row >= sheet.rowCount) throw new Error(`Pivot drill-down source row is invalid: ${path.row}`);
  }

  const columns = drillDownColumns(context, pivot);
  const target = context.workbook.addSheet(params.targetSheetId, createPivotDrillDownSheetName(pivot, params.label));
  columns.forEach((column, index) => target.cells.set(params.target.row, params.target.column + index, { value: column.label }));

  const ranges = pivotSourceRanges(pivot, context.workbook);
  const rowsPerResult = Math.max(ranges.length, 1);
  const resultRowCount = Math.ceil(params.sourceRowPaths.length / rowsPerResult);
  for (let rowOffset = 0; rowOffset < resultRowCount; rowOffset += 1) {
    // Joined pivots flatten one path per source range into the result tree.
    // Re-group those paths deterministically before reading detail values.
    const paths = params.sourceRowPaths.slice(rowOffset * rowsPerResult, (rowOffset + 1) * rowsPerResult);
    columns.forEach((column, columnOffset) => {
      const path = paths.find((entry) => entry.sheetId === column.range.sheetId);
      const source = path ? context.workbook.getSheet(path.sheetId) : undefined;
      const value = source && path ? sourceCellValue(source.cells.get(path.row, column.column)) : null;
      target.cells.set(params.target.row + rowOffset + 1, params.target.column + columnOffset, { value });
    });
  }
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
    layout: structuredClone(pivot.layout),
  };
}

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null;
const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.length > 0;
const isRange = (value: unknown): value is RangeRef => isRecord(value)
  && isNonEmptyString(value.sheetId)
  && Number.isSafeInteger(value.startRow) && Number.isSafeInteger(value.endRow)
  && Number.isSafeInteger(value.startColumn) && Number.isSafeInteger(value.endColumn)
  && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
  && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
const isPivotLayout = (value: unknown): value is PivotLayout => isRecord(value)
  && Array.isArray(value.rows) && Array.isArray(value.columns) && Array.isArray(value.filters)
  && Array.isArray(value.values)
  && typeof value.showSubtotals === 'boolean' && typeof value.showGrandTotals === 'boolean'
  && typeof value.compact === 'boolean' && typeof value.repeatLabels === 'boolean';
const isPivotWorksheetSource = (value: unknown): boolean => {
  if (!isRecord(value)) return false;
  if (value.kind === 'worksheet-range') return isRange(value.range);
  if (value.kind === 'worksheet-ranges') return Array.isArray(value.ranges) && value.ranges.every(isRange)
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
  && isRecord(value.refreshPolicy) && ['manual', 'on-open', 'on-change'].includes(String(value.refreshPolicy.mode));
const isPivotUpdate = (value: unknown): value is PivotUpdateParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && (value.source === undefined || isPivotSource(value.source))
  && (value.target === undefined || (isRecord(value.target) && isNonEmptyString(value.target.sheetId) && isRecord(value.target.anchor)))
  && (value.layout === undefined || isPivotLayout(value.layout))
  && ['source', 'target', 'fieldCatalog', 'refreshPolicy', 'nativeMetadata', 'layout'].some((key) => value[key] !== undefined);
const isPivotDrillDown = (value: unknown): value is PivotDrillDownParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && typeof value.label === 'string' && isNonEmptyString(value.targetSheetId)
  && Array.isArray(value.sourceRowPaths)
  && value.sourceRowPaths.every((entry) => isRecord(entry) && isNonEmptyString(entry.sheetId) && Number.isSafeInteger(entry.row) && Number(entry.row) >= 0)
  && isRecord(value.target)
  && Number.isSafeInteger(value.target.row) && Number.isSafeInteger(value.target.column)
  && Number(value.target.row) >= 0 && Number(value.target.column) >= 0;
const isPivotDrillDownRemove = (value: unknown): value is PivotDrillDownRemoveParams => isRecord(value) && isNonEmptyString(value.targetSheetId);

function pivotMutationRanges(value: unknown): RangeRef[] {
  if (isPivotModel(value)) return pivotSourceRanges(value);
  if (isPivotUpdate(value)) {
    if (value.source?.kind === 'worksheet-ranges') return structuredClone(value.source.ranges);
    if (value.source?.kind === 'worksheet-range') return [structuredClone(value.source.range)];
    return sheetRange(value.sheetId);
  }
  if (isRecord(value) && isNonEmptyString(value.sheetId)) return sheetRange(value.sheetId);
  return [];
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

  runtime.registry.registerMutation<PivotModel>('pivot.add', (item, context) => {
    applyPivotAdd(context, item.params);
  }, {
    schema: { name: 'PivotModel', validate: isPivotModel },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.remove'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<string>('pivot.remove', (item, context) => {
    applyPivotRemove(context, item.params, item.sheetId);
  }, {
    schema: { name: 'PivotId', validate: isNonEmptyString },
    permission: { capability: 'pivot.delete' },
    affectedRanges: { resolve: () => [], mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.add'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<PivotUpdateParams>('pivot.update', (item, context) => {
    applyPivotUpdate(context, item.params);
  }, {
    schema: { name: 'PivotUpdateParams', validate: isPivotUpdate },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.update'], minCount: 1, maxCount: 1 },
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

  runtime.registry.registerMutation<PivotDrillDownParams>('pivot.drilldown.add', (item, context) => {
    writePivotDrillDown(context, item.params);
  }, {
    schema: { name: 'PivotDrillDownParams', validate: isPivotDrillDown },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.drilldown.remove'], minCount: 1, maxCount: 1 },
  });
  runtime.registry.registerMutation<PivotDrillDownRemoveParams>('pivot.drilldown.remove', (item, context) => {
    if (!context.workbook.sheets.has(item.params.targetSheetId)) throw new Error(`Unknown drill-down target: ${item.params.targetSheetId}`);
    context.workbook.removeSheet(item.params.targetSheetId);
  }, {
    schema: { name: 'PivotDrillDownRemoveParams', validate: isPivotDrillDownRemove },
    permission: { capability: 'pivot.edit' },
    affectedRanges: { resolve: (value) => isPivotDrillDownRemove(value) ? sheetRange(value.targetSheetId) : [], mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.drilldown.add'], minCount: 1, maxCount: 1 },
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
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      if (context.workbook.sheets.has(params.targetSheetId)) throw new Error(`Drill-down target already exists: ${params.targetSheetId}`);
      const columns = drillDownColumns(context, pivot).length;
      const sourceRangeCount = Math.max(pivotSourceRanges(pivot, context.workbook).length, 1);
      const detailRows = Math.ceil(params.sourceRowPaths.length / sourceRangeCount);
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
