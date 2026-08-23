import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';
import type {
  PivotAggregateFunction,
  PivotChartReference,
  PivotDataSource,
  PivotGroup,
  PivotLayout,
  PivotModel,
  PivotShowAs,
  PivotSlicer,
  PivotTimeline,
  RangeRef,
} from '@react-sheets/core-model';
import { assertPivotDefinition, assertPivotField, createPivotDrillDownSheetName, setPivotAggregate, setPivotGroup, setPivotShowAs, upsertPivotSlicer, upsertPivotTimeline } from './panel-state';

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
  sourceRange?: RangeRef;
  dataSource?: PivotDataSource;
  layout?: PivotLayout;
  slicers?: PivotSlicer[];
  timelines?: PivotTimeline[];
  chartReferences?: PivotChartReference[];
}

export interface PivotRefreshParams {
  sheetId: string;
  pivotId: string;
  refreshRevision: number;
  lastRefreshedAt: string;
}

export interface PivotLayoutCommandParams {
  sheetId: string;
  pivotId: string;
  layout: PivotLayout;
}

export interface PivotAggregateParams {
  sheetId: string;
  pivotId: string;
  field: string;
  summarizeBy: PivotAggregateFunction;
}

export interface PivotShowAsParams {
  sheetId: string;
  pivotId: string;
  field: string;
  showAs: PivotShowAs;
}

export interface PivotGroupParams {
  sheetId: string;
  pivotId: string;
  axis: 'rows' | 'columns';
  field: string;
  group: PivotGroup;
}

export interface PivotDrillDownParams {
  sheetId: string;
  pivotId: string;
  label: string;
  sourceRowPaths: Array<{ sheetId: string; row: number }>;
  targetSheetId: string;
  targetAnchor: { row: number; column: number };
}

export interface PivotDrillDownRemoveParams {
  targetSheetId: string;
}

export interface PivotSlicerParams {
  sheetId: string;
  pivotId: string;
  slicer: PivotSlicer;
}

export interface PivotTimelineParams {
  sheetId: string;
  pivotId: string;
  timeline: PivotTimeline;
}

function pivotFor(context: CommandContext, sheetId: string, pivotId: string): PivotModel | undefined {
  return context.workbook.getSheet(sheetId).pivots.find((entry) => entry.id === pivotId);
}

function pivotSourceRanges(pivot: PivotModel): RangeRef[] {
  if (pivot.dataSource?.kind === 'worksheet-ranges') return structuredClone(pivot.dataSource.ranges);
  return [structuredClone(pivot.dataSource?.range ?? pivot.sourceRange)];
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
  for (const range of pivotSourceRanges(pivot)) {
    const sheet = context.workbook.getSheet(range.sheetId);
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const raw = sheet.cells.get(range.startRow, column)?.value;
      const base = raw == null || raw === '' ? `Column ${column - range.startColumn + 1}` : String(raw);
      let label = base;
      if (labels.has(label) && pivotSourceRanges(pivot).length > 1) label = `${sheet.name}.${base}`;
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
  for (const range of pivotSourceRanges(pivot)) {
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
  columns.forEach((column, index) => target.cells.set(params.targetAnchor.row, params.targetAnchor.column + index, { value: column.label }));

  const ranges = pivotSourceRanges(pivot);
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
      target.cells.set(params.targetAnchor.row + rowOffset + 1, params.targetAnchor.column + columnOffset, { value });
    });
  }
}

function applyPivotUpdate(context: CommandContext, params: PivotUpdateParams): void {
  const pivot = pivotFor(context, params.sheetId, params.pivotId);
  if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
  const next = structuredClone(pivot);
  if (params.layout) next.layout = structuredClone(params.layout);
  if (params.sourceRange) {
    next.sourceRange = structuredClone(params.sourceRange);
    next.fieldCatalog = undefined;
    // A single-range source update must not leave a stale dataSource pointing
    // at the previous worksheet/range.  Multi-range sources are replaced only
    // when the caller supplies the explicit dataSource below.
    if (next.dataSource?.kind === 'worksheet-range') next.dataSource = { kind: 'worksheet-range', range: structuredClone(params.sourceRange) };
    else if (next.dataSource?.kind === 'worksheet-ranges' && !params.dataSource) next.dataSource = undefined;
  }
  if (params.dataSource) next.dataSource = structuredClone(params.dataSource);
  if (params.slicers) next.slicers = structuredClone(params.slicers);
  if (params.timelines) next.timelines = structuredClone(params.timelines);
  if (params.chartReferences) next.chartReferences = structuredClone(params.chartReferences);
  assertPivotDefinition(context.workbook, next);
  Object.assign(pivot, next);
}

function previousPivotUpdate(pivot: PivotModel): PivotUpdateParams {
  return {
    sheetId: pivot.sheetId,
    pivotId: pivot.id,
    sourceRange: structuredClone(pivot.sourceRange),
    dataSource: structuredClone(pivot.dataSource),
    layout: structuredClone(pivot.layout),
    slicers: structuredClone(pivot.slicers),
    timelines: structuredClone(pivot.timelines),
    chartReferences: structuredClone(pivot.chartReferences),
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
const isPivotDataSource = (value: unknown): value is PivotDataSource => {
  if (!isRecord(value)) return false;
  if (value.kind === 'worksheet-range') return isRange(value.range);
  return value.kind === 'worksheet-ranges' && Array.isArray(value.ranges) && value.ranges.every(isRange)
    && Array.isArray(value.relationships);
};
const isPivotModel = (value: unknown): value is PivotModel => isRecord(value)
  && isNonEmptyString(value.id) && isNonEmptyString(value.sheetId)
  && isRange(value.sourceRange) && isPivotLayout(value.layout)
  && (value.dataSource === undefined || isPivotDataSource(value.dataSource));
const isPivotUpdate = (value: unknown): value is PivotUpdateParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && (value.sourceRange === undefined || isRange(value.sourceRange))
  && (value.dataSource === undefined || isPivotDataSource(value.dataSource))
  && (value.layout === undefined || isPivotLayout(value.layout))
  && (value.slicers === undefined || Array.isArray(value.slicers))
  && (value.timelines === undefined || Array.isArray(value.timelines))
  && (value.chartReferences === undefined || Array.isArray(value.chartReferences))
  && ['sourceRange', 'dataSource', 'layout', 'slicers', 'timelines', 'chartReferences'].some((key) => value[key] !== undefined);
const isPivotRefresh = (value: unknown): value is PivotRefreshParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && Number.isSafeInteger(value.refreshRevision) && Number(value.refreshRevision) >= 0
  && typeof value.lastRefreshedAt === 'string';
const isPivotDrillDown = (value: unknown): value is PivotDrillDownParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.pivotId)
  && typeof value.label === 'string' && isNonEmptyString(value.targetSheetId)
  && Array.isArray(value.sourceRowPaths)
  && value.sourceRowPaths.every((entry) => isRecord(entry) && isNonEmptyString(entry.sheetId) && Number.isSafeInteger(entry.row) && Number(entry.row) >= 0)
  && isRecord(value.targetAnchor)
  && Number.isSafeInteger(value.targetAnchor.row) && Number.isSafeInteger(value.targetAnchor.column)
  && Number(value.targetAnchor.row) >= 0 && Number(value.targetAnchor.column) >= 0;
const isPivotDrillDownRemove = (value: unknown): value is PivotDrillDownRemoveParams => isRecord(value) && isNonEmptyString(value.targetSheetId);

function pivotMutationRanges(value: unknown): RangeRef[] {
  if (isPivotModel(value)) return pivotSourceRanges(value);
  if (isPivotUpdate(value)) {
    if (value.dataSource?.kind === 'worksheet-ranges') return structuredClone(value.dataSource.ranges);
    if (value.dataSource?.kind === 'worksheet-range') return [structuredClone(value.dataSource.range)];
    if (value.sourceRange) return [structuredClone(value.sourceRange)];
    return sheetRange(value.sheetId);
  }
  if (isRecord(value) && isNonEmptyString(value.sheetId)) return sheetRange(value.sheetId);
  return [];
}

function applyPivotAdd(context: CommandContext, params: PivotModel): void {
  const sheet = context.workbook.getSheet(params.sheetId);
  if (context.workbook.getSheets().some((candidate) => candidate.pivots.some((entry) => entry.id === params.id))) throw new Error(`Pivot already exists: ${params.id}`);
  assertPivotDefinition(context.workbook, params);
  sheet.pivots.push(structuredClone(params));
}

function applyPivotRemove(context: CommandContext, params: string, sheetId: string): void {
  if (!removeById(context.workbook.getSheet(sheetId).pivots, params)) throw new Error(`Unknown pivot: ${params}`);
}

function applyPivotRefresh(context: CommandContext, params: PivotRefreshParams): void {
  const pivot = pivotFor(context, params.sheetId, params.pivotId);
  if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
  pivot.refreshRevision = params.refreshRevision;
  pivot.lastRefreshedAt = params.lastRefreshedAt;
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
  runtime.registry.registerMutation<PivotRefreshParams>('pivot.refresh', (item, context) => {
    applyPivotRefresh(context, item.params);
  }, {
    schema: { name: 'PivotRefreshParams', validate: isPivotRefresh },
    permission: { capability: 'pivot.refresh' },
    affectedRanges: { resolve: pivotMutationRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['pivot.refresh'], minCount: 1, maxCount: 1 },
  });

  runtime.registry.registerCommand<PivotModel>({
    id: 'pivot.add',
    execute: (params, context) => {
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'pivot.add',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: structuredClone(params),
        affectedRanges,
        inverse: [{ id: 'pivot.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: params.id, affectedRanges }],
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

  runtime.registry.registerCommand<{ sheetId: string; pivotId: string }>({
    id: 'pivot.refresh',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      const next: PivotRefreshParams = {
        pivotId: params.pivotId,
        sheetId: params.sheetId,
        refreshRevision: (pivot.refreshRevision ?? 0) + 1,
        lastRefreshedAt: new Date().toISOString(),
      };
      const previous: PivotRefreshParams = {
        pivotId: params.pivotId,
        sheetId: params.sheetId,
        refreshRevision: pivot.refreshRevision ?? 0,
        lastRefreshedAt: pivot.lastRefreshedAt ?? '',
      };
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'pivot.refresh',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: next,
        affectedRanges,
        inverse: [{ id: 'pivot.refresh', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }],
        apply: () => applyPivotRefresh(context, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
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

  const registerLayoutPatch = <P extends { sheetId: string; pivotId: string; field: string }>(
    commandId: string,
    buildLayout: (layout: PivotLayout, params: P) => PivotLayout,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const pivot = pivotFor(context, params.sheetId, params.pivotId);
        if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
        assertPivotField(context.workbook, pivot, params.field);
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

  registerLayoutPatch<PivotAggregateParams>('pivot.setAggregate', (layout, params) => setPivotAggregate(layout, params.field, params.summarizeBy));
  registerLayoutPatch<PivotShowAsParams>('pivot.setShowAs', (layout, params) => setPivotShowAs(layout, params.field, params.showAs));
  registerLayoutPatch<PivotGroupParams>('pivot.setGroup', (layout, params) => setPivotGroup(layout, params.axis, params.field, params.group));

  runtime.registry.registerCommand<PivotSlicerParams>({
    id: 'pivot.slicer.set',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      assertPivotField(context.workbook, pivot, params.slicer.field);
      const previous = structuredClone(pivot.slicers ?? []);
      const next = upsertPivotSlicer(pivot, params.slicer);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'pivot.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { pivotId: params.pivotId, sheetId: params.sheetId, slicers: next },
        affectedRanges,
        inverse: [{ id: 'pivot.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { pivotId: params.pivotId, sheetId: params.sheetId, slicers: previous }, affectedRanges }],
        apply: () => applyPivotUpdate(context, { pivotId: params.pivotId, sheetId: params.sheetId, slicers: next }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.slicer.set');

  runtime.registry.registerCommand<PivotTimelineParams>({
    id: 'pivot.timeline.set',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      assertPivotField(context.workbook, pivot, params.timeline.field);
      const previous = structuredClone(pivot.timelines ?? []);
      const next = upsertPivotTimeline(pivot, params.timeline);
      const affectedRanges = sheetRange(params.sheetId);
      context.applyMutation({
        id: 'pivot.update',
        unitId: context.workbook.unitId,
        sheetId: params.sheetId,
        params: { pivotId: params.pivotId, sheetId: params.sheetId, timelines: next },
        affectedRanges,
        inverse: [{ id: 'pivot.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { pivotId: params.pivotId, sheetId: params.sheetId, timelines: previous }, affectedRanges }],
        apply: () => applyPivotUpdate(context, { pivotId: params.pivotId, sheetId: params.sheetId, timelines: next }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.timeline.set');

  runtime.registry.registerCommand<PivotDrillDownParams>({
    id: 'pivot.drillDown',
    execute: (params, context) => {
      const pivot = pivotFor(context, params.sheetId, params.pivotId);
      if (!pivot) throw new Error(`Unknown pivot: ${params.pivotId}`);
      if (context.workbook.sheets.has(params.targetSheetId)) throw new Error(`Drill-down target already exists: ${params.targetSheetId}`);
      const columns = drillDownColumns(context, pivot).length;
      const sourceRangeCount = pivot.dataSource?.kind === 'worksheet-ranges' ? Math.max(pivot.dataSource.ranges.length, 1) : 1;
      const detailRows = Math.ceil(params.sourceRowPaths.length / sourceRangeCount);
      const affectedRanges: RangeRef[] = [
        { sheetId: params.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        { sheetId: params.targetSheetId, startRow: params.targetAnchor.row, endRow: params.targetAnchor.row + detailRows, startColumn: params.targetAnchor.column, endColumn: params.targetAnchor.column + Math.max(columns - 1, 0) },
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

export const PIVOT_MUTATION_IDS = ['pivot.add', 'pivot.remove', 'pivot.update', 'pivot.refresh', 'pivot.drilldown.add', 'pivot.drilldown.remove'] as const;
