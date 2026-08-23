import type { CommandRuntime } from '@react-sheets/command-runtime';
import type {
  PivotAggregateFunction,
  PivotGroup,
  PivotLayout,
  PivotShowAs,
  PivotSlicer,
  PivotTimeline,
  RangeRef,
} from '@react-sheets/core-model';
import { applyTrackedMutation, registerMutationHandler, sheetRange } from './command-helpers';
import { createPivotDrillDownSheetName, setPivotAggregate, setPivotGroup, setPivotShowAs, upsertPivotSlicer, upsertPivotTimeline } from './pivot-panel-state';

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

function updatePivotLayout(
  runtime: CommandRuntime,
  context: import('@react-sheets/command-runtime').CommandContext,
  params: PivotLayoutCommandParams,
  previousLayout: PivotLayout,
): void {
  const affectedRanges = sheetRange(params.sheetId);
  applyTrackedMutation(context, {
    id: 'pivot.update',
    sheetId: params.sheetId,
    params: { pivotId: params.pivotId, sheetId: params.sheetId, layout: params.layout },
    inverseParams: { pivotId: params.pivotId, sheetId: params.sheetId, layout: previousLayout },
    affectedRanges,
    apply: () => {
      const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
      if (pivot) pivot.layout = structuredClone(params.layout);
    },
  });
}

export function registerExtendedPivotCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<PivotDrillDownParams>(runtime, 'pivot.drilldown', (params, context) => {
    const sourceSheet = context.workbook.getSheet(params.sheetId);
    const pivot = sourceSheet.pivots.find((entry) => entry.id === params.pivotId);
    if (!pivot) return;
    if (!context.workbook.sheets.has(params.targetSheetId)) {
      context.workbook.addSheet(params.targetSheetId, createPivotDrillDownSheetName(pivot, params.label));
    }
    const target = context.workbook.getSheet(params.targetSheetId);
    const header = pivot.layout.values.map((field) => field.displayName ?? field.field);
    header.forEach((name, index) => target.cells.set(params.targetAnchor.row, params.targetAnchor.column + index, { value: name }));
    params.sourceRowPaths.forEach((path, rowOffset) => {
      const source = context.workbook.getSheet(path.sheetId);
      pivot.layout.values.forEach((field, columnOffset) => {
        const value = source.cells.get(path.row, pivot.sourceRange.startColumn + columnOffset)?.value ?? null;
        target.cells.set(params.targetAnchor.row + rowOffset + 1, params.targetAnchor.column + columnOffset, { value });
      });
    });
  });

  const registerLayoutPatch = <P extends { sheetId: string; pivotId: string }>(
    commandId: string,
    buildLayout: (layout: PivotLayout, params: P) => PivotLayout,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
        if (!pivot) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
        const previousLayout = structuredClone(pivot.layout);
        const nextLayout = buildLayout(previousLayout, params);
        updatePivotLayout(runtime, context, { sheetId: params.sheetId, pivotId: params.pivotId, layout: nextLayout }, previousLayout);
        return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
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
      const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
      if (!pivot) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const previous = structuredClone(pivot.slicers ?? []);
      const next = upsertPivotSlicer(pivot, params.slicer);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation(context, {
        id: 'pivot.update',
        sheetId: params.sheetId,
        params: { pivotId: params.pivotId, sheetId: params.sheetId, slicers: next },
        inverseParams: { pivotId: params.pivotId, sheetId: params.sheetId, slicers: previous },
        affectedRanges,
        apply: () => {
          const current = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
          if (current) current.slicers = structuredClone(next);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.slicer.set');

  runtime.registry.registerCommand<PivotTimelineParams>({
    id: 'pivot.timeline.set',
    execute: (params, context) => {
      const pivot = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
      if (!pivot) return { operationId: context.operationId, mutationCount: 0, affectedRanges: sheetRange(params.sheetId) };
      const previous = structuredClone(pivot.timelines ?? []);
      const next = upsertPivotTimeline(pivot, params.timeline);
      const affectedRanges = sheetRange(params.sheetId);
      applyTrackedMutation(context, {
        id: 'pivot.update',
        sheetId: params.sheetId,
        params: { pivotId: params.pivotId, sheetId: params.sheetId, timelines: next },
        inverseParams: { pivotId: params.pivotId, sheetId: params.sheetId, timelines: previous },
        affectedRanges,
        apply: () => {
          const current = context.workbook.getSheet(params.sheetId).pivots.find((entry) => entry.id === params.pivotId);
          if (current) current.timelines = structuredClone(next);
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.timeline.set');

  runtime.registry.registerCommand<PivotDrillDownParams>({
    id: 'pivot.drillDown',
    execute: (params, context) => {
      const affectedRanges: RangeRef[] = [
        { sheetId: params.sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
        { sheetId: params.targetSheetId, startRow: params.targetAnchor.row, endRow: params.targetAnchor.row + params.sourceRowPaths.length, startColumn: params.targetAnchor.column, endColumn: params.targetAnchor.column + 3 },
      ];
      applyTrackedMutation(context, {
        id: 'pivot.drilldown',
        sheetId: params.sheetId,
        params,
        inverseParams: params,
        affectedRanges,
        apply: () => runtime.registry.getMutation('pivot.drilldown')({ id: 'pivot.drilldown', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges }, context),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('pivot.drillDown');

  return commandIds;
}
