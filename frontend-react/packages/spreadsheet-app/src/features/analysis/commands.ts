import type {
  AnalysisViewDefinition,
  AnalysisViewLayout,
  DataViewField,
  TableScalar,
} from '@react-sheets/core-model';
import type { CommandContext, CommandRuntime } from '@react-sheets/command-runtime';

export interface AnalysisViewReplaceParams {
  view: AnalysisViewDefinition | null;
  viewId?: string;
}

export interface AnalysisViewSetParams {
  view: AnalysisViewDefinition;
}

export interface AnalysisViewRemoveParams {
  viewId: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isScalar(value: unknown): value is TableScalar {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
}

function isField(value: unknown): value is DataViewField {
  return isRecord(value)
    && typeof value.fieldId === 'string'
    && value.fieldId.trim().length > 0
    && typeof value.caption === 'string'
    && value.caption.trim().length > 0
    && (value.formula === undefined || typeof value.formula === 'string')
    && (value.widthPx === undefined || (typeof value.widthPx === 'number' && Number.isFinite(value.widthPx) && value.widthPx > 0));
}

function isLayout(value: unknown): value is AnalysisViewLayout {
  return isRecord(value)
    && Number.isSafeInteger(value.columns) && Number(value.columns) >= 1 && Number(value.columns) <= 12
    && typeof value.rowHeightPx === 'number' && Number.isFinite(value.rowHeightPx) && value.rowHeightPx >= 1
    && typeof value.gapPx === 'number' && Number.isFinite(value.gapPx) && value.gapPx >= 0;
}

function isAnalysisView(value: unknown): value is AnalysisViewDefinition {
  if (!isRecord(value)
    || value.kind !== 'analysis'
    || typeof value.id !== 'string' || value.id.trim().length === 0
    || typeof value.name !== 'string' || value.name.trim().length === 0
    || typeof value.tableId !== 'string' || value.tableId.trim().length === 0
    || !Array.isArray(value.fields) || !value.fields.every(isField)
    || !Array.isArray(value.filters) || !Array.isArray(value.charts)
    || !isLayout(value.layout)
    || !Number.isSafeInteger(value.revision) || Number(value.revision) < 0) return false;
  for (const filter of value.filters) {
    if (!isRecord(filter)
      || typeof filter.id !== 'string' || filter.id.trim().length === 0
      || typeof filter.fieldId !== 'string' || filter.fieldId.trim().length === 0
      || !['equals', 'not-equals', 'contains', 'in', 'between'].includes(String(filter.operator))
      || !Array.isArray(filter.values) || !filter.values.every(isScalar)) return false;
  }
  for (const chart of value.charts) {
    if (!isRecord(chart) || typeof chart.chartId !== 'string' || chart.chartId.trim().length === 0 || !isRecord(chart.fieldMap)) return false;
    for (const fieldId of Object.values(chart.fieldMap)) if (typeof fieldId !== 'string' || fieldId.trim().length === 0) return false;
  }
  return true;
}

function isAnalysisViewReplace(value: unknown): value is AnalysisViewReplaceParams {
  if (!isRecord(value) || !Object.prototype.hasOwnProperty.call(value, 'view')) return false;
  if (value.view === null) return typeof value.viewId === 'string' && value.viewId.trim().length > 0;
  return isAnalysisView(value.view);
}

function validateAnalysisView(context: CommandContext, value: AnalysisViewDefinition): AnalysisViewDefinition {
  if (!isAnalysisView(value)) throw new Error('Analysis view payload is invalid');
  const table = context.workbook.dataModel.tables.get(value.tableId);
  if (!table) throw new Error(`Analysis view table not found: ${value.tableId}`);
  const fields = new Set(table.fields.map((field) => field.id));
  const viewFields = new Set<string>();
  for (const field of value.fields) {
    if (!fields.has(field.fieldId)) throw new Error(`Analysis view field not found: ${field.fieldId}`);
    if (!viewFields.add(field.fieldId)) throw new Error(`Analysis view field is duplicated: ${field.fieldId}`);
  }
  for (const filter of value.filters) if (!fields.has(filter.fieldId)) throw new Error(`Analysis filter field not found: ${filter.fieldId}`);
  for (const chart of value.charts) {
    for (const fieldId of Object.values(chart.fieldMap)) if (fieldId !== undefined && !fields.has(fieldId)) throw new Error(`Analysis chart field not found: ${fieldId}`);
  }
  return structuredClone(value);
}

function applyAnalysisView(context: CommandContext, params: AnalysisViewReplaceParams): void {
  if (params.view === null) {
    context.workbook.removeAnalysisView(params.viewId!);
    return;
  }
  context.workbook.setAnalysisView(params.view);
}

export function registerAnalysisCommands(runtime: CommandRuntime): string[] {
  runtime.registry.registerMutation<AnalysisViewReplaceParams>({
    id: 'analysis.view.replace',
    handler: (item, context) => {
      if (!isAnalysisViewReplace(item.params)) throw new Error('analysis.view.replace requires a canonical view or viewId');
      if (item.params.view) applyAnalysisView(context, { view: validateAnalysisView(context, item.params.view) });
      else applyAnalysisView(context, item.params);
    },
    metadata: {
      schema: { name: 'AnalysisViewReplaceParams', validate: isAnalysisViewReplace },
      permission: { capability: 'analysis.view.write', roles: ['owner', 'editor'] },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inversePolicy: { allowedMutationIds: ['analysis.view.replace'], minCount: 1, maxCount: 1 },
    },
  });

  runtime.registry.registerCommand<AnalysisViewSetParams>({
    id: 'analysis.view.set',
    execute: (params, context) => {
      const view = validateAnalysisView(context, params.view);
      const previous = context.workbook.getAnalysisView(view.id);
      context.applyMutation({
        id: 'analysis.view.replace',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params: { view },
        affectedRanges: [],
        inverse: [{
          id: 'analysis.view.replace',
          unitId: context.workbook.unitId,
          sheetId: context.workbook.primarySheetId,
          params: previous ? { view: previous } : { view: null, viewId: view.id },
          affectedRanges: [],
        }],
        apply: () => applyAnalysisView(context, { view }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  runtime.registry.registerCommand<AnalysisViewRemoveParams>({
    id: 'analysis.view.remove',
    execute: (params, context) => {
      const previous = context.workbook.getAnalysisView(params.viewId);
      if (!previous) throw new Error(`Analysis view not found: ${params.viewId}`);
      context.applyMutation({
        id: 'analysis.view.replace',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.primarySheetId,
        params: { view: null, viewId: params.viewId },
        affectedRanges: [],
        inverse: [{ id: 'analysis.view.replace', unitId: context.workbook.unitId, sheetId: context.workbook.primarySheetId, params: { view: previous }, affectedRanges: [] }],
        apply: () => applyAnalysisView(context, { view: null, viewId: params.viewId }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  return ['analysis.view.set', 'analysis.view.remove'];
}
