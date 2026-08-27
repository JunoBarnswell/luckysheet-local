import type { PivotDefinition, PivotLayout, RangeRef, WorkbookModel } from '@react-sheets/core-model';
import { DEFAULT_PIVOT_COLLATION } from '@react-sheets/core-model';
import { getPivotFieldCatalog } from './engine';
import { buildDefaultPivotLayout } from './helpers';

export interface PivotTableRecommendation {
  id: string;
  title: string;
  confidence: number;
  source: Extract<PivotDefinition['source'], { kind: 'worksheet-range' }>;
  layout: PivotLayout;
  summary: string;
}

export function recommendPivotTables(workbook: WorkbookModel, sheetId: string, sourceRange: RangeRef): readonly PivotTableRecommendation[] {
  if (sourceRange.endRow <= sourceRange.startRow || sourceRange.endColumn < sourceRange.startColumn) {
    throw new Error('PIVOT_SOURCE_INVALID: Recommended PivotTables requires headers and at least one data row');
  }
  const defaultLayout = buildDefaultPivotLayout(workbook, sheetId, sourceRange);
  if (!defaultLayout) throw new Error('PIVOT_SOURCE_INVALID: Recommended PivotTables could not resolve a value field');
  const draft: PivotDefinition = {
    schema: 'PivotDefinition',
    id: 'recommended-pivot-analysis',
    source: { kind: 'worksheet-range', range: structuredClone(sourceRange) },
    target: { sheetId, anchor: { row: 0, column: 0 } },
    fieldCatalog: { fields: [] },
    refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
    layout: defaultLayout,
  };
  const catalog = getPivotFieldCatalog(workbook, draft);
  const dimensions = catalog.fields.filter((field) => field.dataType !== 'number');
  const measures = catalog.fields.filter((field) => field.dataType === 'number');
  const primaryMeasure = measures[0] ?? catalog.fields[0];
  if (!primaryMeasure) throw new Error('PIVOT_SOURCE_INVALID: Recommended PivotTables found no fields');
  const source = { kind: 'worksheet-range' as const, range: structuredClone(sourceRange) };
  const base = {
    filters: [],
    allowMultipleFiltersPerField: true,
    collation: { ...DEFAULT_PIVOT_COLLATION },
    subtotalLocation: 'bottom' as const,
    showRowGrandTotals: true,
    showColumnGrandTotals: true,
    reportLayout: 'compact' as const,
    calculatedFields: [],
    calculatedItems: [],
    expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
  };
  const value = { valueId: `value:${primaryMeasure.fieldId}`, fieldId: primaryMeasure.fieldId, summarizeBy: primaryMeasure.dataType === 'number' ? 'sum' as const : 'count' as const };
  const candidates: PivotTableRecommendation[] = [];
  if (dimensions[0]) candidates.push({
    id: 'recommended-pivot-rows',
    title: `${primaryMeasure.name} by ${dimensions[0].name}`,
    confidence: 0.98,
    source,
    summary: `${dimensions[0].name} → Rows; ${primaryMeasure.name} → Values`,
    layout: { ...base, rows: [{ fieldId: dimensions[0].fieldId }], columns: [], values: [value] },
  });
  if (dimensions[0] && dimensions[1]) candidates.push({
    id: 'recommended-pivot-matrix',
    title: `${primaryMeasure.name}: ${dimensions[0].name} × ${dimensions[1].name}`,
    confidence: 0.92,
    source,
    summary: `${dimensions[0].name} → Rows; ${dimensions[1].name} → Columns; ${primaryMeasure.name} → Values`,
    layout: { ...base, rows: [{ fieldId: dimensions[0].fieldId }], columns: [{ fieldId: dimensions[1].fieldId }], values: [value] },
  });
  if (measures.length > 1 && dimensions[0]) candidates.push({
    id: 'recommended-pivot-values',
    title: `Measures by ${dimensions[0].name}`,
    confidence: 0.86,
    source,
    summary: `${dimensions[0].name} → Rows; ${measures.slice(0, 3).map((field) => field.name).join(', ')} → Values`,
    layout: { ...base, rows: [{ fieldId: dimensions[0].fieldId }], columns: [], values: measures.slice(0, 3).map((field) => ({ valueId: `value:${field.fieldId}`, fieldId: field.fieldId, summarizeBy: 'sum' as const })) },
  });
  if (!candidates.length) candidates.push({
    id: 'recommended-pivot-summary',
    title: `${primaryMeasure.name} summary`,
    confidence: 0.75,
    source,
    summary: `${primaryMeasure.name} → Values`,
    layout: { ...base, rows: [], columns: [], values: [value] },
  });
  return candidates;
}
