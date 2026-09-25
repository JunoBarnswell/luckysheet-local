import type { ChartDrawingPayload, ChartTextFormulaField, ChartTextModel } from './domain';

const CHART_TEXT_FORMULA_FIELDS: readonly ChartTextFormulaField[] = [
  'titleText.linkedFormula',
  'legend.text.linkedFormula',
  'categoryAxis.titleText.linkedFormula',
  'valueAxis.titleText.linkedFormula',
  'secondaryCategoryAxis.titleText.linkedFormula',
  'secondaryValueAxis.titleText.linkedFormula',
  'dataTable.font.linkedFormula',
];

function chartTextModel(payload: ChartDrawingPayload, field: ChartTextFormulaField): ChartTextModel | undefined {
  const elements: unknown = payload.elements;
  if (!elements || typeof elements !== 'object') {
    throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: chart ${payload.chartId} has no canonical elements`);
  }
  const path = field.slice(0, -'.linkedFormula'.length).split('.');
  let current: unknown = elements;
  for (const segment of path) {
    if (current === undefined) return undefined;
    if (current === null || typeof current !== 'object' || Array.isArray(current)) {
      throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: chart ${payload.chartId} ${field} owner path is not an object`);
    }
    current = (current as Record<string, unknown>)[segment];
  }
  if (current === undefined) return undefined;
  if (current === null || typeof current !== 'object' || Array.isArray(current)) {
    throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: chart ${payload.chartId} ${field} owner is not an object`);
  }
  return current as ChartTextModel;
}

export function chartTextFormulaEntries(
  payload: ChartDrawingPayload,
): Array<{ field: ChartTextFormulaField; formula: string }> {
  const entries: Array<{ field: ChartTextFormulaField; formula: string }> = [];
  for (const field of CHART_TEXT_FORMULA_FIELDS) {
    const formula = chartTextModel(payload, field)?.linkedFormula;
    if (formula === undefined) continue;
    if (typeof formula !== 'string' || formula.trim().length === 0) {
      throw new Error(`STRUCTURAL_REFERENCE_OWNER_INVALID: chart ${payload.chartId} ${field} is not a non-empty formula`);
    }
    entries.push({ field, formula });
  }
  return entries;
}

export function readChartTextFormula(
  payload: ChartDrawingPayload,
  field: ChartTextFormulaField,
): string | undefined {
  return chartTextModel(payload, field)?.linkedFormula;
}

export function writeChartTextFormula(
  payload: ChartDrawingPayload,
  field: ChartTextFormulaField,
  formula: string,
): void {
  const model = chartTextModel(payload, field);
  if (!model) throw new Error(`STRUCTURAL_PATCH_INVARIANT: chart ${payload.chartId} ${field} owner disappeared`);
  model.linkedFormula = formula;
}
