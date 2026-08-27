import { defaultChartSubtype, type ChartDrawingPayload, type RangeRef, type WorkbookModel } from '@react-sheets/core-model';

export interface ChartRecommendation {
  id: string;
  chartType: ChartDrawingPayload['chartType'];
  subtype: ChartDrawingPayload['subtype'];
  title: string;
  confidence: number;
  reason: 'time-series' | 'category-comparison' | 'part-to-whole' | 'correlation' | 'hierarchy';
  source: { kind: 'worksheet-ranges'; ranges: [RangeRef] };
}

interface ColumnProfile {
  index: number;
  header: string;
  numericCount: number;
  textCount: number;
  dateCount: number;
  nonBlankCount: number;
}

function isDateSemantic(value: unknown, numberFormat: string | undefined): boolean {
  if (typeof value === 'number') return /(?:^|[^a-z])[ymdhis]+(?:[^a-z]|$)/i.test(numberFormat ?? '');
  return typeof value === 'string' && /^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:[ T].*)?$/.test(value.trim());
}

function profileColumns(workbook: WorkbookModel, range: RangeRef): ColumnProfile[] {
  const sheet = workbook.getSheet(range.sheetId);
  const profiles: ColumnProfile[] = [];
  for (let column = range.startColumn; column <= range.endColumn; column += 1) {
    const headerCell = sheet.cells.get(range.startRow, column);
    const profile: ColumnProfile = {
      index: column,
      header: String(headerCell?.value ?? `Column ${column - range.startColumn + 1}`),
      numericCount: 0,
      textCount: 0,
      dateCount: 0,
      nonBlankCount: 0,
    };
    for (let row = range.startRow + 1; row <= range.endRow; row += 1) {
      const cell = sheet.cells.get(row, column);
      const value = cell?.value;
      if (value === null || value === undefined || value === '') continue;
      profile.nonBlankCount += 1;
      if (isDateSemantic(value, cell?.numberFormat)) profile.dateCount += 1;
      else if (typeof value === 'number' && Number.isFinite(value)) profile.numericCount += 1;
      else profile.textCount += 1;
    }
    profiles.push(profile);
  }
  return profiles;
}

function addCandidate(target: ChartRecommendation[], range: RangeRef, input: Omit<ChartRecommendation, 'source' | 'subtype'>): void {
  if (target.some((candidate) => candidate.chartType === input.chartType)) return;
  target.push({ ...input, subtype: defaultChartSubtype(input.chartType), source: { kind: 'worksheet-ranges', ranges: [structuredClone(range)] } });
}

/** Deterministic selection analyzer used by the Recommended Charts dialog. */
export function recommendCharts(workbook: WorkbookModel, range: RangeRef): readonly ChartRecommendation[] {
  if (range.endRow <= range.startRow || range.endColumn < range.startColumn) {
    throw new Error('INVALID_CHART_SOURCE: Recommended Charts requires a header row and at least one data row');
  }
  const profiles = profileColumns(workbook, range);
  const populated = profiles.filter((profile) => profile.nonBlankCount > 0);
  const numeric = populated.filter((profile) => profile.numericCount > 0 && profile.numericCount >= profile.textCount);
  const dates = populated.filter((profile) => profile.dateCount > 0 && profile.dateCount >= profile.numericCount);
  const categories = populated.filter((profile) => profile.textCount > 0 && profile.textCount >= profile.numericCount);
  if (numeric.length === 0) throw new Error('INVALID_CHART_SOURCE: Recommended Charts requires at least one numeric field');

  const rowCount = range.endRow - range.startRow;
  const recommendations: ChartRecommendation[] = [];
  if (dates.length > 0) addCandidate(recommendations, range, { id: 'recommended-line', chartType: 'line', title: `${numeric[0]!.header} trend`, confidence: 0.98, reason: 'time-series' });
  if (categories.length > 0) addCandidate(recommendations, range, { id: 'recommended-column', chartType: 'column', title: `${numeric[0]!.header} by ${categories[0]!.header}`, confidence: dates.length ? 0.86 : 0.96, reason: 'category-comparison' });
  if (numeric.length >= 2 && categories.length === 0 && dates.length === 0) addCandidate(recommendations, range, { id: 'recommended-scatter', chartType: 'scatter', title: `${numeric[1]!.header} by ${numeric[0]!.header}`, confidence: 0.94, reason: 'correlation' });
  if (numeric.length === 1 && categories.length > 0 && rowCount <= 12) addCandidate(recommendations, range, { id: 'recommended-pie', chartType: 'pie', title: `${numeric[0]!.header} share`, confidence: rowCount <= 7 ? 0.9 : 0.72, reason: 'part-to-whole' });
  if (categories.length >= 2 && numeric.length > 0) addCandidate(recommendations, range, { id: 'recommended-treemap', chartType: 'treemap', title: `${numeric[0]!.header} hierarchy`, confidence: 0.84, reason: 'hierarchy' });
  if (recommendations.length === 0) addCandidate(recommendations, range, { id: 'recommended-column', chartType: 'column', title: numeric[0]!.header, confidence: 0.75, reason: 'category-comparison' });
  return recommendations.slice().sort((left, right) => right.confidence - left.confidence);
}
