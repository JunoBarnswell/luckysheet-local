import type { RangeRef, SparklineGroup, SparklineModel, WorkbookModel, WorksheetModel } from '@react-sheets/core-model';

export interface SparklineInsertLocationParams {
  sheetId: string;
  sparklineId: string;
  dataRange: RangeRef;
  location: { row: number; column: number };
  type: SparklineModel['type'];
  groupId?: string;
  color?: string;
  negativeColor?: string;
  highlightMax?: boolean;
  highlightMin?: boolean;
  highlightFirst?: boolean;
  highlightLast?: boolean;
  highlightNegative?: boolean;
}

export function buildSparklineInsertParams(sparkline: SparklineModel): { sheetId: string; sparkline: SparklineModel } {
  return {
    sheetId: sparkline.sheetId,
    sparkline: structuredClone(sparkline),
  };
}

export function buildSparklineDataLocationParams(
  sheetId: string,
  sparklineId: string,
  dataRange: RangeRef,
  location: { row: number; column: number },
  type: SparklineModel['type'] = 'line',
  options?: Partial<Pick<SparklineModel, 'color' | 'negativeColor' | 'highlightMax' | 'highlightMin' | 'highlightFirst' | 'highlightLast' | 'highlightNegative' | 'groupId'>>,
): SparklineInsertLocationParams {
  return {
    sheetId,
    sparklineId,
    // The destination sheet owns the sparkline, while the source range may
    // intentionally live on another worksheet.  Do not rewrite the source
    // sheet id during command construction.
    dataRange: { ...dataRange },
    location,
    type,
    color: options?.color,
    negativeColor: options?.negativeColor,
    highlightMax: options?.highlightMax,
    highlightMin: options?.highlightMin,
    highlightFirst: options?.highlightFirst,
    highlightLast: options?.highlightLast,
    highlightNegative: options?.highlightNegative,
    groupId: options?.groupId,
  };
}

export function resolveQuickSparklinePlacement(range: RangeRef): { dataRange: RangeRef; location: { row: number; column: number } } {
  const singleRow = range.startRow === range.endRow;
  const dataRange: RangeRef = {
    sheetId: range.sheetId,
    startRow: singleRow ? range.startRow : range.startRow,
    endRow: singleRow ? range.startRow : range.endRow,
    startColumn: range.startColumn,
    endColumn: range.endColumn,
  };
  return {
    dataRange,
    location: {
      row: dataRange.startRow,
      column: dataRange.endColumn + 1,
    },
  };
}

export function extractSparklineValues(workbookOrSheet: WorkbookModel | WorksheetModel, sparkline: SparklineModel): number[] {
  const values: number[] = [];
  const source = sparkline.sourceRange;
  const sheet = 'getSheet' in workbookOrSheet
    ? workbookOrSheet.getSheet(source.sheetId)
    : workbookOrSheet.id === source.sheetId
      ? workbookOrSheet
      : undefined;
  if (!sheet) throw new Error(`Unknown sparkline source sheet: ${source.sheetId}`);
  for (let row = source.startRow; row <= source.endRow; row++) {
    for (let column = source.startColumn; column <= source.endColumn; column++) {
      const cell = sheet.cells.get(row, column);
      if (!cell) continue;
      const raw = cell.formulaValue ?? cell.value;
      if (raw == null) continue;
      const numeric = Number(String(raw).replace(/[$,%]/g, ''));
      if (Number.isFinite(numeric)) values.push(numeric);
    }
  }
  return values;
}

export function buildSparklineGroup(
  sheetId: string,
  groupId: string,
  sparklineIds: string[],
  type: SparklineModel['type'] = 'line',
  patch?: Partial<Pick<SparklineGroup, 'showAxis' | 'showMarkers'>>,
): SparklineGroup {
  return {
    id: groupId,
    sheetId,
    type,
    sparklineIds: [...sparklineIds],
    showAxis: patch?.showAxis ?? false,
    showMarkers: patch?.showMarkers ?? false,
  };
}
