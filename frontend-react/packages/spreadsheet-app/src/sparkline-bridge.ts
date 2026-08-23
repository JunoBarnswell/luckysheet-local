import type { RangeRef, SparklineGroup, SparklineModel, WorksheetModel } from '@react-sheets/core-model';

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
  options?: Partial<Pick<SparklineModel, 'color' | 'negativeColor' | 'highlightMax' | 'highlightMin' | 'groupId'>>,
): SparklineInsertLocationParams {
  return {
    sheetId,
    sparklineId,
    dataRange: { ...dataRange, sheetId },
    location,
    type,
    color: options?.color,
    negativeColor: options?.negativeColor,
    highlightMax: options?.highlightMax,
    highlightMin: options?.highlightMin,
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

export function extractSparklineValues(sheet: WorksheetModel, sparkline: SparklineModel): number[] {
  const values: number[] = [];
  const source = sparkline.sourceRange;
  for (let row = source.startRow; row <= source.endRow; row++) {
    for (let column = source.startColumn; column <= source.endColumn; column++) {
      const cell = sheet.cells.get(row, column);
      if (!cell || cell.value == null) continue;
      const numeric = Number(String(cell.value).replace(/[$,%]/g, ''));
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
