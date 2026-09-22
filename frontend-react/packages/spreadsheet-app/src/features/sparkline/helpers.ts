import type { RangeRef, SparklineGroup, SparklineModel, WorkbookModel, WorksheetModel } from '@react-sheets/core-model';
import { resolveSparklineSeries, type ResolvedSparklineSeries, type StructuredChartSheet } from '../chart/data';

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
  showAxis?: boolean;
  showMarkers?: boolean;
  lineWeight?: number;
  dateAxis?: boolean;
  dataOrientation?: SparklineModel['dataOrientation'];
  rightToLeft?: boolean;
  hiddenCells?: SparklineModel['hiddenCells'];
  emptyCells?: SparklineModel['emptyCells'];
  verticalAxis?: SparklineModel['verticalAxis'];
  axisColor?: string;
  firstColor?: string;
  lastColor?: string;
  highColor?: string;
  lowColor?: string;
  markerColor?: string;
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
  options?: Partial<Pick<SparklineModel, 'color' | 'negativeColor' | 'highlightMax' | 'highlightMin' | 'highlightFirst' | 'highlightLast' | 'highlightNegative' | 'groupId' | 'showAxis' | 'showMarkers' | 'lineWeight' | 'dateAxis' | 'dataOrientation' | 'rightToLeft' | 'hiddenCells' | 'emptyCells' | 'verticalAxis' | 'axisColor' | 'firstColor' | 'lastColor' | 'highColor' | 'lowColor' | 'markerColor'>>,
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
    showAxis: options?.showAxis,
    showMarkers: options?.showMarkers,
    groupId: options?.groupId,
    lineWeight: options?.lineWeight,
    dateAxis: options?.dateAxis,
    dataOrientation: options?.dataOrientation,
    rightToLeft: options?.rightToLeft,
    hiddenCells: options?.hiddenCells,
    emptyCells: options?.emptyCells,
    verticalAxis: options?.verticalAxis ? structuredClone(options.verticalAxis) : undefined,
    axisColor: options?.axisColor,
    firstColor: options?.firstColor,
    lastColor: options?.lastColor,
    highColor: options?.highColor,
    lowColor: options?.lowColor,
    markerColor: options?.markerColor,
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
  const resolved = resolveSparklineSeries(sparkline, (sheetId) => {
    const sheet = 'getSheet' in workbookOrSheet
      ? workbookOrSheet.getSheet(sheetId)
      : workbookOrSheet.id === sheetId
        ? workbookOrSheet
        : undefined;
    return sheet ? { getCell: (row: number, column: number) => sheet.cells.get(row, column), hiddenRows: sheet.hiddenRows, hiddenColumns: sheet.hiddenColumns } : undefined;
  });
  return resolved.values.filter((value): value is number => value !== null);
}

export function buildSparklineGroup(
  sheetId: string,
  groupId: string,
  sparklineIds: string[],
  type: SparklineModel['type'] = 'line',
  patch?: Partial<Omit<SparklineGroup, 'id' | 'sheetId' | 'type' | 'sparklineIds'>>,
): SparklineGroup {
  return {
    id: groupId,
    sheetId,
    type,
    sparklineIds: [...sparklineIds],
    showAxis: patch?.showAxis ?? false,
    showMarkers: patch?.showMarkers ?? false,
    ...structuredClone(patch ?? {}),
  };
}

export function resolveSparklineData(
  sparkline: SparklineModel,
  getSheet: (sheetId: string) => StructuredChartSheet | undefined,
  group?: SparklineGroup,
): ResolvedSparklineSeries {
  return resolveSparklineSeries(sparkline, getSheet, group);
}
