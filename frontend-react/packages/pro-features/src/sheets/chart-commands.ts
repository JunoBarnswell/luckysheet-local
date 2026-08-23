import type { ChartDrawingPayload, RangeRef } from '@react-sheets/core-model';

/** Public chart command payload types are kept here for exchange/UI callers. */
export type P1ChartType = ChartDrawingPayload['chartType'];

export interface ChartInsertParams {
  sheetId: string;
  chartId: string;
  drawingId: string;
  bounds: { x: number; y: number; width: number; height: number };
  payload: ChartDrawingPayload;
  zIndex?: number;
  rotation?: number;
}

export interface ChartUpdateParams {
  sheetId: string;
  chartId: string;
  payload: Partial<ChartDrawingPayload>;
}

export interface ChartSetTypeParams {
  sheetId: string;
  chartId: string;
  chartType: P1ChartType;
  stacked?: ChartDrawingPayload['stacked'];
}

export interface ChartSetLegendParams {
  sheetId: string;
  chartId: string;
  legendPosition: NonNullable<ChartDrawingPayload['legendPosition']>;
}

export interface ChartSetDataLabelsParams {
  sheetId: string;
  chartId: string;
  showDataLabels: boolean;
}

export interface ChartSetSeriesParams {
  sheetId: string;
  chartId: string;
  sourceRanges: RangeRef[];
  series?: ChartDrawingPayload['series'];
  categoryRange?: RangeRef;
}
