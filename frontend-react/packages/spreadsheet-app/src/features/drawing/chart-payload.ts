import type { DrawingObject } from '@react-sheets/core-model';
import type { ChartInsertParams, ChartPayload } from '../chart/commands';

export function buildChartInsertParams(drawing: DrawingObject, payload: ChartPayload): ChartInsertParams {
  if (drawing.kind !== 'chart') throw new Error(`Drawing is not a chart: ${drawing.id}`);
  if (drawing.payloadId !== payload.chartId) throw new Error(`Chart payload identity mismatch: ${drawing.payloadId}`);
  return {
    sheetId: drawing.sheetId,
    drawing: structuredClone(drawing),
    payload: structuredClone(payload),
  };
}
