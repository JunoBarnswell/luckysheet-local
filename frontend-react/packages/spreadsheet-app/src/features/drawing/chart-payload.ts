import type { ChartDrawingPayload, DrawingObject } from '@react-sheets/core-model';
import type { ChartInsertParams } from '@react-sheets/pro-features';

export function buildChartInsertParams(drawing: DrawingObject, payload: ChartDrawingPayload): ChartInsertParams {
  if (drawing.kind !== 'chart') throw new Error(`Drawing is not a chart: ${drawing.id}`);
  if (drawing.payloadId !== payload.chartId) throw new Error(`Chart payload identity mismatch: ${drawing.payloadId}`);
  return {
    sheetId: drawing.sheetId,
    chartId: payload.chartId,
    drawingId: drawing.id,
    bounds: {
      x: drawing.transform.x,
      y: drawing.transform.y,
      width: drawing.transform.width,
      height: drawing.transform.height,
    },
    payload: structuredClone(payload),
    zIndex: drawing.zIndex,
    rotation: drawing.transform.rotation,
  };
}
