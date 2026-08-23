import type {
  DrawingObject,
  DrawingPayload,
  DrawingTransform,
  WorksheetModel,
} from '@react-sheets/core-model';
import type { DrawingAddParams } from './commands';

export function findDrawingByPayloadId(sheet: WorksheetModel, payloadId: string): DrawingObject | undefined {
  return sheet.drawings.find((entry) => entry.payloadId === payloadId || entry.id === payloadId);
}

export function boundsToTransform(
  bounds: { x: number; y: number; width: number; height: number },
  rotation = 0,
): DrawingTransform {
  return { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height, rotation };
}

/** Build the sole drawing aggregate insertion payload.
 *
 * Callers are required to provide the canonical persisted pair. The helper
 * validates their identity and makes a defensive copy; it does not translate
 * legacy per-kind models into drawing data.
 */
export function buildDrawingAdd(drawing: DrawingObject, payload: DrawingPayload): DrawingAddParams {
  if (drawing.sheetId.length === 0) throw new Error(`Drawing sheet is required: ${drawing.id}`);
  if (drawing.payloadId.length === 0) throw new Error(`Drawing payload is required: ${drawing.id}`);
  if (drawing.kind !== payload.kind) throw new Error(`Drawing payload kind mismatch: ${drawing.id}`);
  if (payload.kind === 'chart' && payload.chartId !== drawing.payloadId) {
    throw new Error(`Drawing payload identity mismatch: ${drawing.payloadId}`);
  }
  return {
    sheetId: drawing.sheetId,
    drawing: structuredClone(drawing),
    payload: structuredClone(payload),
  };
}

export function resolveDrawingMoveTransform(
  sheet: WorksheetModel,
  payloadId: string,
  bounds: Pick<DrawingTransform, 'x' | 'y' | 'width' | 'height'> & { rotation?: number },
): { drawingId: string; transform: DrawingTransform } | undefined {
  const drawing = findDrawingByPayloadId(sheet, payloadId);
  if (!drawing) return undefined;
  return {
    drawingId: drawing.id,
    transform: boundsToTransform(bounds, bounds.rotation ?? drawing.transform.rotation ?? 0),
  };
}
