import type {
  ChartModel,
  DrawingObject,
  DrawingTransform,
  FloatingImage,
  ShapeModel,
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

export function buildShapeDrawingAdd(sheetId: string, shape: ShapeModel, drawingId: string): DrawingAddParams {
  return {
    sheetId,
    drawing: {
      id: drawingId,
      sheetId,
      kind: 'shape',
      anchor: { kind: 'absolute' },
      transform: boundsToTransform(shape.bounds, shape.rotation ?? 0),
      zIndex: 0,
      payloadId: shape.id,
    },
    payload: {
      kind: 'shape',
      type: shape.type,
      fill: shape.fill,
      stroke: shape.stroke,
      strokeWidth: shape.strokeWidth,
      text: shape.text,
      textColor: shape.textColor,
      fontSize: shape.fontSize,
    },
  };
}

export function buildImageDrawingAdd(sheetId: string, image: FloatingImage, drawingId: string): DrawingAddParams {
  return {
    sheetId,
    drawing: {
      id: drawingId,
      sheetId,
      kind: 'image',
      anchor: { kind: 'absolute' },
      transform: boundsToTransform(image.bounds),
      zIndex: 0,
      payloadId: image.id,
    },
    payload: {
      kind: 'image',
      src: image.src,
      name: image.name,
    },
  };
}

export function resolveDrawingMoveTransform(
  sheet: WorksheetModel,
  payloadId: string,
  bounds: ChartModel['bounds'],
): { drawingId: string; transform: DrawingTransform } | undefined {
  const drawing = findDrawingByPayloadId(sheet, payloadId);
  if (!drawing) return undefined;
  return {
    drawingId: drawing.id,
    transform: boundsToTransform(bounds, drawing.transform.rotation),
  };
}
