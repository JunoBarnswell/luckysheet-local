export type DimensionResizeAxis = 'column' | 'row';

export interface DimensionResizeGesture {
  axis: DimensionResizeAxis;
  boundaryIndex: number;
  startModelSizePx: number;
  startPointerScreenPx: number;
  currentPointerScreenPx: number;
  currentModelSizePx: number;
  zoomScale: number;
  minimumModelSizePx: number;
}

export interface BeginDimensionResizeGestureInput {
  axis: DimensionResizeAxis;
  boundaryIndex: number;
  startModelSizePx: number;
  startPointerScreenPx: number;
  zoomScale: number;
  minimumModelSizePx: number;
}

export function beginDimensionResizeGesture(input: BeginDimensionResizeGestureInput): DimensionResizeGesture {
  if (!Number.isSafeInteger(input.boundaryIndex) || input.boundaryIndex < 0) throw new Error('Dimension resize boundary index must be a non-negative integer');
  if (!Number.isFinite(input.startModelSizePx) || input.startModelSizePx <= 0) throw new Error('Dimension resize start size must be positive pixels');
  if (!Number.isFinite(input.startPointerScreenPx)) throw new Error('Dimension resize start pointer must be finite');
  if (!Number.isFinite(input.zoomScale) || input.zoomScale <= 0) throw new Error('Dimension resize zoom must be positive');
  if (!Number.isFinite(input.minimumModelSizePx) || input.minimumModelSizePx <= 0) throw new Error('Dimension resize minimum must be positive');
  const currentModelSizePx = Math.max(input.minimumModelSizePx, Math.round(input.startModelSizePx));
  return {
    ...input,
    currentPointerScreenPx: input.startPointerScreenPx,
    currentModelSizePx,
  };
}

export function updateDimensionResizeGesture(
  gesture: DimensionResizeGesture,
  currentPointerScreenPx: number,
): DimensionResizeGesture {
  if (!Number.isFinite(currentPointerScreenPx)) throw new Error('Dimension resize pointer must be finite');
  const deltaModelPx = (currentPointerScreenPx - gesture.startPointerScreenPx) / gesture.zoomScale;
  return {
    ...gesture,
    currentPointerScreenPx,
    currentModelSizePx: Math.max(
      gesture.minimumModelSizePx,
      Math.round(gesture.startModelSizePx + deltaModelPx),
    ),
  };
}
