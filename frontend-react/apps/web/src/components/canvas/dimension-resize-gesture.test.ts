import test from 'node:test';
import assert from 'node:assert/strict';
import { beginDimensionResizeGesture, updateDimensionResizeGesture } from './dimension-resize-gesture';

test('column resize continuously reaches 1px and re-expands from the same locked boundary', () => {
  const started = beginDimensionResizeGesture({
    axis: 'column',
    boundaryIndex: 2,
    startModelSizePx: 64,
    startPointerScreenPx: 200,
    zoomScale: 1,
    minimumModelSizePx: 1,
  });
  assert.equal(updateDimensionResizeGesture(started, 151).currentModelSizePx, 15);
  assert.equal(updateDimensionResizeGesture(started, 137).currentModelSizePx, 1);
  const widened = updateDimensionResizeGesture(updateDimensionResizeGesture(started, 137), 256);
  assert.equal(widened.currentModelSizePx, 120);
  assert.equal(widened.boundaryIndex, 2);
});

test('dimension resize converts screen delta through one zoom scale for preview and commit', () => {
  const zoomed = beginDimensionResizeGesture({
    axis: 'column',
    boundaryIndex: 0,
    startModelSizePx: 80,
    startPointerScreenPx: 100,
    zoomScale: 1.25,
    minimumModelSizePx: 1,
  });
  assert.equal(updateDimensionResizeGesture(zoomed, 75).currentModelSizePx, 60);
  assert.equal(updateDimensionResizeGesture(zoomed, 125).currentModelSizePx, 100);
});

test('dimension resize rejects invalid ownership inputs before a gesture starts', () => {
  assert.throws(() => beginDimensionResizeGesture({
    axis: 'column',
    boundaryIndex: -1,
    startModelSizePx: 64,
    startPointerScreenPx: 0,
    zoomScale: 1,
    minimumModelSizePx: 1,
  }), /boundary index/);
});
