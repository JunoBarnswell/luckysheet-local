import test from 'node:test';
import assert from 'node:assert/strict';
import {
  CanvasRenderEngine,
  DirtyRangeSet,
  SheetSkeleton,
  Viewport,
  calculateRenderPlan,
  calculateScrollDelta,
  createEmptyChromeState,
  defaultHeaderOffset,
  mergeCellRanges,
  rangeToViewportRect,
} from './index';

const skeleton = new SheetSkeleton({
  rowCount: 20,
  columnCount: 10,
  defaultRowHeight: 20,
  defaultColumnWidth: 80,
  rowHeights: new Map([[2, 30]]),
  columnWidths: new Map([[1, 120]]),
});

function viewport(overrides: Partial<ReturnType<Viewport['getSnapshot']>> = {}) {
  return {
    width: 240,
    height: 100,
    scrollX: 0,
    scrollY: 0,
    devicePixelRatio: 1,
    ...overrides,
  };
}

test('SheetSkeleton calculates variable cell geometry and visible ranges', () => {
  assert.deepEqual(skeleton.getCellRect(2, 1), { x: 80, y: 40, width: 120, height: 30 });
  assert.deepEqual(skeleton.getVisibleRange({ x: 70, y: 35, width: 140, height: 40 }), {
    startRow: 1,
    endRow: 3,
    startColumn: 0,
    endColumn: 2,
  });
  assert.deepEqual(skeleton.getCellAtPoint({ x: 100, y: 50 }), { row: 2, column: 1 });
});

test('DirtyRangeSet normalizes, merges, and returns defensive copies', () => {
  const dirty = new DirtyRangeSet();
  dirty.add({ startRow: 3, endRow: 3, startColumn: 3, endColumn: 3 });
  dirty.add({ startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 });
  dirty.add({ startRow: 0, endRow: 0, startColumn: 8, endColumn: 9 });
  assert.deepEqual(dirty.toArray(), [
    { startRow: 0, endRow: 0, startColumn: 8, endColumn: 9 },
    { startRow: 2, endRow: 3, startColumn: 2, endColumn: 3 },
  ]);
  const copy = dirty.toArray();
  copy[0]!.startRow = 99;
  assert.equal(dirty.toArray()[0]?.startRow, 0);
  assert.deepEqual(mergeCellRanges([{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }]), [
    { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
  ]);
});

test('small scroll delta produces a reusable copy and exposed strips', () => {
  const plan = calculateScrollDelta(viewport(), viewport({ scrollX: 12, scrollY: 8 }));
  assert.equal(plan.canBlit, true);
  assert.deepEqual(plan.source, { x: 12, y: 8, width: 228, height: 92 });
  assert.deepEqual(plan.destination, { x: 0, y: 0, width: 228, height: 92 });
  assert.deepEqual(plan.exposedRects, [
    { x: 0, y: 92, width: 240, height: 8 },
    { x: 228, y: 0, width: 12, height: 100 },
  ]);
  assert.equal(calculateScrollDelta(viewport(), viewport({ scrollX: 240 })).canBlit, false);
});

test('RenderPlan selects initial, dirty, scroll, and overlay lifecycle modes', () => {
  const initial = calculateRenderPlan({ skeleton, viewport: viewport() });
  assert.equal(initial.fullRedraw, true);
  assert.equal(initial.reason, 'initial');
  assert.deepEqual(initial.layers.map((layer) => layer.mode), ['full', 'full', 'full', 'full', 'full']);

  const dirty = calculateRenderPlan({
    skeleton,
    viewport: viewport(),
    previousViewport: viewport(),
    dirtyRanges: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
  });
  assert.equal(dirty.fullRedraw, false);
  assert.equal(dirty.layers[0]?.mode, 'dirty');
  assert.equal(dirty.dirtyRects.length, 1);
  assert.deepEqual(rangeToViewportRect({ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }, skeleton, viewport()), {
    x: 80,
    y: 20,
    width: 120,
    height: 20,
  });

  const scroll = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollX: 10 }),
    previousViewport: viewport(),
  });
  assert.equal(scroll.layers[0]?.mode, 'scroll');
  assert.equal(scroll.layers[1]?.mode, 'scroll');
  assert.equal(scroll.layers[2]?.mode, 'scroll');
  assert.equal(scroll.layers[3]?.mode, 'full');
});

test('RenderPlan scroll blits scrollable layers even with header offset', () => {
  const scroll = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollX: 10 }),
    previousViewport: viewport(),
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(scroll.fullRedraw, false);
  assert.equal(scroll.layers[0]?.mode, 'scroll');
  assert.equal(scroll.layers[4]?.mode, 'full');
});

test('RenderPlan chromeDirty only redraws chrome layer', () => {
  const chromeOnly = calculateRenderPlan({
    skeleton,
    viewport: viewport(),
    previousViewport: viewport(),
    chromeDirty: true,
  });
  assert.equal(chromeOnly.fullRedraw, false);
  assert.deepEqual(chromeOnly.layers.map((layer) => layer.mode), ['none', 'none', 'none', 'none', 'full']);
});

test('CanvasRenderEngine invalidateChrome only marks chrome layer', () => {
  const engine = new CanvasRenderEngine({ skeleton, viewport: viewport() });
  engine.render();
  engine.setChrome(createEmptyChromeState());
  const chromeOnly = engine.render();
  assert.equal(chromeOnly.fullRedraw, false);
  assert.equal(chromeOnly.layers.find((layer) => layer.layerId === 'chrome')?.mode, 'full');
  assert.equal(chromeOnly.layers.find((layer) => layer.layerId === 'grid')?.mode, 'none');
  engine.dispose();
});

test('RenderPlan uses scroll blit on grid layers even with header offset', () => {
  const scroll = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollX: 10 }),
    previousViewport: viewport(),
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(scroll.fullRedraw, false);
  assert.equal(scroll.layers[0]?.mode, 'scroll');
  assert.equal(scroll.layers[4]?.mode, 'full');
});

test('RenderPlan chromeDirty only repaints chrome layer', () => {
  const chromeOnly = calculateRenderPlan({
    skeleton,
    viewport: viewport(),
    previousViewport: viewport(),
    chromeDirty: true,
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(chromeOnly.fullRedraw, false);
  assert.equal(chromeOnly.layers[0]?.mode, 'none');
  assert.equal(chromeOnly.layers[4]?.mode, 'full');
});

test('CanvasRenderEngine keeps rendering state independent from DOM mounting', () => {
  const engine = new CanvasRenderEngine({ skeleton, viewport: viewport() });
  const initial = engine.render();
  assert.equal(initial.fullRedraw, true);
  engine.scrollTo(10, 0);
  const scrolled = engine.render();
  assert.equal(scrolled.scrollDelta.canBlit, true);
  engine.invalidate([{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }]);
  const dirty = engine.render();
  assert.equal(dirty.layers[0]?.mode, 'dirty');
  engine.dispose();
  assert.throws(() => engine.render(), /disposed/);
});

test('SheetSkeleton handles custom dimensions and coordinate lookups accurately', () => {
  const customSkeleton = new SheetSkeleton({
    rowCount: 100,
    columnCount: 50,
    defaultRowHeight: 24,
    defaultColumnWidth: 100,
  });

  customSkeleton.setRowHeight(0, 48);
  customSkeleton.setColumnWidth(0, 150);

  assert.equal(customSkeleton.getRowHeight(0), 48);
  assert.equal(customSkeleton.getRowHeight(1), 24);
  assert.equal(customSkeleton.getColumnWidth(0), 150);
  assert.equal(customSkeleton.getColumnWidth(1), 100);

  assert.equal(customSkeleton.getRowTop(0), 0);
  assert.equal(customSkeleton.getRowTop(1), 48);
  assert.equal(customSkeleton.getRowTop(2), 72);

  assert.equal(customSkeleton.getColumnLeft(0), 0);
  assert.equal(customSkeleton.getColumnLeft(1), 150);
  assert.equal(customSkeleton.getColumnLeft(2), 250);

  assert.deepEqual(customSkeleton.getCellAtPoint({ x: 50, y: 20 }), { row: 0, column: 0 });
  assert.deepEqual(customSkeleton.getCellAtPoint({ x: 160, y: 50 }), { row: 1, column: 1 });
});

test('SheetSkeleton virtualizes very large uniform dimensions without dense arrays', () => {
  const largeSkeleton = new SheetSkeleton({
    rowCount: 1_000_000,
    columnCount: 32,
    defaultRowHeight: 20,
    defaultColumnWidth: 80,
  });

  assert.equal(largeSkeleton.totalHeight, 20_000_000);
  assert.equal(largeSkeleton.getRowTop(999_999), 19_999_980);
  assert.deepEqual(largeSkeleton.findRowAt(19_999_985), 999_999);
  assert.deepEqual(largeSkeleton.getVisibleRange({ x: 0, y: 19_999_980, width: 160, height: 20 }), {
    startRow: 999_999,
    endRow: 999_999,
    startColumn: 0,
    endColumn: 1,
  });
  assert.equal(largeSkeleton.getVisibleRowModels().length, 0);
});
