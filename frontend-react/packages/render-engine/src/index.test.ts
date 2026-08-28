import test from 'node:test';
import assert from 'node:assert/strict';
import {
  CanvasRenderEngine,
  DirtyRangeSet,
  SheetSkeleton,
  Viewport,
  calculateRenderPlan,
  computePaneMap,
  calculateScrollDelta,
  createEmptyChromeState,
  defaultHeaderOffset,
  isColumnSelected,
  isRowSelected,
  isSelectAllRange,
  mergeCellRanges,
  rangeToViewportRect,
  hasMeasurableCellContent,
} from './index';
import { drawCellLayer, drawGridLayer } from './cell-renderer';
import { drawChromeLayer } from './chrome-renderer';
import { DEFAULT_RENDER_THEME, type CellRenderData, type RenderPane } from './types';

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

function percentile95(samples: readonly number[]): number {
  const ordered = [...samples].sort((left, right) => left - right);
  return ordered[Math.min(ordered.length - 1, Math.ceil(ordered.length * 0.95) - 1)] ?? 0;
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

test('scroll delta records movement without exposing bitmap reuse state', () => {
  const plan = calculateScrollDelta(viewport(), viewport({ scrollX: 12, scrollY: 8 }));
  assert.deepEqual(plan, {
    delta: { x: 12, y: 8 },
    dx: 12,
    dy: 8,
    hasDelta: true,
    redrawRects: [],
  });
});

test('fractional device-pixel scroll redraws the affected pane', () => {
  const fractional = calculateScrollDelta(viewport(), viewport({ scrollY: 8.5 }));
  assert.equal(fractional.hasDelta, true);

  const plan = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollY: 8.5 }),
    previousViewport: viewport(),
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(plan.reason, 'scroll-redraw');
  assert.deepEqual(plan.scrollDelta.redrawRects, [
    { x: 39, y: 20, width: 201, height: 80 },
  ]);
  assert.deepEqual(plan.layers.find((layer) => layer.layerId === 'grid')?.clearRects, [
    { x: 39, y: 20, width: 201, height: 80 },
  ]);
});

test('device-pixel-aligned scroll still redraws instead of reusing canvas pixels', () => {
  const highDpiViewport = viewport({ devicePixelRatio: 1.25 });
  const aligned = calculateRenderPlan({
    skeleton,
    previousViewport: highDpiViewport,
    viewport: { ...highDpiViewport, scrollY: 0.8 },
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(aligned.reason, 'scroll-redraw');
  assert.equal(aligned.layers[0]?.mode, 'dirty');
  assert.deepEqual(aligned.scrollDelta.redrawRects, [{ x: 39, y: 20, width: 201, height: 80 }]);
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
  assert.equal(scroll.layers[0]?.mode, 'dirty');
  assert.equal(scroll.layers[1]?.mode, 'dirty');
  assert.equal(scroll.layers[2]?.mode, 'dirty');
  assert.equal(scroll.layers[3]?.mode, 'full');
});

test('RenderPlan redraws scrollable panes while preserving header offset', () => {
  const scroll = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollX: 10 }),
    previousViewport: viewport(),
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(scroll.fullRedraw, false);
  assert.equal(scroll.layers[0]?.mode, 'dirty');
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

test('RenderPlan redraws grid layers without canvas self-copy', () => {
  const scroll = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollX: 10 }),
    previousViewport: viewport(),
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(scroll.fullRedraw, false);
  assert.equal(scroll.layers[0]?.mode, 'dirty');
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
  assert.equal(scrolled.reason, 'scroll-redraw');
  assert.deepEqual(scrolled.scrollDelta.redrawRects, [{ x: 39, y: 20, width: 201, height: 80 }]);
  engine.invalidate([{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }]);
  const dirty = engine.render();
  assert.equal(dirty.layers[0]?.mode, 'dirty');
  engine.dispose();
  assert.throws(() => engine.render(), /disposed/);
});

test('CanvasRenderEngine clears and redraws the visible pane on every scroll', () => {
  const engine = new CanvasRenderEngine({ skeleton, viewport: viewport() });
  const initial = engine.render();
  assert.equal(initial.fullRedraw, true);
  engine.scrollTo(0, 40);
  const scrolled = engine.render();
  assert.equal(scrolled.fullRedraw, false);
  assert.equal(scrolled.layers.find((layer) => layer.layerId === 'grid')?.mode, 'dirty');
  assert.deepEqual(scrolled.layers.find((layer) => layer.layerId === 'grid')?.clearRects, [
    { x: 39, y: 20, width: 201, height: 80 },
  ]);
  assert.equal(scrolled.layers.find((layer) => layer.layerId === 'chrome')?.mode, 'full');
  engine.dispose();
});

test('RenderPlan keeps frozen panes disjoint while scrolling only their movable axes', () => {
  const frozen = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollX: 10, scrollY: 8 }),
    previousViewport: viewport(),
    pane: { kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 0, startColumn: 0, state: 'frozen' },
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(frozen.fullRedraw, false);
  assert.deepEqual(frozen.scrollDelta.redrawRects, [
    { x: 119, y: 20, width: 121, height: 20 },
    { x: 39, y: 40, width: 80, height: 60 },
    { x: 119, y: 40, width: 121, height: 60 },
  ]);
  assert.ok(frozen.scrollDelta.redrawRects.every((rect) => rect.x >= 39 && rect.y >= 20));
});

test('RenderPlan redraws a scrollbar jump inside the current pane only', () => {
  const jumped = calculateRenderPlan({
    skeleton,
    viewport: viewport({ scrollY: 400 }),
    previousViewport: viewport(),
    headerOffset: defaultHeaderOffset(),
  });
  assert.equal(jumped.fullRedraw, false);
  assert.equal(jumped.reason, 'scroll-redraw');
  assert.deepEqual(jumped.layers.find((layer) => layer.layerId === 'grid')?.clearRects, [{ x: 39, y: 20, width: 201, height: 80 }]);
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

test('nearest column boundary remains stable for adjacent narrow columns', () => {
  const narrow = new SheetSkeleton({
    rowCount: 2,
    columnCount: 4,
    defaultRowHeight: 20,
    defaultColumnWidth: 3,
  });
  assert.deepEqual(narrow.findNearestColumnBoundary(4.2, 4), { index: 0, deltaPx: -1.2000000000000002 });
  assert.deepEqual(narrow.findNearestColumnBoundary(4.8, 4), { index: 1, deltaPx: 1.2000000000000002 });
  assert.deepEqual(narrow.findNearestColumnBoundary(11.5, 4), { index: 3, deltaPx: 0.5 });
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

test('scroll planning stays inside one frame for 100k and 1m logical rows', () => {
  for (const rowCount of [100_000, 1_000_000]) {
    const largeSkeleton = new SheetSkeleton({
      rowCount,
      columnCount: 20,
      defaultRowHeight: 20,
      defaultColumnWidth: 80,
    });
    const samples: number[] = [];
    let previous = { width: 1440, height: 760, scrollX: 0, scrollY: 0, devicePixelRatio: 1 };
    for (let index = 1; index <= 120; index += 1) {
      const next = { ...previous, scrollY: (largeSkeleton.totalHeight - previous.height) * (index / 121) };
      const started = performance.now();
      const plan = calculateRenderPlan({ skeleton: largeSkeleton, viewport: next, previousViewport: previous, headerOffset: defaultHeaderOffset() });
      samples.push(performance.now() - started);
      assert.ok(plan.visibleRange && plan.visibleRange.endRow < rowCount);
      assert.ok(plan.layers.some((layer) => layer.layerId === 'grid' && layer.mode === 'dirty'));
      previous = next;
    }
    assert.ok(percentile95(samples) <= 16.7, `${rowCount} rows p95 must stay inside a 60fps frame, got ${percentile95(samples).toFixed(3)}ms`);
  }
});

test('hidden rows collapse layout geometry without losing model row identity', () => {
  const hiddenSkeleton = new SheetSkeleton({ rowCount: 6, columnCount: 3, defaultRowHeight: 20, defaultColumnWidth: 50, hiddenRows: new Set([2, 4]) });
  assert.equal(hiddenSkeleton.getRowHeight(2), 0);
  assert.equal(hiddenSkeleton.getRowTop(3), 40);
  assert.equal(hiddenSkeleton.getCellRect(2, 0), null);
  assert.deepEqual(hiddenSkeleton.getCellAtPoint({ x: 25, y: 41 }), { row: 3, column: 0 });
  assert.equal(hiddenSkeleton.totalHeight, 80);
});

test('chrome renders hidden row and column double-line indicators at the canonical collapsed boundary', () => {
  const hiddenSkeleton = new SheetSkeleton({ rowCount: 6, columnCount: 6, defaultRowHeight: 20, defaultColumnWidth: 50, hiddenRows: new Set([1, 2]), hiddenColumns: new Set([1, 2]) });
  const plan = calculateRenderPlan({ skeleton: hiddenSkeleton, viewport: viewport({ width: 220, height: 120 }), headerOffset: defaultHeaderOffset() });
  const { context, lineCalls } = recordingContext();
  drawChromeLayer({ context, skeleton: hiddenSkeleton, plan, chrome: createEmptyChromeState(), theme: DEFAULT_RENDER_THEME });
  assert.ok(lineCalls.some(({ from, to }) => Math.abs(from[0] - 87) < 0.01 && Math.abs(to[0] - 87) < 0.01));
  assert.ok(lineCalls.some(({ from, to }) => Math.abs(from[1] - 38) < 0.01 && Math.abs(to[1] - 38) < 0.01));
  const engine = new CanvasRenderEngine({ skeleton: hiddenSkeleton, viewport: viewport({ width: 220, height: 120 }) });
  assert.deepEqual(engine.headerHitAtLocal({ x: 89, y: 10 })?.hiddenIndices, [1, 2]);
  assert.deepEqual(engine.headerHitAtLocal({ x: 10, y: 40 })?.hiddenIndices, [1, 2]);
});

function recordingContext() {
  const textCalls: Array<{ text: string; x: number; y: number }> = [];
  const lineCalls: Array<{ from: [number, number]; to: [number, number] }> = [];
  const fillCalls: Array<{ x: number; y: number; width: number; height: number }> = [];
  const textAlignValues: CanvasTextAlign[] = [];
  let current: [number, number] = [0, 0];
  let currentTextAlign: CanvasTextAlign = 'left';
  const context = {
    beginPath() {},
    closePath() {},
    save() {},
    restore() {},
    moveTo(x: number, y: number) { current = [x, y]; },
    lineTo(x: number, y: number) {
      lineCalls.push({ from: current, to: [x, y] });
      current = [x, y];
    },
    stroke() {},
    fill() {},
    fillRect(x: number, y: number, width: number, height: number) { fillCalls.push({ x, y, width, height }); },
    strokeRect() {},
    fillText(text: string, x: number, y: number) { textCalls.push({ text, x, y }); },
    measureText() { return { width: 8 }; },
    setLineDash() {},
    translate() {},
    rotate() {},
    arc() {},
    rect() {},
    ellipse() {},
    roundRect() {},
    clip() {},
    clearRect() {},
    setTransform() {},
    globalAlpha: 1,
    lineWidth: 1,
    fillStyle: '',
    strokeStyle: '',
    font: '',
    textBaseline: 'middle' as CanvasTextBaseline,
    get textAlign() { return currentTextAlign; },
    set textAlign(value: CanvasTextAlign) { currentTextAlign = value; textAlignValues.push(value); },
  } as unknown as CanvasRenderingContext2D;
  return { context, textCalls, lineCalls, fillCalls, textAlignValues };
}

function mainPane(range: { startRow: number; endRow: number; startColumn: number; endColumn: number }): RenderPane {
  return {
    id: 'main',
    screenRect: { x: 39, y: 20, width: 260, height: 180 },
    contentOrigin: { x: 0, y: 0 },
    visibleRange: range,
  };
}

test('cell rendering calls the provider address and paints the value only in that cell', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 12, columnCount: 8, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 10, startColumn: 0, endColumn: 4 };
  const provider = ({ row, column }: { row: number; column: number }): CellRenderData | undefined =>
    row === 8 && column === 2 ? { value: 4, displayValue: '4' } : undefined;
  const { context, textCalls } = recordingContext();

  drawCellLayer({
    context,
    skeleton: renderSkeleton,
    pane: mainPane(range),
    visibleRange: range,
    cellProvider: provider,
    theme: DEFAULT_RENDER_THEME,
  });

  assert.deepEqual(textCalls, [{ text: '4', x: 144, y: 170 }]);
});

test('cell rendering applies top, middle, and bottom vertical alignment to unwrapped text', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 1, columnCount: 3, defaultRowHeight: 60, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 };
  const provider = ({ column }: { row: number; column: number }): CellRenderData => ({ value: column, displayValue: String(column), style: { verticalAlignment: (['top', 'middle', 'bottom'] as const)[column] } });
  const { context, textCalls } = recordingContext();
  drawCellLayer({ context, skeleton: renderSkeleton, pane: mainPane(range), visibleRange: range, cellProvider: provider, theme: DEFAULT_RENDER_THEME });
  assert.equal(textCalls[0]?.y, 12.5);
  assert.equal(textCalls[1]?.y, 30);
  assert.equal(textCalls[2]?.y, 47.5);
});

test('extended alignment renders through a Canvas-safe projection and center-across span', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 1, columnCount: 5, defaultRowHeight: 30, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 4 };
  const provider = ({ column }: { row: number; column: number }): CellRenderData | undefined => {
    if (column <= 2) {
      return {
        value: 'Across',
        displayValue: 'Across',
        style: { horizontalAlignment: 'centerContinuous' },
        alignmentSpan: { startColumn: 0, endColumn: 2, isAnchor: column === 0 },
      };
    }
    const horizontalAlignment = column === 3 ? ('justify' as const) : ('fill' as const);
    return { value: 'Text', displayValue: 'Text', style: { horizontalAlignment } };
  };
  const { context, textCalls, textAlignValues } = recordingContext();
  drawCellLayer({ context, skeleton: renderSkeleton, pane: mainPane(range), visibleRange: range, cellProvider: provider, theme: DEFAULT_RENDER_THEME });

  assert.equal(textCalls.filter((call) => call.text === 'Across').length, 1);
  assert.equal(textCalls.find((call) => call.text === 'Across')?.x, 75);
  assert.ok(textAlignValues.every((value) => value === 'left' || value === 'center' || value === 'right'));
});

test('stacked orientation renders explicitly while unsupported native alignment stays out of Canvas', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 1, columnCount: 2, defaultRowHeight: 60, defaultColumnWidth: 60 });
  const range = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 };
  const provider = ({ column }: { row: number; column: number }): CellRenderData => column === 0
    ? { value: 'AB', displayValue: 'AB', style: { textOrientation: 'stacked' } }
    : { value: 'vendor', displayValue: 'vendor', style: { unsupportedAlignment: { horizontal: 'vendorAlignment' } } };
  const { context, textCalls, textAlignValues } = recordingContext();
  drawCellLayer({ context, skeleton: renderSkeleton, pane: mainPane(range), visibleRange: range, cellProvider: provider, theme: DEFAULT_RENDER_THEME });

  assert.deepEqual(textCalls.map((call) => call.text), ['A', 'B', 'Unsupported alignment']);
  assert.ok(textAlignValues.every((value) => value === 'left' || value === 'center' || value === 'right'));
});

test('blank cells retain complete horizontal and vertical grid boundaries', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 3, columnCount: 3, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 };
  const { context, lineCalls } = recordingContext();
  drawGridLayer({
    context,
    skeleton: renderSkeleton,
    pane: mainPane(range),
    visibleRange: range,
    cellProvider: () => undefined,
    theme: DEFAULT_RENDER_THEME,
  });

  const horizontal = new Set(lineCalls
    .filter(({ from, to }) => from[1] === to[1] && from[0] !== to[0])
    .map(({ from, to }) => `${from[1]}:${from[0]}-${to[0]}`));
  const vertical = new Set(lineCalls
    .filter(({ from, to }) => from[0] === to[0] && from[1] !== to[1])
    .map(({ from, to }) => `${from[0]}:${from[1]}-${to[1]}`));
  for (const y of [0.5, 20.5, 40.5, 60.5]) {
    for (const segment of ['0.5-50.5', '50.5-100.5', '100.5-150.5']) assert.ok(horizontal.has(`${y}:${segment}`));
  }
  for (const x of [0.5, 50.5, 100.5, 150.5]) {
    for (const segment of ['0.5-20.5', '20.5-40.5', '40.5-60.5']) assert.ok(vertical.has(`${x}:${segment}`));
  }
});

test('incremental grid redraw fills only the exposed content strip', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 3, columnCount: 3, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 };
  const { context, fillCalls } = recordingContext();
  drawGridLayer({
    context,
    skeleton: renderSkeleton,
    pane: mainPane(range),
    visibleRange: range,
    cellProvider: () => undefined,
    theme: DEFAULT_RENDER_THEME,
    drawRects: [{ x: 100, y: 40, width: 50, height: 20 }],
  });
  assert.deepEqual(fillCalls, [{ x: 100, y: 40, width: 50, height: 20 }]);
});

test('incremental cell redraw reads only the exposed strip instead of the visible viewport', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 100, columnCount: 20, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const range = { startRow: 20, endRow: 49, startColumn: 0, endColumn: 19 };
  let reads = 0;
  const { context } = recordingContext();
  drawCellLayer({
    context,
    skeleton: renderSkeleton,
    pane: mainPane(range),
    visibleRange: range,
    cellProvider: () => {
      reads += 1;
      return undefined;
    },
    theme: DEFAULT_RENDER_THEME,
    drawRects: [{ x: 0, y: 980, width: 1000, height: 20 }],
  });

  // One exposed row has 20 cells. Each candidate is resolved once to retain
  // merge/center-across semantics and once while painting. A full viewport
  // scan would require 30 * 20 * 2 reads.
  assert.ok(reads <= 40, `expected at most 40 provider reads for an exposed row, got ${reads}`);
});

test('incremental cell redraw includes a merged anchor that begins outside the exposed strip', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 5, columnCount: 3, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 };
  const { context, textCalls } = recordingContext();
  drawCellLayer({
    context,
    skeleton: renderSkeleton,
    pane: mainPane(range),
    visibleRange: range,
    cellProvider: ({ row, column }) => {
      if (column !== 0 || (row !== 1 && row !== 2)) return undefined;
      return {
        value: 'Merged anchor',
        displayValue: 'Merged anchor',
        merge: { startRow: 1, endRow: 2, startColumn: 0, endColumn: 0, isAnchor: row === 1 },
      };
    },
    theme: DEFAULT_RENDER_THEME,
    drawRects: [{ x: 0, y: 40, width: 50, height: 20 }],
  });

  assert.equal(textCalls.filter((call) => call.text === 'Merged anchor').length, 1);
});

test('merged blank cells suppress only their internal grid boundaries', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 3, columnCount: 3, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const range = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 };
  const { context, lineCalls } = recordingContext();
  drawGridLayer({
    context,
    skeleton: renderSkeleton,
    pane: mainPane(range),
    visibleRange: range,
    cellProvider: ({ row, column }) => {
      if (row > 1 || column > 1) return undefined;
      return {
        value: undefined,
        merge: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1, isAnchor: row === 0 && column === 0 },
      };
    },
    theme: DEFAULT_RENDER_THEME,
  });

  const internalHorizontal = lineCalls.some(({ from, to }) => from[1] === 20.5 && from[0] < 100.5 && to[0] <= 100.5);
  const internalVertical = lineCalls.some(({ from, to }) => from[0] === 50.5 && from[1] < 40.5 && to[1] <= 40.5);
  assert.equal(internalHorizontal, false);
  assert.equal(internalVertical, false);
});

test('pane translation preserves model coordinates for C9', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 20, columnCount: 10, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const engine = new CanvasRenderEngine({
    skeleton: renderSkeleton,
    viewport: { width: 300, height: 220, scrollX: 0, scrollY: 0, devicePixelRatio: 1 },
  });
  engine.render();
  const cell = renderSkeleton.getCellRect(8, 2)!;
  assert.deepEqual(engine.cellAtLocalPoint({ x: 39 + cell.x + cell.width / 2, y: 20 + cell.y + cell.height / 2 }), { row: 8, column: 2 });
  engine.dispose();
});

test('contentToScreen selects the cell pane for frozen rows and columns', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 20, columnCount: 10, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const engine = new CanvasRenderEngine({
    skeleton: renderSkeleton,
    viewport: { width: 300, height: 220, scrollX: 0, scrollY: 0, devicePixelRatio: 1 },
  });
  engine.setPane({ kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1, state: 'frozen' });
  engine.render();
  const topLeft = renderSkeleton.getCellRect(0, 0)!;
  const main = renderSkeleton.getCellRect(4, 2)!;
  assert.deepEqual(engine.contentToScreen({ x: topLeft.x, y: topLeft.y }, { row: 0, column: 0 }), { x: 39, y: 20 });
  assert.deepEqual(engine.contentToScreen({ x: main.x, y: main.y }, { row: 4, column: 2 }), { x: 139, y: 100 });
  engine.dispose();
});

test('pivot-control child hits are returned before the generic floating move hit', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 20, columnCount: 10, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const engine = new CanvasRenderEngine({
    skeleton: renderSkeleton,
    viewport: { width: 300, height: 220, scrollX: 0, scrollY: 0, devicePixelRatio: 1 },
  });
  engine.render();
  engine.setFloating([{
    kind: 'pivot-control',
    id: 'slicer-1',
    bounds: { x: 0, y: 0, width: 160, height: 100 },
    draw: () => undefined,
    hitTest: (point) => point.y >= 26 && point.y < 46 ? { action: 'pivot.slicer.member', data: { member: 'text:A' } } : null,
  }]);
  const memberHit = engine.hitTestFloating({ x: 39 + 20, y: 20 + 35 });
  assert.equal(memberHit?.kind, 'pivot-control');
  assert.equal(memberHit?.control?.action, 'pivot.slicer.member');
  assert.deepEqual(memberHit?.control?.data, { member: 'text:A' });
  const bodyHit = engine.hitTestFloating({ x: 39 + 20, y: 20 + 70 });
  assert.equal(bodyHit?.kind, 'pivot-control');
  assert.equal(bodyHit?.control, undefined);
  engine.dispose();
});

test('saved frozen top-left cell seeds initial scroll without blocking earlier rows', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 200, columnCount: 10, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const engine = new CanvasRenderEngine({ skeleton: renderSkeleton, viewport: { width: 500, height: 320, scrollX: 0, scrollY: 0, devicePixelRatio: 1 } });
  engine.setPane({ kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 100, startColumn: 1, state: 'frozen' });
  assert.equal(engine.viewport.getSnapshot().scrollY, renderSkeleton.getRowTop(100));
  engine.scrollTo(0, 0);
  assert.equal(engine.viewport.getSnapshot().scrollY, 0);
  engine.render();
  assert.equal(engine.lastRenderPlan?.paneMap.paneForCell({ row: 1, column: 1 })?.id, 'main');
  engine.dispose();
});

test('split pane point geometry is independent of display DPR', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 50, columnCount: 20, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const pane = { kind: 'split' as const, xSplit: 1440, ySplit: 720, startRow: 0, startColumn: 0, state: 'split' as const };
  const one = computePaneMap(renderSkeleton, { width: 800, height: 500, scrollX: 0, scrollY: 0, devicePixelRatio: 1 }, pane, defaultHeaderOffset());
  const two = computePaneMap(renderSkeleton, { width: 800, height: 500, scrollX: 0, scrollY: 0, devicePixelRatio: 2 }, pane, defaultHeaderOffset());
  assert.equal(one.panes.find((entry) => entry.id === 'main')?.screenRect.x, two.panes.find((entry) => entry.id === 'main')?.screenRect.x);
  assert.equal(one.panes.find((entry) => entry.id === 'main')?.screenRect.y, two.panes.find((entry) => entry.id === 'main')?.screenRect.y);
});

test('frozen 2x2 pane map clamps the main origin and keeps four ranges disjoint', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 20, columnCount: 10, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const map = computePaneMap(renderSkeleton, { width: 500, height: 320, scrollX: 0, scrollY: 0, devicePixelRatio: 1 }, {
    kind: 'frozen', xSplit: 2, ySplit: 2, startRow: 0, startColumn: 0, state: 'frozen',
  }, defaultHeaderOffset());
  const byId = new Map(map.panes.map((pane) => [pane.id, pane]));
  assert.deepEqual(byId.get('topLeft')?.visibleRange, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  assert.equal(byId.get('topRight')?.visibleRange?.startRow, 0);
  assert.equal(byId.get('topRight')?.visibleRange?.startColumn, 2);
  assert.equal(byId.get('bottomLeft')?.visibleRange?.startRow, 2);
  assert.equal(byId.get('bottomLeft')?.visibleRange?.startColumn, 0);
  assert.equal(byId.get('main')?.visibleRange?.startRow, 2);
  assert.equal(byId.get('main')?.visibleRange?.startColumn, 2);
  assert.equal(map.paneForCell({ row: 0, column: 0 })?.id, 'topLeft');
  assert.equal(map.paneForCell({ row: 0, column: 2 })?.id, 'topRight');
  assert.equal(map.paneForCell({ row: 2, column: 0 })?.id, 'bottomLeft');
  assert.equal(map.paneForCell({ row: 2, column: 2 })?.id, 'main');
});

test('selection header projection highlights ordinary rectangles and explicit skeleton bounds', () => {
  const chrome = createEmptyChromeState();
  chrome.selection.ranges = [{ startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 }];
  assert.equal(isColumnSelected(chrome, 0), true);
  assert.equal(isColumnSelected(chrome, 2), true);
  assert.equal(isColumnSelected(chrome, 3), false);
  assert.equal(isRowSelected(chrome, 1), true);
  assert.equal(isRowSelected(chrome, 8), true);
  assert.equal(isRowSelected(chrome, 9), false);
  assert.equal(isSelectAllRange({ startRow: 0, endRow: 19, startColumn: 0, endColumn: 9 }, skeleton), true);
  assert.equal(isSelectAllRange({ startRow: 0, endRow: 999, startColumn: 0, endColumn: 25 }, skeleton), true);
  assert.equal(isSelectAllRange({ startRow: 0, endRow: 18, startColumn: 0, endColumn: 9 }, skeleton), false);
});

test('content range screen geometry splits overlays across frozen panes', () => {
  const renderSkeleton = new SheetSkeleton({ rowCount: 20, columnCount: 10, defaultRowHeight: 20, defaultColumnWidth: 50 });
  const engine = new CanvasRenderEngine({
    skeleton: renderSkeleton,
    viewport: { width: 300, height: 220, scrollX: 30, scrollY: 20, devicePixelRatio: 1 },
  });
  engine.setPane({ kind: 'frozen', xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1, state: 'frozen' });
  engine.render();
  const rects = engine.contentRangeToScreenRects({ startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 });
  assert.ok(rects.some((rect) => rect.x === 39 && rect.y === 20));
  assert.ok(rects.length >= 2);
  engine.dispose();
});

test('AutoFit content gate preserves typed zero, FALSE, formatted text and errors', () => {
  assert.equal(hasMeasurableCellContent(undefined), false);
  assert.equal(hasMeasurableCellContent({ value: null }), false);
  assert.equal(hasMeasurableCellContent({ value: undefined }), false);
  assert.equal(hasMeasurableCellContent({ value: '' }), false);
  assert.equal(hasMeasurableCellContent({ value: 0 }), true);
  assert.equal(hasMeasurableCellContent({ value: false }), true);
  assert.equal(hasMeasurableCellContent({ value: 0, displayValue: '0.00' }), true);
  assert.equal(hasMeasurableCellContent({ value: false, displayValue: 'FALSE' }), true);
  assert.equal(hasMeasurableCellContent({ value: null, displayValue: '#DIV/0!' }), true);
  assert.equal(hasMeasurableCellContent({ value: 'raw', displayValue: '' }), false);
});
