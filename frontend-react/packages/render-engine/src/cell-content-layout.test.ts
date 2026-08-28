import assert from 'node:assert/strict';
import test from 'node:test';
import { DEFAULT_RENDER_THEME, resolveCellContentLayout, type CellRenderData } from './index';

function context(): CanvasRenderingContext2D {
  return {
    save() {},
    restore() {},
    measureText(text: string) { return { width: text.length * 8 }; },
    font: '',
  } as unknown as CanvasRenderingContext2D;
}

function cell(value: CellRenderData['value'], style: CellRenderData['style'] = {}): CellRenderData {
  return { value, displayValue: value === null ? '' : String(value), style };
}

test('CellContentLayoutDomain expands static overflow in the alignment direction and stops at occupied cells', () => {
  const base = { x: 40, y: 0, width: 40, height: 20 };
  const right = resolveCellContentLayout({
    context: context(),
    cell: cell('abcdefghij', { horizontalAlignment: 'left' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'abcdefghij',
    cellRect: base,
    cellRange: { startColumn: 1, endColumn: 1 },
    mode: 'display',
    neighborOccupancy: { left: [{ column: 0, widthPx: 40, occupied: false }], right: [{ column: 2, widthPx: 40, occupied: false }, { column: 3, widthPx: 40, occupied: true }] },
  });
  assert.equal(right.displayRect.x, 40);
  assert.equal(right.displayRect.width, 80);
  assert.deepEqual(right.overflowSpan, { startColumn: 1, endColumn: 2 });

  const left = resolveCellContentLayout({
    context: context(),
    cell: cell('abcdefghij', { horizontalAlignment: 'right' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'abcdefghij',
    cellRect: base,
    cellRange: { startColumn: 1, endColumn: 1 },
    mode: 'display',
    neighborOccupancy: { left: [{ column: 0, widthPx: 40, occupied: false }], right: [{ column: 2, widthPx: 40, occupied: true }] },
  });
  assert.equal(left.displayRect.x, 0);
  assert.equal(left.displayRect.width, 80);
  assert.deepEqual(left.overflowSpan, { startColumn: 0, endColumn: 1 });

  const centered = resolveCellContentLayout({
    context: context(),
    cell: cell('abcdefghij', { horizontalAlignment: 'center' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'abcdefghij',
    cellRect: base,
    cellRange: { startColumn: 1, endColumn: 1 },
    mode: 'display',
    neighborOccupancy: { left: [{ column: 0, widthPx: 40, occupied: false }], right: [{ column: 2, widthPx: 40, occupied: false }] },
  });
  assert.ok(centered.displayRect.x < base.x);
  assert.ok(centered.displayRect.width > base.width);
  assert.deepEqual(centered.overflowSpan, { startColumn: 0, endColumn: 2 });
});

test('edit geometry grows over occupied neighbors, preserves alignment, and remains inside the viewport', () => {
  const layout = resolveCellContentLayout({
    context: context(),
    cell: cell('long draft value', { horizontalAlignment: 'left' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'long draft value',
    cellRect: { x: 120, y: 20, width: 40, height: 20 },
    cellRange: { startColumn: 3, endColumn: 3 },
    mode: 'edit',
    viewportRect: { x: 0, y: 0, width: 180, height: 120 },
    neighborOccupancy: { left: [{ column: 2, widthPx: 40, occupied: true }, { column: 1, widthPx: 40, occupied: true }], right: [{ column: 4, widthPx: 40, occupied: true }] },
    caret: { start: 16, end: 16 },
  });
  assert.ok(layout.editRect.width >= 40);
  assert.ok(layout.editRect.x >= 0);
  assert.ok(layout.editRect.x + layout.editRect.width <= 180);
  assert.equal(layout.caretGeometry?.visible, true);
  assert.equal(layout.horizontalAlignment, 'left');
});

test('static overflow is blocked by numeric, wrapped, presented, controlled, and merged content contracts', () => {
  const neighbors = { left: [], right: [{ column: 1, widthPx: 120, occupied: false }] };
  for (const candidate of [
    cell(42, { horizontalAlignment: 'left' }),
    { ...cell('long text', { horizontalAlignment: 'left', wrapText: true }), value: 'long text' },
    { ...cell('long text', { horizontalAlignment: 'left' }), formula: '=A1' },
    { ...cell('long text', { horizontalAlignment: 'left' }), presentation: { kind: 'barcode', symbology: 'qr', parameters: { symbology: 'qr' }, options: { quietZone: 1, foreground: '#000000', background: '#ffffff', showText: false, labelPosition: 'none' } } },
    { ...cell('long text', { horizontalAlignment: 'left' }), editor: { kind: 'checkbox' } },
  ] as CellRenderData[]) {
    const result = resolveCellContentLayout({
      context: context(),
      cell: candidate,
      theme: DEFAULT_RENDER_THEME,
      text: String(candidate.value),
      cellRect: { x: 0, y: 0, width: 24, height: 20 },
      mode: 'display',
      neighborOccupancy: neighbors,
    });
    assert.deepEqual(result.displayRect, { x: 0, y: 0, width: 24, height: 20 });
  }
});

test('wrap and shrink use the same measured font metrics for editor and AutoFit geometry', () => {
  const wrapped = resolveCellContentLayout({
    context: context(),
    cell: cell('a\nbb', { wrapText: true }),
    theme: DEFAULT_RENDER_THEME,
    text: 'a\nbb',
    cellRect: { x: 0, y: 0, width: 32, height: 20 },
    mode: 'edit',
    caret: { start: 4, end: 4 },
  });
  assert.deepEqual(wrapped.lines, ['a', 'bb']);
  assert.ok(wrapped.editRect.height > 20);
  assert.equal(wrapped.multiline, true);
  assert.equal(wrapped.caretGeometry?.lineIndex, 1);

  const shrunk = resolveCellContentLayout({
    context: context(),
    cell: cell('abcdefghij', { shrinkToFit: true }),
    theme: DEFAULT_RENDER_THEME,
    text: 'abcdefghij',
    cellRect: { x: 0, y: 0, width: 40, height: 20 },
    mode: 'display',
  });
  assert.ok(shrunk.fontSizePx < 13);
  assert.ok(shrunk.fontSizePx >= 8);
});

test('merged and center-across cells use their span as the layout boundary', () => {
  const merged = resolveCellContentLayout({
    context: context(),
    cell: cell('merged value', { horizontalAlignment: 'center' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'merged value',
    cellRect: { x: 0, y: 0, width: 40, height: 20 },
    mergedRect: { x: 0, y: 0, width: 120, height: 20 },
    cellRange: { startColumn: 0, endColumn: 2 },
    mode: 'display',
    neighborOccupancy: { left: [], right: [] },
  });
  assert.equal(merged.displayRect.width, 120);
  assert.equal(merged.contentRect.width, 108);

  const centerAcross = resolveCellContentLayout({
    context: context(),
    cell: cell('across', { horizontalAlignment: 'centerContinuous' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'across',
    cellRect: { x: 40, y: 0, width: 40, height: 20 },
    alignmentSpan: { x: 40, y: 0, width: 80, height: 20 },
    cellRange: { startColumn: 1, endColumn: 2 },
    mode: 'display',
  });
  assert.equal(centerAcross.displayRect.width, 80);
  assert.equal(centerAcross.horizontalAlignment, 'centerContinuous');
});

test('rich-text runs participate in the shared font and line-height measurement', () => {
  const rich = resolveCellContentLayout({
    context: context(),
    cell: { value: 'AB', displayValue: 'AB', richText: [{ text: 'A', style: { bold: true, fontSizePx: 20 } }, { text: 'B', style: { italic: true } }] },
    theme: DEFAULT_RENDER_THEME,
    text: 'AB',
    cellRect: { x: 0, y: 0, width: 80, height: 20 },
    mode: 'display',
  });
  assert.equal(rich.fontRuns.length, 2);
  assert.ok(rich.fontRuns[0]!.font.includes('700'));
  assert.ok(rich.lineHeightPx >= 25);
});

test('rotated and stacked text contribute their occupied geometry instead of using an unrotated width', () => {
  const rotated = resolveCellContentLayout({
    context: context(),
    cell: cell('rotate', { textOrientation: 'rotateUp' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'rotate',
    cellRect: { x: 0, y: 0, width: 40, height: 20 },
    mode: 'display',
  });
  assert.ok(rotated.heightPx > rotated.widthPx / 2);
  const stacked = resolveCellContentLayout({
    context: context(),
    cell: cell('AB', { textOrientation: 'stacked' }),
    theme: DEFAULT_RENDER_THEME,
    text: 'AB',
    cellRect: { x: 0, y: 0, width: 40, height: 20 },
    mode: 'display',
  });
  assert.deepEqual(stacked.lines, ['A', 'B']);
  assert.ok(stacked.heightPx > 20);
});
