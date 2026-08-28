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
