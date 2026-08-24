import assert from 'node:assert/strict';
import test from 'node:test';
import { excelColumnWidthToPixels, pixelsToExcelColumnWidth, pixelsToPoints, pointsToPixels } from './ooxml-metrics';

test('OOXML metrics use 96-DPI CSS pixels and the official maximum-digit-width formula', () => {
  assert.equal(excelColumnWidthToPixels(8.7109375, 7), 61);
  assert.equal(pointsToPixels(15), 20);
  assert.equal(pointsToPixels(18), 24);
  assert.equal(pixelsToPoints(20), 15);
});
test('column width round-trip stays within one rendered pixel', () => {
  for (const pixels of [1, 8, 20, 61, 68, 120, 255, 500, 1_785]) {
    const width = pixelsToExcelColumnWidth(pixels, 7);
    assert.ok(Math.abs(excelColumnWidthToPixels(width, 7) - pixels) <= 1, `${pixels}px -> ${width}`);
  }
});
