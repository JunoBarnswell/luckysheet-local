import assert from 'node:assert/strict';
import { test } from 'node:test';
import { isChartHistogramOptions, type ChartHistogramOptions } from '@react-sheets/core-model';
import { chartHistogramOptionsForMode, chartHistogramOptionsWithNumber } from './chart-editor-state';

test('histogram editor mode changes preserve compatible tail thresholds and remove mode-owned fields', () => {
  const initial: ChartHistogramOptions = { mode: 'bin-width', binWidth: 2.5, underflow: 0, overflow: 10 };
  const byCount = chartHistogramOptionsForMode(initial, 'bin-count');
  assert.deepEqual(byCount, { mode: 'bin-count', binCount: 10, underflow: 0, overflow: 10 });
  assert.equal(isChartHistogramOptions(byCount), true);
  assert.deepEqual(chartHistogramOptionsForMode(byCount, 'by-category'), { mode: 'by-category' });
  assert.deepEqual(chartHistogramOptionsForMode(undefined, 'bin-width'), { mode: 'bin-width', binWidth: 1 });
});

test('histogram editor numeric changes clear optional bounds and leave invalid values rejectable', () => {
  const initial: ChartHistogramOptions = { mode: 'bin-width', binWidth: 2, underflow: 0, overflow: 10 };
  const updated = chartHistogramOptionsWithNumber(initial, 'underflow', '-5');
  assert.equal(updated.underflow, -5);
  assert.equal(isChartHistogramOptions(updated), true);
  const withoutOverflow = chartHistogramOptionsWithNumber(updated, 'overflow', '');
  assert.equal(withoutOverflow.overflow, undefined);
  assert.equal(isChartHistogramOptions(withoutOverflow), true);
  const invalidWidth = chartHistogramOptionsWithNumber(initial, 'binWidth', '0');
  assert.equal(isChartHistogramOptions(invalidWidth), false);
});
