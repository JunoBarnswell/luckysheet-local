import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { isChartHistogramOptions } from './index';

describe('ChartHistogramOptions contract', () => {
  it('accepts finite Excel bin settings and ordered tail boundaries', () => {
    assert.equal(isChartHistogramOptions({ mode: 'automatic', underflow: 0, overflow: 10 }), true);
    assert.equal(isChartHistogramOptions({ mode: 'bin-width', binWidth: 2.5, underflow: 0, overflow: 10 }), true);
    assert.equal(isChartHistogramOptions({ mode: 'bin-count', binCount: 4, underflow: 0, overflow: 10 }), true);
    assert.equal(isChartHistogramOptions({ mode: 'bin-count', binCount: 2, underflow: 5, overflow: 5 }), true);
    assert.equal(isChartHistogramOptions({ mode: 'by-category' }), true);
  });

  it('rejects ignored, malformed, non-finite, and contradictory settings', () => {
    assert.equal(isChartHistogramOptions({ mode: 'bin-width', binWidth: 0 }), false);
    assert.equal(isChartHistogramOptions({ mode: 'bin-width', binWidth: Number.POSITIVE_INFINITY }), false);
    assert.equal(isChartHistogramOptions({ mode: 'bin-count', binCount: 2.5 }), false);
    assert.equal(isChartHistogramOptions({ mode: 'automatic', binWidth: 2 }), false);
    assert.equal(isChartHistogramOptions({ mode: 'automatic', binWidth: undefined }), false);
    assert.equal(isChartHistogramOptions({ mode: 'by-category', underflow: 0 }), false);
    assert.equal(isChartHistogramOptions({ mode: 'automatic', underflow: 11, overflow: 10 }), false);
    assert.equal(isChartHistogramOptions({ mode: 'automatic', unknown: true }), false);
    assert.equal(isChartHistogramOptions(Object.create({ mode: 'automatic' }) as unknown), false);
  });
});
