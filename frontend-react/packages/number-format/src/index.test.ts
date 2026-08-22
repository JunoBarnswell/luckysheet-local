import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { formatValue } from './index';

describe('number-format', () => {
  it('formats thousands separator', () => {
    assert.equal(formatValue(1234567.89, '#,##0'), '1,234,568');
    assert.equal(formatValue(1234.5, '#,##0.00'), '1,234.50');
  });

  it('scales percent formats', () => {
    assert.equal(formatValue(0.256, '0%'), '26%');
    assert.equal(formatValue(0.256, '0.0%'), '25.6%');
  });

  it('applies currency literals', () => {
    assert.equal(formatValue(1234, '"$"#,##0'), '$1,234');
  });

  it('falls back for text values', () => {
    assert.equal(formatValue('abc' as never, '#,##0'), 'abc');
    assert.equal(formatValue(null as never), '');
  });

  it('uses general formatting without a format code', () => {
    assert.equal(formatValue(3.5), '3.5');
    assert.equal(formatValue(1000), '1000');
  });
});
