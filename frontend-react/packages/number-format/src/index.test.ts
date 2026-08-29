import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { formatNumberValue, formatValue, parseNumberFormat, transformNumberFormatPrecision } from './index';

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

  it('parses conditional/color/locale sections into one AST and selects them deterministically', () => {
    const ast = parseNumberFormat('[Red][>=100]#,##0.00;[Blue]#,##0.00');
    assert.equal(ast.sections[0]?.condition?.operator, '>=');
    assert.equal(ast.sections[0]?.color, 'Red');
    assert.equal(formatNumberValue(1250.5, ast).text, '1,250.50');
    assert.equal(formatNumberValue(1250.5, ast).color, 'Red');
    assert.equal(formatNumberValue(12.5, ast).color, 'Blue');
  });

  it('uses the workbook date system and locale context instead of a fixed epoch', () => {
    assert.equal(formatValue(1, 'yyyy-mm-dd', { dateSystem: '1900' }), '1900-01-01');
    assert.equal(formatValue(0, 'yyyy-mm-dd', { dateSystem: '1904' }), '1904-01-01');
    assert.equal(formatValue(1234.5, '#,##0.0', { locale: 'de-DE' }), '1.234,5');
    assert.equal(formatValue(48.5, '[h]:mm', { dateSystem: '1900' }), '1164:00');
  });

  it('changes only decimal placeholders in ordinary numeric formats', () => {
    assert.deepEqual(transformNumberFormatPrecision('#,##0', 1), { ok: true, format: '#,##0.0', decimalPlaces: 1 });
    assert.deepEqual(transformNumberFormatPrecision('$#,##0.00', -1), { ok: true, format: '$#,##0.0', decimalPlaces: 1 });
    assert.deepEqual(transformNumberFormatPrecision('0.00%', 1), { ok: true, format: '0.000%', decimalPlaces: 3 });
    assert.deepEqual(transformNumberFormatPrecision('0.0E+00', -1), { ok: true, format: '0E+00', decimalPlaces: 0 });
  });

  it('preserves sections, colors, conditions, quoted literals, and locale markers', () => {
    assert.deepEqual(
      transformNumberFormatPrecision('[Red]#,##0;[Blue]-#,##0', 1),
      { ok: true, format: '[Red]#,##0.0;[Blue]-#,##0.0', decimalPlaces: 1 },
    );
    assert.deepEqual(
      transformNumberFormatPrecision('[>=100]"USD "#,##0.00;[Red]"USD "#,##0.00', -1),
      { ok: true, format: '[>=100]"USD "#,##0.0;[Red]"USD "#,##0.0', decimalPlaces: 1 },
    );
    assert.deepEqual(
      transformNumberFormatPrecision('[$-409]$#,##0.00;[Red]-[$-409]$#,##0.00', -1),
      { ok: true, format: '[$-409]$#,##0.0;[Red]-[$-409]$#,##0.0', decimalPlaces: 1 },
    );
    assert.deepEqual(
      transformNumberFormatPrecision('0.00;[Red]-0.00;"zero";@', -1),
      { ok: true, format: '0.0;[Red]-0.0;"zero";@', decimalPlaces: 1 },
    );
  });

  it('fails closed for date, time, fraction, general, malformed, and boundary formats', () => {
    assert.equal(transformNumberFormatPrecision('yyyy-mm-dd', 1).ok, false);
    assert.equal(transformNumberFormatPrecision('h:mm:ss', -1).ok, false);
    assert.equal(transformNumberFormatPrecision('# ?/?', 1).ok, false);
    assert.equal(transformNumberFormatPrecision('general', 1).ok, false);
    assert.equal(transformNumberFormatPrecision('0.00"unterminated', 1).ok, false);
    assert.equal(transformNumberFormatPrecision('0.00', -1).ok, true);
    assert.equal(transformNumberFormatPrecision('0', -1).ok, false);
    assert.equal(transformNumberFormatPrecision(`0.${'0'.repeat(30)}`, 1).ok, false);
  });
});
