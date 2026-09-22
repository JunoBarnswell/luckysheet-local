import test from 'node:test';
import assert from 'node:assert/strict';
import { canonicalDateToSerial, canonicalExcelDateFromSerial, dateToSerial, isExcelDateFormat, parseDateSystem, serialToCanonicalDate, serialToExcelDate } from './date-system';

test('1900 date system serial round-trip', () => {
  const date = new Date(Date.UTC(2024, 0, 1));
  const serial = dateToSerial(date, '1900');
  const back = serialToExcelDate(serial, '1900');
  assert.equal(back.getUTCFullYear(), 2024);
  assert.equal(back.getUTCMonth(), 0);
  assert.equal(back.getUTCDate(), 1);
});

test('parseDateSystem detects 1904', () => {
  assert.equal(parseDateSystem('<workbookPr date1904="1"/>'), '1904');
  assert.equal(parseDateSystem('<workbookPr/>'), '1900');
});

test('1900 system preserves Excel leap-year gap and fractional time', () => {
  assert.equal(serialToCanonicalDate(59, '1900'), '1900-02-28T00:00:00.000Z');
  assert.equal(serialToCanonicalDate(61, '1900'), '1900-03-01T00:00:00.000Z');
  assert.equal(serialToCanonicalDate(61.5, '1900'), '1900-03-01T12:00:00.000Z');
  assert.throws(() => serialToExcelDate(60, '1900'), /non-existent/);
});

test('1904 system maps serial zero to its epoch', () => {
  assert.equal(serialToCanonicalDate(0, '1904'), '1904-01-01T00:00:00.000Z');
  assert.equal(canonicalDateToSerial('1904-01-02T06:00:00.000Z', '1904'), 1.25);
});

test('date system and date format parsing fail closed', () => {
  assert.equal(parseDateSystem('<x:workbookPr date1904="true"/>'), '1904');
  assert.throws(() => parseDateSystem('<workbookPr date1904="maybe"/>'), /Invalid Excel date1904/);
  assert.equal(isExcelDateFormat('yyyy-mm-dd'), true);
  assert.equal(isExcelDateFormat('0.00'), false);
  assert.throws(() => canonicalDateToSerial('2024-01-01', '1900'), /UTC ISO/);
});

test('canonical serial parts are timezone-independent for both workbook calendars', () => {
  assert.deepEqual(canonicalExcelDateFromSerial(59, '1900'), { system: '1900', serial: 59, year: 1900, month: 2, day: 28, hour: 0, minute: 0, second: 0, millisecond: 0 });
  assert.deepEqual(canonicalExcelDateFromSerial(61.5, '1900'), { system: '1900', serial: 61.5, year: 1900, month: 3, day: 1, hour: 12, minute: 0, second: 0, millisecond: 0 });
  assert.deepEqual(canonicalExcelDateFromSerial(0.25, '1904'), { system: '1904', serial: 0.25, year: 1904, month: 1, day: 1, hour: 6, minute: 0, second: 0, millisecond: 0 });
  assert.throws(() => canonicalExcelDateFromSerial(60, '1900'), /non-existent/);
  assert.throws(() => canonicalExcelDateFromSerial(Number.NaN, '1900'), /finite/);
});
