import test from 'node:test';
import assert from 'node:assert/strict';
import { dateToSerial, serialToExcelDate, parseDateSystem } from './date-system';

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
