import test from 'node:test';
import assert from 'node:assert/strict';
import { protectionResolver, type ProtectionRule } from './index';

const sheetId = 'sheet-1';
const cell = (row: number, column: number) => ({ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column });

test('ProtectionResolver uses cell locked style under active sheet protection', () => {
  const rules: ProtectionRule[] = [{ id: 'sheet-lock', scope: 'sheet', sheetId, locked: true, allow: {} }];
  const unlocked = protectionResolver.resolve({
    sheetId, rules, ranges: [cell(0, 0)], action: 'edit-cell', rowCount: 10, columnCount: 10,
    readCellStyle: () => ({ locked: false }),
  });
  assert.equal(unlocked.allowed, true);

  const defaultLocked = protectionResolver.resolve({
    sheetId, rules, ranges: [cell(0, 1)], action: 'edit-cell', rowCount: 10, columnCount: 10,
  });
  assert.equal(defaultLocked.allowed, false);
  assert.match(defaultLocked.reason ?? '', /locked/);
});

test('ProtectionResolver applies native allow flags to operation actions', () => {
  const allowed: ProtectionRule = {
    id: 'sheet-lock', scope: 'sheet', sheetId, locked: true,
    allow: { formatCells: true, sort: true },
  };
  assert.equal(protectionResolver.resolve({ sheetId, rules: [allowed], ranges: [cell(0, 0)], action: 'format', rowCount: 10, columnCount: 10 }).allowed, true);
  assert.equal(protectionResolver.resolve({ sheetId, rules: [allowed], ranges: [cell(0, 0)], action: 'sort', rowCount: 10, columnCount: 10 }).allowed, true);
  assert.equal(protectionResolver.resolve({ sheetId, rules: [allowed], ranges: [cell(0, 0)], action: 'auto-filter', rowCount: 10, columnCount: 10 }).allowed, false);
});

test('ProtectionResolver rejects a mixed locked range before any write', () => {
  const rules: ProtectionRule[] = [{ id: 'sheet-lock', scope: 'sheet', sheetId, locked: true, allow: {} }];
  const result = protectionResolver.resolve({
    sheetId, rules, ranges: [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }], action: 'edit-cell', rowCount: 10, columnCount: 10,
    readCellStyle: (row, column) => ({ locked: column === 1 }),
  });
  assert.equal(result.allowed, false);
  assert.equal(result.lockedCells, 1);
  assert.equal(result.unlockedCells, 1);
});
