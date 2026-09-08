import assert from 'node:assert/strict';
import test from 'node:test';
import { registerSheetCommands, formatTsv, parseTsv, isCellSetMutationParams } from './index';
import { openCanonicalTestRuntime, seedCanonicalCells } from '../../core-model/src/canonical-test-runtime.test';

const INPUT_CONTEXT = {
  sourceKind: 'clipboard-text' as const, cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900' as const,
  referenceDate: { year: 2026, month: 8, day: 27, hour: 0, minute: 0, second: 0, millisecond: 0 },
};

test('typed sheet commands plan through the canonical runtime and commit page values', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('sheet-features-command');
  try {
    registerSheetCommands(runtime);
    await runtime.execute('sheet.cell.set', { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'Title', style: { bold: true } } });
    const titlePageChecksum = workbook.manifest().pages[0]?.checksum;
    await runtime.execute('sheet.range.set', {
      sheetId: 'sheet-1', startRow: 1, startColumn: 0,
      values: [[{ value: 10 }, { value: 20 }], [{ value: 30 }, { value: 40 }]],
    });
    const sheet = workbook.getSheet('sheet-1');
    assert.equal(sheet.cells.get(0, 0)?.value, 'Title');
    assert.equal(sheet.cells.get(0, 0)?.style?.bold, true);
    assert.equal(sheet.cells.get(2, 1)?.value, 40);
    assert.equal(workbook.revision, 2);
    const rangePageChecksum = workbook.manifest().pages[0]?.checksum;
    assert.notEqual(rangePageChecksum, titlePageChecksum);
    assert.equal(runtime.getHistoryDepth().undo, 2);
    assert.equal(await runtime.undo(), true);
    assert.equal(workbook.manifest().pages[0]?.checksum, titlePageChecksum);
    assert.equal(await runtime.redo(), true);
    assert.equal(workbook.manifest().pages[0]?.checksum, rangePageChecksum);
  } finally {
    close();
  }
});

test('typed command rejection is fail-closed before a manifest revision is committed', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('sheet-features-rejection');
  try {
    registerSheetCommands(runtime);
    await assert.rejects(runtime.execute('sheet.cell.set', {
      sheetId: 'sheet-1', row: -1, column: 0, value: { value: 1 },
    }), /Invalid cell set parameters|invalid/i);
    assert.equal(workbook.revision, 0);
    assert.equal(runtime.getHistoryDepth().undo, 0);
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0), undefined);
  } finally {
    close();
  }
});

test('cell.set authority is canonical and independent of JSON key order', () => {
  const value = { value: 'ordered', style: { bold: true, italic: false } };
  const params = {
    sheetId: 'sheet-1', row: 0, column: 0, value,
    writeAuthority: { kind: 'script', target: { sheetId: 'sheet-1', row: 0, column: 0 }, validationDecision: { status: 'accepted' }, candidate: value },
  };
  assert.equal(isCellSetMutationParams({ ...params, value: { style: value.style, value: value.value } }), true);
});

test('clipboard serialization keeps typed cell values and formulas host-neutral', () => {
  const text = formatTsv([[{ value: 'A\tB' }, { value: 2 }, { value: null }]]);
  assert.deepEqual(parseTsv(text, INPUT_CONTEXT), [[{ value: 'A\tB' }, { value: 2 }, { value: null }]]);
});

test('canonical seed helper uses typed commands rather than mutable cell storage', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('sheet-features-seed');
  try {
    registerSheetCommands(runtime);
    await seedCanonicalCells(runtime, 'sheet-1', [{ row: 2, column: 2, value: 42 }]);
    assert.equal(workbook.getSheet('sheet-1').cells.get(2, 2)?.value, 42);
    assert.equal('set' in workbook.getSheet('sheet-1').cells, false);
  } finally {
    close();
  }
});
