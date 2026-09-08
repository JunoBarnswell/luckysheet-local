import assert from 'node:assert/strict';
import test from 'node:test';
import { registerSheetCommands } from './index';
import { parseReplacementValue } from './home-commands';
import { openCanonicalTestRuntime } from '../../core-model/src/canonical-test-runtime.test';

const context = {
  sourceKind: 'find-replace' as const, cultureId: 'en-US', decimalSeparator: '.', groupSeparator: ',', dateSystem: '1900' as const,
  referenceDate: { year: 2026, month: 8, day: 27, hour: 0, minute: 0, second: 0, millisecond: 0 },
};

test('AutoSum resolves canonical values and publishes one typed command transaction', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('home-autosum');
  try {
    registerSheetCommands(runtime);
    const sheetId = workbook.primarySheetId;
    await runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 10 } });
    await runtime.execute('sheet.cell.set', { sheetId, row: 1, column: 0, value: { value: 20 } });
    runtime.setCellValueResolver((_sheet, row, column) => column === 0 && row < 2 ? (row + 1) * 10 : undefined);
    const before = workbook.revision;
    await runtime.execute('formula.autosum', { sheetId, range: { sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }, target: { row: 2, column: 0 } });
    assert.equal(workbook.revision, before + 1);
    assert.equal(workbook.getSheet(sheetId).cells.get(2, 0)?.formula, '=SUM(A1:A2)');
    assert.equal(runtime.getHistoryDepth().undo, 3);
  } finally {
    close();
  }
});

test('AutoSum rejects an unsafe target before changing the canonical manifest', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('home-autosum-reject');
  try {
    registerSheetCommands(runtime);
    const sheetId = workbook.primarySheetId;
    await runtime.execute('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 10 } });
    const revision = workbook.revision;
    await assert.rejects(runtime.execute('formula.autosum', { sheetId, range: { sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }, target: { row: 1, column: 0 } }), /not blank|include its target/i);
    assert.equal(workbook.revision, revision);
  } finally {
    close();
  }
});

test('replacement values preserve the typed command contract', () => {
  assert.deepEqual(parseReplacementValue('0', context), { kind: 'number', value: 0 });
  assert.deepEqual(parseReplacementValue('=A1+1', context), { kind: 'formula', value: null, formula: '=A1+1' });
  assert.equal(parseReplacementValue('', context).kind, 'empty');
});
