import assert from 'node:assert/strict';
import test from 'node:test';
import { registerSheetCommands } from './index';
import { openCanonicalTestRuntime } from '../../core-model/src/canonical-test-runtime.test';

const metadata = { visible: true, type: 'hiragana' as const, alignment: 'center' as const, runs: [{ text: 'とうきょう', start: 0, end: 2 }] };
const target = (sheetId: string) => ({ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 });

test('phonetic guide commits through one canonical reversible cell mutation', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('phonetic-command');
  try {
    registerSheetCommands(runtime);
    await runtime.execute('sheet.cell.set', { sheetId: 'sheet-1', row: 0, column: 0, value: { value: '東京' } });
    await runtime.execute('sheet.phonetic.set', { sheetId: 'sheet-1', range: target('sheet-1'), metadata });
    assert.deepEqual(workbook.getSheet('sheet-1').cells.get(0, 0)?.phonetic, metadata);
    assert.equal(await runtime.undo(), true);
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0)?.phonetic, undefined);
    assert.equal(await runtime.redo(), true);
    assert.deepEqual(workbook.getSheet('sheet-1').cells.get(0, 0)?.phonetic, metadata);
  } finally {
    close();
  }
});

test('phonetic guide rejects non-text targets before a commit or history entry', async () => {
  const { workbook, runtime, close } = await openCanonicalTestRuntime('phonetic-command-reject');
  try {
    registerSheetCommands(runtime);
    await runtime.execute('sheet.cell.set', { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 42 } });
    const revision = workbook.revision;
    await assert.rejects(runtime.execute('sheet.phonetic.set', { sheetId: 'sheet-1', range: target('sheet-1'), metadata }), /PHONETIC_TEXT_REQUIRED/);
    assert.equal(workbook.revision, revision);
    assert.equal(runtime.getHistoryDepth().undo, 1);
  } finally {
    close();
  }
});
