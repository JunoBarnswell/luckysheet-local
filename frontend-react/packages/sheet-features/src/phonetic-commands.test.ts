import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { registerSheetCommands } from './index';

test('phonetic guide is one canonical reversible cell mutation', () => {
  const workbook = new WorkbookModel('phonetic-command', 'Phonetic Command');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: '東京' });
  const metadata = { visible: true, type: 'hiragana' as const, alignment: 'center' as const, runs: [{ text: 'とうきょう', start: 0, end: 2 }] };
  runtime.execute('sheet.phonetic.set', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, metadata });
  assert.deepEqual(sheet.cells.get(0, 0)?.phonetic, metadata);
  assert.equal(runtime.undo(), true);
  assert.equal(sheet.cells.get(0, 0)?.phonetic, undefined);
  assert.equal(runtime.redo(), true);
  assert.deepEqual(sheet.cells.get(0, 0)?.phonetic, metadata);
});

test('phonetic guide rejects non-text targets before history or model changes', () => {
  const workbook = new WorkbookModel('phonetic-command-reject', 'Phonetic Command Reject');
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  const sheet = workbook.getSheet('sheet-1');
  sheet.cells.set(0, 0, { value: 42 });
  const before = workbook.snapshot();
  assert.throws(() => runtime.execute('sheet.phonetic.set', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, metadata: { visible: true, type: 'no-conversion', alignment: 'center', runs: [{ text: 'forty-two', start: 0, end: 2 }] } }), /PHONETIC_TEXT_REQUIRED/);
  assert.deepEqual(workbook.snapshot(), before);
  assert.equal(runtime.getHistoryDepth().undo, 0);
});
