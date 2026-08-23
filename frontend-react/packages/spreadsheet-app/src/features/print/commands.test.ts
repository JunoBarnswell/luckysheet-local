import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRegistry, CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  getPrintDocument,
  registerPrintCommands,
  type PageSetup,
} from './index';

const pageSetup: PageSetup = {
  paperSize: 'letter',
  orientation: 'landscape',
  margins: { top: 36, right: 36, bottom: 36, left: 36, header: 18, footer: 18 },
  scale: 90,
  fitToWidth: 1,
  printGridlines: true,
  printHeadings: false,
  centerHorizontally: true,
  centerVertically: false,
};

function runtime(workbook: WorkbookModel): CommandRuntime {
  const commands = new CommandRuntime(workbook, new CommandRegistry({ requireMutationMetadata: true }));
  registerPrintCommands(commands.registry);
  commands.registry.assertComplete();
  return commands;
}

describe('print document commands', () => {
  it('stores page setup, area and page breaks as one canonical document', () => {
    const workbook = new WorkbookModel('wb-print-command', 'Print');
    const commands = runtime(workbook);
    const sheetId = workbook.activeSheetId;
    const range = { sheetId, startRow: 2, endRow: 20, startColumn: 1, endColumn: 7 };

    commands.execute('print.pageSetup', { sheetId, pageSetup });
    commands.execute('print.area.set', { sheetId, range });
    commands.execute('print.pageBreak.set', { sheetId, pageBreak: { sheetId, row: 10 } });

    const document = getPrintDocument(workbook, sheetId);
    assert.deepEqual(document.pageSetup, pageSetup);
    assert.deepEqual(document.printAreas, [{ sheetId, range }]);
    assert.deepEqual(document.pageBreaks, [{ sheetId, row: 10 }]);
    assert.equal(commands.getHistoryDepth().undo, 3);
  });

  it('undoes and redoes document changes without creating a second write path', () => {
    const workbook = new WorkbookModel('wb-print-history', 'Print');
    const commands = runtime(workbook);
    const sheetId = workbook.activeSheetId;
    const range = { sheetId, startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 };

    commands.execute('print.area.set', { sheetId, range });
    assert.equal(getPrintDocument(workbook, sheetId).printAreas.length, 1);
    assert.equal(commands.undo(), true);
    assert.equal(getPrintDocument(workbook, sheetId).printAreas.length, 0);
    assert.equal(commands.redo(), true);
    assert.deepEqual(getPrintDocument(workbook, sheetId).printAreas[0]?.range, range);
  });

  it('replays the exact typed document mutation on another workbook', () => {
    const source = new WorkbookModel('wb-print-remote', 'Print');
    const target = new WorkbookModel('wb-print-remote', 'Print');
    const sourceRuntime = runtime(source);
    const targetRuntime = runtime(target);
    const sheetId = source.activeSheetId;
    sourceRuntime.execute('print.pageBreak.set', { sheetId, pageBreak: { sheetId, column: 4 } });
    const mutation = sourceRuntime.getUndoEntries().at(-1)?.redo[0];
    assert.ok(mutation);
    targetRuntime.applyRemoteMutations([mutation]);
    assert.deepEqual(getPrintDocument(target, sheetId).pageBreaks, [{ sheetId, column: 4 }]);
  });
});
