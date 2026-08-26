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
    const sheetId = workbook.primarySheetId;
    const range = { sheetId, startRow: 2, endRow: 20, startColumn: 1, endColumn: 7 };

    commands.execute('pageLayout.pageSetup.set', { sheetId, pageSetup });
    commands.execute('pageLayout.printArea.set', { sheetId, range });
    commands.execute('pageLayout.pageBreak.insert', { sheetId, pageBreak: { sheetId, row: 10 } });

    const document = getPrintDocument(workbook, sheetId);
    assert.deepEqual(document.pageSetup, pageSetup);
    assert.deepEqual(document.printAreas, [{ sheetId, range }]);
    assert.deepEqual(document.pageBreaks, [{ sheetId, row: 10 }]);
    assert.equal(commands.getHistoryDepth().undo, 3);
  });

  it('undoes and redoes document changes without creating a second write path', () => {
    const workbook = new WorkbookModel('wb-print-history', 'Print');
    const commands = runtime(workbook);
    const sheetId = workbook.primarySheetId;
    const range = { sheetId, startRow: 0, endRow: 5, startColumn: 0, endColumn: 2 };

    commands.execute('pageLayout.printArea.set', { sheetId, range });
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
    const sheetId = source.primarySheetId;
    sourceRuntime.execute('pageLayout.pageBreak.insert', { sheetId, pageBreak: { sheetId, column: 4 } });
    const mutation = sourceRuntime.getUndoEntries().at(-1)?.redo[0];
    assert.ok(mutation);
    targetRuntime.applyRemoteMutations([mutation]);
    assert.deepEqual(getPrintDocument(target, sheetId).pageBreaks, [{ sheetId, column: 4 }]);
  });

  it('persists print titles, scale, gridlines and headings through explicit commands', () => {
    const workbook = new WorkbookModel('wb-print-layout-fields', 'Print fields');
    const commands = runtime(workbook);
    const sheetId = workbook.primarySheetId;

    commands.execute('pageLayout.printTitles.set', {
      sheetId,
      repeatRows: { start: 0, end: 1 },
      repeatColumns: { start: 0, end: 0 },
    });
    commands.execute('pageLayout.scaleToFit.set', { sheetId, scale: 80, fitToWidth: 1 });
    commands.execute('pageLayout.printGridlines.set', { sheetId, enabled: true });
    commands.execute('pageLayout.printHeadings.set', { sheetId, enabled: true });

    const document = getPrintDocument(workbook, sheetId);
    assert.deepEqual(document.repeatRows, { start: 0, end: 1 });
    assert.deepEqual(document.repeatColumns, { start: 0, end: 0 });
    assert.equal(document.pageSetup.scale, 80);
    assert.equal(document.pageSetup.fitToWidth, 1);
    assert.equal(document.pageSetup.printGridlines, true);
    assert.equal(document.pageSetup.printHeadings, true);

    assert.equal(commands.undo(), true);
    assert.equal(getPrintDocument(workbook, sheetId).pageSetup.printHeadings, false);
    assert.equal(commands.redo(), true);
    assert.equal(getPrintDocument(workbook, sheetId).pageSetup.printHeadings, true);
    commands.execute('pageLayout.printTitles.clear', { sheetId });
    assert.equal(getPrintDocument(workbook, sheetId).repeatRows, undefined);
    assert.equal(getPrintDocument(workbook, sheetId).repeatColumns, undefined);
  });

  it('persists page-layout gridline and heading view options through undoable mutations', () => {
    const workbook = new WorkbookModel('wb-page-layout-view', 'Page Layout view');
    const commands = runtime(workbook);
    const sheetId = workbook.primarySheetId;
    const sheet = workbook.getSheet(sheetId);

    commands.execute('pageLayout.viewGridlines.set', { sheetId, enabled: false });
    commands.execute('pageLayout.viewHeadings.set', { sheetId, enabled: false });
    assert.equal(sheet.showGridlines, false);
    assert.equal(sheet.showHeaders, false);
    assert.equal(commands.undo(), true);
    assert.equal(sheet.showHeaders, true);
    assert.equal(commands.undo(), true);
    assert.equal(sheet.showGridlines, true);
    assert.equal(commands.redo(), true);
    assert.equal(sheet.showGridlines, false);
  });
});
