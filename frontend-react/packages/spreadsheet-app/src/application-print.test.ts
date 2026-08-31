import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession print integration', () => {
  it('exposes print metadata in ui snapshot after preview', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'Print me' },
    });
    app.printWorkbook({
      paper: 'A4',
      orientation: 'portrait',
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
    });
    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.dialogs.active, 'print-preview');
    assert.ok(snapshot.printPageCount >= 1);
    assert.ok(snapshot.printPages.length >= 1);
    assert.ok(snapshot.printArea);
    assert.match(snapshot.notice, /page\(s\)/i);
  });

  it('runs print.export through the real page projection path', async () => {
    const app = new WorkbookSession();
    await app.exportPdf({
      paper: 'Letter',
      orientation: 'landscape',
      margin: { top: 10, right: 10, bottom: 10, left: 10 },
    });
    const printSnapshot = app.getPrintSnapshot();
    assert.ok(printSnapshot);
    assert.equal(printSnapshot?.layout.paper, 'Letter');
    assert.equal(printSnapshot?.layout.orientation, 'landscape');
  });

  it('keeps saved print area, current selection, and active-sheet scopes distinct', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    sheet.cells.set(0, 0, { value: 'used-start' });
    sheet.cells.set(20, 5, { value: 'used-end' });
    app.setPrintArea({ sheetId, startRow: 2, endRow: 6, startColumn: 1, endColumn: 3 });
    app.selectRange({ startRow: 10, endRow: 12, startColumn: 2, endColumn: 4 });
    const layout = { paper: 'A4' as const, orientation: 'portrait' as const, margin: { top: 20, right: 20, bottom: 20, left: 20 } };
    app.printWorkbook(layout, 'saved-area');
    assert.deepEqual(app.getPrintSnapshot()?.printArea, { sheetId, startRow: 2, endRow: 6, startColumn: 1, endColumn: 3 });
    app.printWorkbook(layout, 'selection');
    assert.deepEqual(app.getPrintSnapshot()?.printArea, { sheetId, startRow: 10, endRow: 12, startColumn: 2, endColumn: 4 });
    app.printWorkbook(layout, 'active-sheet');
    assert.deepEqual(app.getPrintSnapshot()?.printArea, sheet.usedRange);
  });

  it('updates print area through pageLayout.printArea.set', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.setPrintArea({
      sheetId,
      startRow: 1,
      endRow: 5,
      startColumn: 0,
      endColumn: 3,
    });
    const area = app.getPrintSnapshot()?.printArea;
    assert.equal(area?.startRow, 1);
    assert.equal(area?.endRow, 5);
  });

  it('allows viewers to preview print output', () => {
    const app = new WorkbookSession();
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);
    app.printWorkbook({
      paper: 'A4',
      orientation: 'portrait',
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
    });
    assert.equal(app.getUiSnapshot().dialogs.active, 'print-preview');
    assert.ok(app.getUiSnapshot().printPageCount >= 1);
  });

  it('exposes persisted page-layout fields without replacing the saved print area', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.setPrintArea({ sheetId, startRow: 2, endRow: 8, startColumn: 1, endColumn: 4 });
    app.updatePrintPageSetup({
      paper: 'Letter',
      orientation: 'landscape',
      margin: { top: 12, right: 12, bottom: 12, left: 12 },
    });
    app.setPrintTitles('rows');
    app.setPrintPageBreak({ row: 5 });
    app.setPrintScale(80, 1, null);
    app.setPrintGridlines(true);
    app.setPrintHeadings(true);
    app.setViewGridlines(false);
    app.setViewHeadings(false);

    const snapshot = app.getPrintSnapshot();
    assert.equal(snapshot?.printArea.startRow, 2);
    assert.equal(snapshot?.printArea.endRow, 8);
    assert.equal(app.getPrintPageSetup().scale, 80);
    assert.equal(app.getPrintPageSetup().fitToWidth, 1);
    assert.equal(app.getPrintPageSetup().printGridlines, true);
    assert.equal(app.getPrintPageSetup().printHeadings, true);
    assert.equal(app['runtime'].model.getSheet(sheetId).showGridlines, false);
    assert.equal(app['runtime'].model.getSheet(sheetId).showHeaders, false);

    app.clearPrintPageBreaks();
    app.clearPrintTitles();
    assert.equal(app.getPrintSnapshot()?.model.pageBreaks.length, 0);
    assert.equal(app.getPrintSnapshot()?.model.repeatRows, undefined);
  });
});
