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
    assert.equal(snapshot.showPrintPreview, true);
    assert.ok(snapshot.printPageCount >= 1);
    assert.ok(snapshot.printPages.length >= 1);
    assert.ok(snapshot.printArea);
    assert.match(snapshot.notice, /page\(s\)/i);
  });

  it('runs print.export through command path', () => {
    const app = new WorkbookSession();
    app.exportPdf({
      paper: 'Letter',
      orientation: 'landscape',
      margin: { top: 10, right: 10, bottom: 10, left: 10 },
    });
    const printSnapshot = app.getPrintSnapshot();
    assert.ok(printSnapshot);
    assert.equal(printSnapshot?.layout.paper, 'Letter');
    assert.equal(printSnapshot?.layout.orientation, 'landscape');
  });

  it('updates print area through print.area.set', () => {
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
    assert.equal(app.getUiSnapshot().showPrintPreview, true);
    assert.ok(app.getUiSnapshot().printPageCount >= 1);
  });
});
