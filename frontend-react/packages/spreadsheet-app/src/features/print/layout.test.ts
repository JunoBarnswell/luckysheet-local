import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  buildPrintLayoutModel,
  buildPrintSnapshot,
  printLayoutToPageSetup,
  resolvePrintArea,
  summarizePrintSnapshot,
} from './layout';
import { computePrintPages, createDefaultPrintLayout } from './index';

describe('print layout', () => {
  it('builds print snapshots from workbook content', () => {
    const workbook = new WorkbookModel('wb-print', 'Print');
    const sheetId = workbook.primarySheetId;
    const sheet = workbook.getSheet(sheetId);
    sheet.cells.set(0, 0, { value: 'Header' });
    sheet.cells.set(1, 0, { value: 42 });
    sheet.cells.set(1, 1, { value: 'East' });

    const layout = {
      paper: 'A4' as const,
      orientation: 'portrait' as const,
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
    };
    const snapshot = buildPrintSnapshot(workbook, sheetId, layout);
    assert.ok(snapshot.pageCount >= 1);
    assert.equal(snapshot.printArea.startRow, 0);
    assert.equal(snapshot.printArea.endRow, 1);
    assert.match(summarizePrintSnapshot(snapshot), /page\(s\)/i);
  });

  it('uses multi-cell selection as print area', () => {
    const workbook = new WorkbookModel('wb-print-selection', 'Selection');
    const sheetId = workbook.primarySheetId;
    const sheet = workbook.getSheet(sheetId);
    sheet.cells.set(0, 0, { value: 'A' });
    sheet.cells.set(9, 4, { value: 'B' });

    const selection = {
      sheetId,
      startRow: 2,
      endRow: 4,
      startColumn: 1,
      endColumn: 2,
    };
    const snapshot = buildPrintSnapshot(workbook, sheetId, {
      paper: 'Letter',
      orientation: 'landscape',
      margin: { top: 10, right: 10, bottom: 10, left: 10 },
    }, selection);

    assert.equal(snapshot.printArea.startRow, 2);
    assert.equal(snapshot.printArea.endColumn, 2);
  });

  it('maps ui layout to page setup', () => {
    const setup = printLayoutToPageSetup({
      paper: 'Legal',
      orientation: 'landscape',
      margin: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
      scale: 90,
      fitToWidth: true,
    });
    assert.equal(setup.paperSize, 'legal');
    assert.equal(setup.orientation, 'landscape');
    assert.equal(setup.scale, 90);
    assert.equal(setup.fitToWidth, 1);
  });

  it('resolves print area from used range for single-cell selection', () => {
    const workbook = new WorkbookModel('wb-used', 'Used');
    const sheetId = workbook.primarySheetId;
    const sheet = workbook.getSheet(sheetId);
    sheet.cells.set(3, 2, { value: 'Only' });

    const area = resolvePrintArea(sheet, {
      sheetId,
      startRow: 0,
      endRow: 0,
      startColumn: 0,
      endColumn: 0,
    });
    assert.equal(area.startRow, 3);
    assert.equal(area.startColumn, 2);
  });

  it('builds print layout model with repeat rows', () => {
    const model = buildPrintLayoutModel('wb-1', 'sheet-1', {
      paper: 'A4',
      orientation: 'portrait',
      margin: { top: 20, right: 20, bottom: 20, left: 20 },
      repeatRows: {
        sheetId: 'sheet-1',
        startRow: 0,
        endRow: 0,
        startColumn: 0,
        endColumn: 5,
      },
    }, {
      sheetId: 'sheet-1',
      startRow: 0,
      endRow: 10,
      startColumn: 0,
      endColumn: 5,
    });
    assert.equal(model.repeatRows?.start, 0);
    assert.equal(model.printAreas[0]?.range.endRow, 10);
  });

  it('honors explicit row and column page breaks with one global page sequence', () => {
    const model = createDefaultPrintLayout('wb-breaks', 'sheet-1');
    model.printAreas = [{ sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 0, endRow: 9, startColumn: 0, endColumn: 5 } }];
    model.pageSetup.margins = { top: 12, right: 12, bottom: 12, left: 12, header: 0, footer: 0 };
    model.pageBreaks = [{ sheetId: 'sheet-1', row: 5 }, { sheetId: 'sheet-1', column: 3 }];
    const pages = computePrintPages(model, 20, 80);
    assert.equal(pages.length, 4);
    assert.deepEqual(pages.map((page) => page.pageIndex), [0, 1, 2, 3]);
    assert.equal(pages[0]?.range.endRow, 4);
    assert.equal(pages[2]?.range.startRow, 5);
    assert.equal(pages[1]?.range.startColumn, 3);
  });

  it('projects persisted print titles when preview uses the stored document', () => {
    const workbook = new WorkbookModel('wb-print-stored-titles', 'Stored titles');
    const sheetId = workbook.primarySheetId;
    const sheet = workbook.getSheet(sheetId);
    sheet.cells.set(0, 0, { value: 'Header' });
    sheet.cells.set(8, 3, { value: 'Tail' });
    workbook.setPrintDocument({
      schema: 'PrintDocument',
      unitId: workbook.unitId,
      sheetId,
      pageSetup: {
        paperSize: 'a4',
        orientation: 'portrait',
        margins: { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36 },
        scale: 100,
        printGridlines: false,
        printHeadings: false,
        centerHorizontally: false,
        centerVertically: false,
      },
      printAreas: [],
      pageBreaks: [],
      repeatRows: { start: 0, end: 1 },
      repeatColumns: { start: 0, end: 0 },
    });

    const snapshot = buildPrintSnapshot(workbook, sheetId);
    assert.deepEqual(snapshot.model.repeatRows, { start: 0, end: 1 });
    assert.deepEqual(snapshot.model.repeatColumns, { start: 0, end: 0 });
  });
});
