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

describe('print layout', () => {
  it('builds print snapshots from workbook content', () => {
    const workbook = new WorkbookModel('wb-print', 'Print');
    const sheetId = workbook.activeSheetId;
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
    const sheetId = workbook.activeSheetId;
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
    const sheetId = workbook.activeSheetId;
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
});
