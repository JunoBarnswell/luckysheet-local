import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { copyRangeToClipboardData, createPasteSpecialSpec } from '@react-sheets/sheet-features';
import { WorkbookSession } from './workbook-session';

function selectCell(app: WorkbookSession, row: number, column: number): void {
  const sheetId = app.getActiveSheetId();
  app.runCommand('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }],
    primaryRangeIndex: 0,
    activeCell: { row, column },
    anchorCell: { row, column },
  });
}

function selectRange(app: WorkbookSession, startRow: number, endRow: number, startColumn = 0, endColumn = 0): void {
  const sheetId = app.getActiveSheetId();
  app.runCommand('selection.set', {
    sheetId,
    ranges: [{ sheetId, startRow, endRow, startColumn, endColumn }],
    primaryRangeIndex: 0,
    activeCell: { row: startRow, column: startColumn },
    anchorCell: { row: startRow, column: startColumn },
  });
}

describe('WorkbookSession core editing integration', () => {
  it('Fill Series keeps the complete selected seed range in the canonical planner', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    sheet.cells.set(0, 0, { value: 1 });
    sheet.cells.set(1, 0, { value: 3 });
    selectRange(app, 0, 4);

    app.fillSelection('down', 'series');
    assert.deepEqual([2, 3, 4].map((row) => sheet.cells.get(row, 0)?.value), [5, 7, 9]);
    assert.deepEqual([sheet.cells.get(0, 0)?.value, sheet.cells.get(1, 0)?.value], [1, 3]);

    app.undo();
    assert.equal(sheet.cells.get(2, 0), undefined);
    assert.deepEqual([sheet.cells.get(0, 0)?.value, sheet.cells.get(1, 0)?.value], [1, 3]);
    app.redo();
    assert.deepEqual([2, 3, 4].map((row) => sheet.cells.get(row, 0)?.value), [5, 7, 9]);
  });

  it('preserves AutoFit dimension results through undo and redo', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    const defaultWidth = sheet.defaultColumnWidthPx;
    const defaultHeight = sheet.defaultRowHeightPx;

    app.applyColumnWidths([{ column: 0, widthPx: 140 }, { column: 1, widthPx: 160 }]);
    app.applyRowHeights([{ row: 0, heightPx: 36 }, { row: 1, heightPx: 42 }]);
    assert.equal(sheet.columnWidthsPx[0], 140);
    assert.equal(sheet.columnWidthsPx[1], 160);
    assert.equal(sheet.rowHeightsPx[0], 36);
    assert.equal(sheet.rowHeightsPx[1], 42);

    app.undo();
    assert.equal(sheet.rowHeightsPx[0], defaultHeight);
    assert.equal(sheet.rowHeightsPx[1], defaultHeight);
    assert.equal(sheet.columnWidthsPx[0], 140);
    assert.equal(sheet.columnWidthsPx[1], 160);
    app.undo();
    assert.equal(sheet.columnWidthsPx[0], defaultWidth);
    assert.equal(sheet.columnWidthsPx[1], defaultWidth);
    assert.equal(sheet.defaultColumnWidthPx, defaultWidth);
    assert.equal(sheet.defaultRowHeightPx, defaultHeight);

    app.redo();
    app.redo();
    assert.equal(sheet.columnWidthsPx[0], 140);
    assert.equal(sheet.columnWidthsPx[1], 160);
    assert.equal(sheet.rowHeightsPx[0], 36);
    assert.equal(sheet.rowHeightsPx[1], 42);
  });

  it('selectAddress jumps to the requested cell', () => {
    const app = new WorkbookSession();
    assert.equal(app.selectAddress('C3'), true);
    assert.equal(app.getUiSnapshot().activeCell, 'C3');
  });

  it('paste special values copies values without formulas', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 42, formula: '=21*2' },
    });
    selectCell(app, 0, 0);
    const range = app.getPrimaryRange();
    app.setClipboard({ ...copyRangeToClipboardData(app['runtime'].model, range), transfer: 'copy' });
    selectCell(app, 0, 1);
    const outcome = await app.pasteSpecial(createPasteSpecialSpec({ content: 'values', formatting: 'none', metadata: { commentsNotes: false, validation: false, columnWidths: false, conditionalFormats: false, hyperlinks: false } }));
    assert.equal(outcome.status, 'committed');

    const target = app['runtime'].model.getSheet(sheetId).cells.get(0, 1);
    assert.equal(target?.value, 42);
    assert.equal(target?.formula, undefined);
  });

  it('returns a typed dispatch outcome and preserves state when permission rejects a paste', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    selectCell(app, 0, 0);
    app.setClipboard({
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      values: [[{ value: 'blocked' }]],
      transfer: 'move',
      rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
    });
    const historyDepth = app['runtime'].commands.getHistoryDepth();
    app['permission'].applyServerAccess('viewer');
    app['permission'].setOnline(true);

    const outcome = await app.paste();

    assert.equal(outcome.status, 'rejected');
    if (outcome.status === 'rejected') assert.equal(outcome.error.code, 'COMMAND_REJECTED');
    assert.equal(app.getClipboard()?.transfer, 'move');
    assert.deepEqual(app['runtime'].commands.getHistoryDepth(), historyDepth);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0), undefined);
    assert.match(app.getUiSnapshot().notice, /permission|viewer|edit/i);
  });

  it('keeps paste pending until data-region materialization rejects and then preserves all state', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app['runtime'].model.getSheet(sheetId).dataRegions.push({
      id: 'missing-source-region',
      sourceId: 'missing-source',
      range: { sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      headerRow: 0,
      revision: 0,
    });
    selectCell(app, 0, 0);
    app.setClipboard({
      range: { sheetId, startRow: 1, endRow: 1, startColumn: 0, endColumn: 0 },
      values: [[{ value: 'materialize-me' }]],
      transfer: 'copy',
      rangeMetadata: { columnWidths: [], validations: [], conditionalFormats: [], notes: [], comments: [], hyperlinks: [] },
    });
    const historyDepth = app['runtime'].commands.getHistoryDepth();

    const pending = app.paste();
    assert.equal(app.getUiSnapshot().pendingCommandCount, 1);
    const outcome = await pending;

    assert.equal(outcome.status, 'rejected');
    if (outcome.status === 'rejected') assert.equal(outcome.error.code, 'MATERIALIZATION_FAILED');
    assert.equal(app.getUiSnapshot().pendingCommandCount, 0);
    assert.deepEqual(app['runtime'].commands.getHistoryDepth(), historyDepth);
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0), undefined);
    assert.equal(app.getClipboard()?.values[0]?.[0]?.value, 'materialize-me');
  });

  it('keeps the private clipboard usable when external clipboard publication is unavailable', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'private-copy' } });
    selectCell(app, 0, 0);
    const historyDepth = app['runtime'].commands.getHistoryDepth();

    const outcome = await app.copy();

    assert.equal(outcome.status, 'failed');
    assert.equal(outcome.privatePayloadStored, true);
    assert.equal(app.getClipboard()?.values[0]?.[0]?.value, 'private-copy');
    assert.equal(app.getUiSnapshot().clipboard.systemStatus, 'failed');
    assert.deepEqual(app['runtime'].commands.getHistoryDepth(), historyDepth);
    assert.match(app.getUiSnapshot().notice, /external clipboard|unavailable|failed/i);
  });

  it('formatCells applies number format through sheet.format.set', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 1,
      column: 1,
      value: { value: 1234.5 },
    });
    selectCell(app, 1, 1);
    app.formatCells({ numberFormat: '#,##0.00' });

    const cell = app['runtime'].model.getSheet(sheetId).cells.get(1, 1);
    assert.equal(cell?.style?.numberFormat, '#,##0.00');
  });

  it('font family formatting canonicalizes listed and imported names through history', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    selectCell(app, 1, 1);
    app.runCommand('sheet.style.set', {
      sheetId,
      range: { sheetId, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      style: { fontFamily: '  aRiAl  ' },
    });
    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(1, 1)?.style?.fontFamily, 'Arial');

    app.undo();
    assert.equal(sheet.cells.get(1, 1)?.style?.fontFamily, undefined);
    app.redo();
    assert.equal(sheet.cells.get(1, 1)?.style?.fontFamily, 'Arial');

    app.runCommand('sheet.style.set', {
      sheetId,
      range: { sheetId, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      style: { fontFamily: '  Imported Local Font  ' },
    });
    assert.equal(sheet.cells.get(1, 1)?.style?.fontFamily, 'Imported Local Font');

    const historyDepth = app['runtime'].commands.getHistoryDepth();
    assert.throws(() => app.runCommand('sheet.style.set', {
      sheetId,
      range: { sheetId, startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      style: { fontFamily: '   ' },
    }), /must not be empty/);
    assert.deepEqual(app['runtime'].commands.getHistoryDepth(), historyDepth);
    assert.equal(sheet.cells.get(1, 1)?.style?.fontFamily, 'Imported Local Font');
  });

  it('font family selection exposes a mixed state without inventing a value', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    const sheet = app['runtime'].model.getSheet(sheetId);
    sheet.cells.set(0, 0, { value: 'left', style: { fontFamily: 'Arial' } });
    sheet.cells.set(0, 1, { value: 'right', style: { fontFamily: 'Calibri' } });
    selectRange(app, 0, 0, 0, 1);

    const home = app.getUiSnapshot().homeRibbon;
    assert.equal(home.style.fontFamily, undefined);
    assert.equal(home.mixedStyleKeys.includes('fontFamily'), true);
  });

  it('supplies active sheet and selection to a catalog-style sheet.style.set descriptor', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    selectCell(app, 3, 2);
    app.dispatch({ commandId: 'sheet.style.set', params: { style: { bold: true } } });
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(3, 2)?.style?.bold, true);
  });

  it('freezeAtPrimary updates worksheet freeze panes', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    selectCell(app, 2, 1);
    app.freezeAtPrimary();

    const pane = app['runtime'].model.getSheet(sheetId).pane;
    assert.equal(pane.kind === 'frozen' ? pane.ySplit : 0, 2);
    assert.equal(pane.kind === 'frozen' ? pane.xSplit : 0, 1);
  });

  it('duplicateSheet creates a copy and switches to it', () => {
    const app = new WorkbookSession();
    const sourceId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId: sourceId,
      row: 0,
      column: 0,
      value: { value: 'duplicate-me' },
    });
    const beforeCount = app.getUiSnapshot().sheets.length;
    app.duplicateSheet(sourceId);

    const snapshot = app.getUiSnapshot();
    assert.equal(snapshot.sheets.length, beforeCount + 1);
    assert.notEqual(app.getActiveSheetId(), sourceId);
    assert.equal(
      app['runtime'].model.getSheet(app.getActiveSheetId()).cells.get(0, 0)?.value,
      'duplicate-me',
    );
  });

  it('selectSheet updates only transient workbook session state', () => {
    const app = new WorkbookSession();
    const sourceId = app.getActiveSheetId();
    app.runCommand('sheet.add', { id: 'sheet-2', name: 'Second' });

    app.selectSheet('sheet-2');
    assert.equal(app.getActiveSheetId(), 'sheet-2');
    assert.equal(app['runtime'].model.primarySheetId, sourceId);
    assert.equal(app['runtime'].commands.getHistoryDepth().undo, 1);
  });

  it('edits the active canvas cell instead of the top-left of a dragged range', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.applyCanvasSelection({
      ranges: [{ sheetId, startRow: 1, endRow: 8, startColumn: 0, endColumn: 2 }],
      primaryRangeIndex: 0,
      activeCell: { row: 8, column: 2 },
      anchorCell: { row: 1, column: 0 },
    });
    assert.equal(app.getUiSnapshot().activeCell, 'C9');

    app.beginEdit('4');
    assert.deepEqual(app.getUiSnapshot().editingCell, { row: 8, column: 2 });
    app.commitEdit('none');

    const cells = app['runtime'].model.getSheet(sheetId).cells;
    assert.equal(cells.get(8, 2)?.value, 4);
    assert.equal(cells.get(1, 0), undefined);
  });

  it('cell insert down preserves the selected data and shifts the following band', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'top' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 1,
      column: 0,
      value: { value: 'bottom' },
    });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });
    app.applyCellShift('insert', 'row');

    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.cells.get(1, 0)?.value, 'top');
    assert.equal(sheet.cells.get(2, 0)?.value, 'bottom');
    assert.equal(sheet.cells.get(0, 0)?.value, undefined);
  });

  it('undo reverts the last core editing mutation', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'before-undo' },
    });
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'after-edit' },
    });
    app.undo();
    assert.equal(app['runtime'].model.getSheet(sheetId).cells.get(0, 0)?.value, 'before-undo');
  });

  it('derives one mixed Home-ribbon state from a multi-cell selection', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'bold', style: { bold: true } } });
    app.runCommand('sheet.cell.set', { sheetId, row: 1, column: 0, value: { value: 'plain' } });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });

    const home = app.getUiSnapshot().homeRibbon;
    assert.ok(home.mixedStyleKeys.includes('bold'));
    assert.equal(home.style.bold, undefined);
    assert.equal(home.canFormat, true);
  });

  it('uses contents as the default clear mode for a context-resolved Home descriptor', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'keep-style', style: { bold: true } } });
    selectCell(app, 0, 0);

    app.dispatch({ commandId: 'sheet.range.clear' });

    const cell = app['runtime'].model.getSheet(sheetId).cells.get(0, 0);
    assert.equal(cell?.value, null);
    assert.equal(cell?.style?.bold, true);
  });

  it('requires confirmation before a merge discards non-anchor content and keeps one undo entry', () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 0, value: { value: 'anchor' } });
    app.runCommand('sheet.cell.set', { sheetId, row: 0, column: 1, value: { value: 'discarded' } });
    app.runCommand('selection.set', {
      sheetId,
      ranges: [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }],
      primaryRangeIndex: 0,
      activeCell: { row: 0, column: 0 },
      anchorCell: { row: 0, column: 0 },
    });

    const before = app['runtime'].commands.getHistoryDepth().undo;
    app.requestMergeAction('center');
  assert.equal(app.getUiSnapshot().dialogs.active, 'merge-confirm');
  assert.equal(app.getUiSnapshot().dialogs.mergeDiscardCount, 1);

    app.confirmMergeAction();
    const sheet = app['runtime'].model.getSheet(sheetId);
    assert.equal(sheet.merges.length, 1);
    assert.equal(sheet.cells.get(0, 0)?.value, 'anchor');
    assert.equal(app['runtime'].commands.getHistoryDepth().undo, before + 1);
    app.undo();
    assert.equal(sheet.merges.length, 0);
    assert.equal(sheet.cells.get(0, 1)?.value, 'discarded');
  });
});
