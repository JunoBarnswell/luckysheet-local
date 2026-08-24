import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type SheetSnapshot, type WorkbookTableModel } from '@react-sheets/core-model';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerDrawingFeature } from '../drawing';
import { registerInsertCommands } from './commands';

function blankTableSheet(): SheetSnapshot {
  return { kind: 'table-sheet', id: 'table-sheet-1', name: '集算表', rowCount: 100, columnCount: 10, cells: {}, merges: [], pane: { kind: 'none' }, pivots: [], sparklines: [], drawings: [], drawingPayloads: {}, defaultRowHeightPx: 20, defaultColumnWidthPx: 80, tableSheet: { viewId: 'table-1', columns: [{ fieldId: 'name', caption: 'Name' }], grouping: [] } };
}

describe('insert feature', () => {
  it('creates an advanced sheet and its data table in one undo transaction', () => {
    const workbook = new WorkbookModel('insert-test', 'Insert');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerInsertCommands(runtime);
    const table: WorkbookTableModel = { id: 'table-1', name: 'Data', rowCount: 0, fields: [{ id: 'name', name: 'Name', ordinal: 0, type: 'text' }], blockSize: 1024, blocks: [], revision: 0 };
    runtime.execute('sheet.create.advanced', { sheet: blankTableSheet(), table });
    assert.equal(workbook.getSheet('table-sheet-1').kind, 'table-sheet');
    assert.equal(workbook.dataModel.tables.has('table-1'), true);
    assert.equal(runtime.undo(), true);
    assert.equal(workbook.sheets.has('table-sheet-1'), false);
    assert.equal(workbook.dataModel.tables.has('table-1'), false);
  });

  it('applies barcode presentation and persists new drawing payload kinds', () => {
    const workbook = new WorkbookModel('insert-object-test', 'Insert Objects');
    const runtime = new CommandRuntime(workbook);
    registerSheetCommands(runtime);
    registerDrawingFeature(runtime);
    registerInsertCommands(runtime);
    workbook.getSheet('sheet-1').cells.set(0, 0, { value: '123456789012' });
    runtime.execute('cell.barcode.apply', { sheetId: 'sheet-1', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], presentation: { kind: 'barcode', symbology: 'ean13', source: { kind: 'cell-value' }, options: { foreground: '#000000', background: '#ffffff', showText: true, quietZone: 2 } } });
    assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0)?.presentation?.kind, 'barcode');
    runtime.execute('drawing.add.camera', { sheetId: 'sheet-1', drawing: { id: 'camera-1', sheetId: 'sheet-1', kind: 'camera', payloadId: 'camera-payload', anchor: { kind: 'absolute' }, transform: { x: 0, y: 0, width: 100, height: 60 }, zIndex: 1 }, payload: { kind: 'camera', sourceRange: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, refreshPolicy: 'live' } });
    assert.equal(workbook.getSheet('sheet-1').drawingPayloads.get('camera-payload')?.kind, 'camera');
    const restored = WorkbookModel.fromSnapshot(workbook.snapshot());
    assert.equal(restored.getSheet('sheet-1').drawingPayloads.get('camera-payload')?.kind, 'camera');
  });
});
