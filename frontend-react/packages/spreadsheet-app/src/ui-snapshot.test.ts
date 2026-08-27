import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { WorkbookModel, type DataSourceManifest, type DataBlockRef } from '@react-sheets/core-model';
import { buildCanvasSheetSnapshot } from './ui-snapshot';
import { setCellHyperlink } from './features/review';
import { LocalDataBlockStore } from './features/persistence/data-block-store';
import { WorkspaceMemoryCoordinator } from './features/persistence/memory';
import {
  computeColumnarBlockChecksum,
  encodeColumnarBlock,
  DataSourceContentQuery,
  migrateDataRegionCellPatches,
} from './features/data-source';

describe('canonical drawing UI projection', () => {
  it('collapses filtered rows in the render projection without blank placeholders', () => {
    const workbook = new WorkbookModel('snapshot-filter', 'Snapshot Filter');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 5;
    sheet.columnCount = 2;
    sheet.cells.set(0, 0, { value: 'State' });
    sheet.cells.set(0, 1, { value: 'Value' });
    sheet.cells.set(1, 0, { value: 'Keep' });
    sheet.cells.set(2, 0, { value: 'Drop' });
    sheet.cells.set(3, 0, { value: 'Keep' });
    sheet.cells.set(4, 0, { value: 'Drop' });
    sheet.autoFilter = {
      sheetId: sheet.id,
      range: { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
      columns: {
        0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['Keep'], includeBlank: false } },
        1: { column: 1, showButton: true, hiddenButton: false },
      },
    };

    const snapshot = buildCanvasSheetSnapshot(workbook, sheet, new FormulaEngine({ defaultSheetId: sheet.id }), true);
    assert.deepEqual(snapshot.hiddenRows, [2, 4]);
    assert.equal(snapshot.getCell(2, 0)?.value, 'Drop');
  });

  it('uses the current FormulaEngine result for filter domains and row visibility', () => {
    const workbook = new WorkbookModel('snapshot-formula-filter', 'Snapshot Formula Filter');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 4;
    sheet.columnCount = 2;
    sheet.cells.set(0, 0, { value: 'Status' });
    sheet.cells.set(0, 1, { value: 'Source' });
    sheet.cells.set(1, 0, { value: null, formula: '=B2' });
    sheet.cells.set(2, 0, { value: null, formula: '=B3' });
    sheet.cells.set(3, 0, { value: null, formula: '=B4' });
    sheet.cells.set(1, 1, { value: 'Open' });
    sheet.cells.set(2, 1, { value: 'Closed' });
    sheet.cells.set(3, 1, { value: 'Open' });
    sheet.autoFilter = {
      sheetId: sheet.id,
      range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
      columns: {
        0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['Open'], includeBlank: false } },
      },
    };
    const formula = new FormulaEngine({ defaultSheetId: sheet.id });
    formula.setValue({ sheetId: sheet.id, row: 1, column: 1 }, 'Open');
    formula.setValue({ sheetId: sheet.id, row: 2, column: 1 }, 'Closed');
    formula.setValue({ sheetId: sheet.id, row: 3, column: 1 }, 'Open');
    formula.setFormula({ sheetId: sheet.id, row: 1, column: 0 }, '=B2');
    formula.setFormula({ sheetId: sheet.id, row: 2, column: 0 }, '=B3');
    formula.setFormula({ sheetId: sheet.id, row: 3, column: 0 }, '=B4');
    formula.recalculate();

    const snapshot = buildCanvasSheetSnapshot(workbook, sheet, formula, true);
    assert.deepEqual(snapshot.getFilterValueDomain(0), ['Closed', 'Open']);
    assert.deepEqual(snapshot.hiddenRows, [2]);
    assert.equal(snapshot.getCell(1, 0)?.value, 'Open');
    assert.equal(snapshot.getCell(2, 0)?.value, 'Closed');
    formula.setValue({ sheetId: sheet.id, row: 2, column: 1 }, 'Open');
    formula.recalculate();
    const recalculated = buildCanvasSheetSnapshot(workbook, sheet, formula, true);
    assert.deepEqual(recalculated.getFilterValueDomain(0), ['Open']);
    assert.deepEqual(recalculated.hiddenRows, []);
  });

  it('exposes conditional-format colors to the filter domain without writing styles into cells', () => {
    const workbook = new WorkbookModel('snapshot-filter-color', 'Snapshot Filter Color');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 3;
    sheet.columnCount = 1;
    sheet.cells.set(0, 0, { value: 'Status' });
    sheet.cells.set(1, 0, { value: 10 });
    sheet.cells.set(2, 0, { value: 1, filterMetadata: { color: { target: 'cell', dxfId: 7, value: '#123456' } } });
    sheet.conditionalFormats.push({
      id: 'snapshot-filter-color',
      sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }],
      type: 'highlight',
      operator: 'greaterThan',
      value1: 5,
      priority: 1,
      style: { background: '#ff0000' },
    });
    sheet.autoFilter = {
      sheetId: sheet.id,
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
      columns: {
        0: {
          column: 0, showButton: true, hiddenButton: false,
          criterion: { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#FF0000' } },
        },
      },
    };

    const snapshot = buildCanvasSheetSnapshot(workbook, sheet, new FormulaEngine({ defaultSheetId: sheet.id }), true);
    assert.deepEqual(snapshot.getFilterColorDomain(0), [
      { target: 'cell', color: '#123456' },
      { target: 'cell', color: '#ff0000' },
    ]);
    assert.deepEqual(snapshot.hiddenRows, [2]);
    assert.equal(sheet.cells.get(1, 0)?.style, undefined);
    assert.equal(snapshot.getCell(1, 0)?.overlay?.style?.background, '#ff0000');
  });

  it('exposes only DrawingObject and DrawingPayload to consumers', () => {
    const workbook = new WorkbookModel('snapshot-drawings', 'Snapshot Drawings');
    const sheet = workbook.getSheet(workbook.getSheets()[0]!.id);
    sheet.drawings.push({
      id: 'drawing-combo',
      sheetId: sheet.id,
      kind: 'chart',
      payloadId: 'chart-combo',
      anchor: { kind: 'absolute' },
      transform: { x: 10, y: 20, width: 320, height: 200 },
      zIndex: 1,
    });
    sheet.drawingPayloads.set('chart-combo', {
      kind: 'chart',
      chartId: 'chart-combo',
      chartType: 'combo',
      stacked: 'percent',
      sourceRanges: [{ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }],
      elements: { hiddenData: 'show' },
    });

    const snapshot = buildCanvasSheetSnapshot(workbook, sheet, new FormulaEngine({ defaultSheetId: sheet.id }), true);
    assert.equal(snapshot.drawings[0]?.id, 'drawing-combo');
    assert.equal(snapshot.drawingPayloads.get('chart-combo')?.kind, 'chart');
    assert.equal((snapshot.drawingPayloads.get('chart-combo') as { chartType?: string; stacked?: string }).chartType, 'combo');
    assert.equal((snapshot.drawingPayloads.get('chart-combo') as { chartType?: string; stacked?: string }).stacked, 'percent');
    assert.equal('charts' in snapshot, false);
    assert.equal('shapes' in snapshot, false);
    assert.equal('images' in snapshot, false);
  });

  it('renders block values and sparse styles through the same resolved-cell surface', async () => {
    const workbook = new WorkbookModel('snapshot-block', 'Snapshot Block');
    const sheet = workbook.getSheet('sheet-1');
    sheet.rowCount = 3;
    sheet.columnCount = 2;
    const fields = [
      { id: 'code', name: 'Code', ordinal: 0, type: 'text' as const },
      { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' as const },
    ];
    const payload = await encodeColumnarBlock({ fields, rows: [['A', 10], ['B', 20]] });
    const sourceId = 'snapshot-block-source';
    const ref: DataBlockRef = {
      id: 'snapshot-block-0',
      dataSourceId: sourceId,
      startRow: 0,
      rowCount: 2,
      storageKey: 'snapshot-block-source/snapshot-block-0',
      checksum: await computeColumnarBlockChecksum(payload),
      byteLength: payload.byteLength,
      encoding: 'columnar-v1',
      revision: 0,
    };
    const manifest: DataSourceManifest = {
      schema: 'DataSourceManifest',
      version: 1,
      id: sourceId,
      name: 'Snapshot block source',
      kind: 'chunked-table',
      rowCount: 2,
      fields,
      blockRowCount: 65_536,
      blocks: [ref],
      revision: 0,
    };
    const store = new LocalDataBlockStore(new WorkspaceMemoryCoordinator());
    await store.put(ref, payload);
    const query = new DataSourceContentQuery(manifest, store);
    workbook.addDataSource(manifest);
    sheet.dataRegions.push({
      id: 'snapshot-block-region',
      sourceId,
      range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      headerRow: 0,
      revision: 0,
    });
    // Legacy sparse metadata entry contains the source value but must not
    // shadow the immutable block value after the block is loaded.
    sheet.cells.set(1, 1, { value: 999, style: { bold: true } });
    migrateDataRegionCellPatches(sheet);
    await query.getRowValues(0);

    const snapshot = buildCanvasSheetSnapshot(
      workbook,
      sheet,
      new FormulaEngine({ defaultSheetId: sheet.id }),
      true,
      {},
      new Map([[sourceId, query]]),
    );
    const cell = snapshot.getCell(1, 1);
    assert.equal(cell?.value, '10');
    assert.equal(cell?.displayValue, '10');
    assert.equal(cell?.style?.bold, true);
  });

  it('projects review metadata from canonical sheet stores only', () => {
    const workbook = new WorkbookModel('snapshot-review', 'Snapshot Review');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, {
      value: 'link',
      // Legacy fields are intentionally ignored by the UI projection.
      hyperlink: 'https://legacy.invalid',
      hyperlinkDetail: { id: 'legacy', target: { kind: 'url', url: 'https://legacy.invalid' } },
      comment: { id: 'legacy-comment', author: 'old', text: 'old', createdAt: '2020-01-01' },
    } as unknown as import('@react-sheets/core-model').CellData);
    setCellHyperlink(sheet, 0, 0, { id: 'canonical', target: { kind: 'url', url: 'https://canonical.invalid' } });
    sheet.review.addThread({
      id: 'thread-1', sheetId: sheet.id, row: 0, column: 0, author: 'Alice', text: 'Review',
      createdAt: '2026-01-01', replies: [],
    });

    const snapshot = buildCanvasSheetSnapshot(workbook, sheet, new FormulaEngine({ defaultSheetId: sheet.id }), true);
    const cell = snapshot.getCell(0, 0);
    assert.equal(cell?.hyperlink, 'https://canonical.invalid');
    assert.equal(cell?.comment?.id, 'thread-1');
    assert.equal(cell?.comment?.text, 'Review');
  });
});
