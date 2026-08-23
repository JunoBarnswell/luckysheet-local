import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { WorkbookModel } from '@react-sheets/core-model';
import { buildCanvasSheetSnapshot } from './ui-snapshot';
import { setCellHyperlink } from './features/review';

describe('canonical drawing UI projection', () => {
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

  it('projects review metadata from canonical sheet stores only', () => {
    const workbook = new WorkbookModel('snapshot-review', 'Snapshot Review');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, {
      value: 'link',
      // Legacy fields are intentionally ignored by the UI projection.
      hyperlink: 'https://legacy.invalid',
      hyperlinkDetail: { id: 'legacy', target: { kind: 'url', url: 'https://legacy.invalid' } },
      comment: { id: 'legacy-comment', author: 'old', text: 'old', createdAt: '2020-01-01' },
    });
    setCellHyperlink(sheet, 0, 0, { id: 'canonical', target: { kind: 'url', url: 'https://canonical.invalid' } });
    sheet.commentThreads.push({
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
