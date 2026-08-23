import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { FormulaEngine } from '@react-sheets/formula-engine';
import { WorkbookModel } from '@react-sheets/core-model';
import { buildCanvasSheetSnapshot } from './ui-snapshot';

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
});
