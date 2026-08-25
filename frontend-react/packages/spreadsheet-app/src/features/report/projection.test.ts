import test from 'node:test';
import assert from 'node:assert/strict';
import { buildReportProjection } from './projection';
import type { CanvasSheetSnapshot } from '../../ui-snapshot';

const table = { id: 'tasks', name: 'Tasks', sourceSheetId: 'report', sourceRange: { sheetId: 'report', startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 }, rowCount: 2, blockSize: 4096, blocks: [], revision: 0, fields: [{ id: 'title', name: 'Title', ordinal: 0, type: 'text' as const }, { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' as const }] };
const definition = { templateSheetId: 'template', tableId: 'tasks', bindings: [{ cell: { row: 1, column: 0 }, expression: 'title', kind: 'field' as const, direction: 'vertical' as const, fill: 'down' as const }], pagination: { enabled: true, rowsPerPage: 1, repeatHeaderRows: [0] }, renderMode: 'preview' as const, layout: { orientation: 'portrait' as const, marginTopPx: 24, marginRightPx: 24, marginBottomPx: 24, marginLeftPx: 24 }, dataEntry: [] };

function sheetWithRows(rows: string[][], nextDefinition = definition): CanvasSheetSnapshot {
  return { id: 'report', name: 'Report', rowCount: 20, columnCount: 20, columns: [], occupiedCellCount: rows.length * 2, kind: 'report-sheet', reportSheet: structuredClone(nextDefinition), getCell: (row: number, column: number) => row > 0 && rows[row - 1]?.[column] !== undefined ? { value: rows[row - 1]![column]!, address: `A${row}`, displayValue: rows[row - 1]![column]! } : undefined } as unknown as CanvasSheetSnapshot;
}

test('Report projection expands field bindings without writing ordinary cells', () => {
  const projection = buildReportProjection(sheetWithRows([['Alpha', '10'], ['Beta', '20']]), [table]);
  assert.equal(projection.status, 'ready');
  assert.equal(projection.cells.length, 2);
  assert.deepEqual(projection.cells.map((cell) => cell.value), ['Alpha', 'Beta']);
  assert.equal(projection.pageCount, 2);
});

test('Report projection rejects unavailable field bindings', () => {
  const projection = buildReportProjection(sheetWithRows([['Alpha', '10']], { ...definition, bindings: [{ ...definition.bindings[0]!, expression: 'missing' }] }), [table]);
  assert.equal(projection.status, 'error');
  assert.match(projection.error ?? '', /unavailable/);
});
