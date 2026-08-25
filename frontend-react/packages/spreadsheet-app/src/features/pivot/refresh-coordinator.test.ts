import assert from 'node:assert/strict';
import { it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import type { MutationInfo } from '@react-sheets/command-runtime';
import { buildPivotModel } from './helpers';
import { pivotIdsToRefresh } from './refresh-coordinator';

function fixture() {
  const workbook = new WorkbookModel('refresh-policy', 'Refresh Policy');
  const sheet = workbook.getSheet('sheet-1');
  [['Region', 'Amount'], ['East', 10], ['West', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  const range = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 };
  const base = buildPivotModel(workbook, sheet.id, 'base', range)!;
  const make = (id: string, mode: 'manual' | 'on-open' | 'on-change') => ({
    ...structuredClone(base),
    id,
    target: { sheetId: sheet.id, anchor: { row: Number(id.slice(-1)) * 5, column: 0 } },
    refreshPolicy: { mode, preserveFormatting: true, refreshOnLoad: mode !== 'manual' },
  });
  const pivots = [make('manual-1', 'manual'), make('open-2', 'on-open'), make('change-3', 'on-change')];
  const sourceMutation = {
    id: 'cell.set', unitId: workbook.unitId, sheetId: sheet.id,
    params: { sheetId: sheet.id, row: 1, column: 1, value: { value: 30 } },
    affectedRanges: [range],
  } satisfies MutationInfo;
  const unrelatedMutation = { ...sourceMutation, affectedRanges: [{ ...range, startRow: 10, endRow: 10 }] } satisfies MutationInfo;
  return { workbook, pivots, sourceMutation, unrelatedMutation };
}

it('selects only canonical policy targets for open, source, and explicit refreshes', () => {
  const { workbook, pivots, sourceMutation, unrelatedMutation } = fixture();
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'open' }), ['open-2']);
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'source-change', mutations: [sourceMutation] }), ['change-3']);
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'source-change', mutations: [unrelatedMutation] }), []);
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'explicit', pivotId: 'manual-1' }), ['manual-1']);
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'explicit', pivotId: 'missing' }), []);
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'explicit-all' }), ['manual-1', 'open-2', 'change-3']);
});

it('does not treat policy metadata or Pivot mutations as source changes', () => {
  const { workbook, pivots, sourceMutation } = fixture();
  const pivotMutation = { ...sourceMutation, id: 'pivot.update', affectedRanges: [{ ...sourceMutation.affectedRanges[0]! }] };
  assert.deepEqual(pivotIdsToRefresh(workbook, pivots, { kind: 'source-change', mutations: [pivotMutation] }), []);
});
