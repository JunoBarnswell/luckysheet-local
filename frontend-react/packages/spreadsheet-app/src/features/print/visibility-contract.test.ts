import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { ResolvedVisibility } from '@react-sheets/sheet-features';
import { computePrintPages, createDefaultPrintLayout } from './index';

function visibility(revision: number, hiddenRows: readonly number[] = []): ResolvedVisibility {
  const rows = new Map(hiddenRows.map((row) => [row, { manualHidden: false, filterHidden: true, outlineHidden: false }]));
  const columns = new Map<number, { manualHidden: boolean }>();
  return {
    revision,
    rows,
    columns,
    isRowHidden: (row) => rows.has(row),
    isColumnHidden: (column) => columns.has(column),
  };
}

test('print pagination projects hidden rows from the pinned visibility object', () => {
  const layout = createDefaultPrintLayout('wb-print-visibility', 'sheet-1', visibility(7, [1]));
  layout.printAreas = [{ sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 } }];
  const pages = computePrintPages(layout, 20, 80);
  assert.deepEqual(pages[0]?.range, { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 });
  assert.equal(pages[0]?.heightPx, 40);
});

test('print pagination has no implicit visibility fallback', () => {
  const layout = createDefaultPrintLayout('wb-print-visibility', 'sheet-1', visibility(7));
  layout.printAreas = [{ sheetId: 'sheet-1', range: { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 } }];
  assert.doesNotThrow(() => computePrintPages(layout));
  assert.throws(() => computePrintPages({ ...layout, resolvedVisibility: undefined as never }), /RESOLVED_VISIBILITY_REQUIRED/);
});
