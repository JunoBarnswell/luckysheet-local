import assert from 'node:assert/strict';
import { test } from 'node:test';
import type { ChartPayload } from './commands';
import { chartSourceRevision, resolveChartDataFromSources } from './data';
import type { ResolvedVisibility } from '@react-sheets/sheet-features';

function visibility(revision: number): ResolvedVisibility {
  const rows = new Map<number, { manualHidden: boolean; filterHidden: boolean; outlineHidden: boolean }>();
  const columns = new Map<number, { manualHidden: boolean }>();
  return {
    revision,
    rows,
    columns,
    isRowHidden: (row) => rows.has(row),
    isColumnHidden: (column) => columns.has(column),
  };
}

function payload(): ChartPayload {
  return {
    kind: 'chart',
    chartId: 'visibility-contract',
    chartType: 'column',
    subtype: 'clustered',
    source: { kind: 'worksheet-ranges', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }] },
    elements: { hiddenData: 'show' },
  };
}

test('chart source revision is pinned to the visibility revision', () => {
  const source = { getCell: () => undefined, resolvedVisibility: visibility(4), revision: 4 };
  const changed = { ...source, resolvedVisibility: visibility(5), revision: 5 };
  const first = chartSourceRevision(payload(), () => source);
  const second = chartSourceRevision(payload(), () => changed);
  assert.notEqual(first, second);
});

test('chart resolution rejects a stale visibility projection', () => {
  const result = resolveChartDataFromSources(payload(), () => ({
    getCell: () => undefined,
    resolvedVisibility: visibility(4),
    revision: 5,
  }));
  assert.equal(result.status.kind, 'invalid');
  assert.match(result.status.message ?? '', /STALE_VISIBILITY/);
});
