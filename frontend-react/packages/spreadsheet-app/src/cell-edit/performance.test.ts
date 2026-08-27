import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import test from 'node:test';
import { CellEditDomain } from './domain';
import { ColumnValueAutocompleteIndex } from './value-autocomplete';
import type { CellEditEntryContext } from './contracts';

function entry(text = ''): CellEditEntryContext {
  return {
    target: { display: { sheetId: 'sheet-1', row: 0, column: 0 }, canonical: { sheetId: 'sheet-1', row: 0, column: 0 } },
    source: 'f2', surface: 'grid', editorKind: 'text', editorSurface: { kind: 'text', inputMode: 'text', multiline: true },
    initialDraft: { kind: 'plain', text }, caret: { start: text.length, end: text.length },
    originalSelection: { unitId: 'unit-1', sheetId: 'sheet-1', ranges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }], primaryRangeIndex: 0, activeCell: { row: 0, column: 0 }, anchorCell: { row: 0, column: 0 } },
    originalCell: text ? { value: text } : null, baseCellFingerprint: text ? JSON.stringify({ value: text }) : 'null', enterMove: 'down', groupedSheetIds: ['sheet-1'],
  };
}

function percentile(values: readonly number[], quantile: number): number {
  const ordered = [...values].sort((left, right) => left - right);
  return ordered[Math.min(ordered.length - 1, Math.floor(ordered.length * quantile))] ?? 0;
}

function metrics(values: readonly number[]) {
  return { p50: percentile(values, 0.5), p95: percentile(values, 0.95), p99: percentile(values, 0.99), max: Math.max(...values) };
}

test('CellEditDomain remains inside the synchronous input budget and history memory bound', () => {
  const domain = new CellEditDomain();
  domain.dispatch({ type: 'begin', entry: entry() });
  const durations: number[] = [];
  for (let index = 0; index < 1_000; index += 1) {
    const started = performance.now();
    domain.dispatch({ type: 'text.insert', text: String(index % 10) });
    durations.push(performance.now() - started);
  }
  const measured = metrics(durations);
  assert.ok(measured.p95 <= 1, `draft reducer p95 ${measured.p95.toFixed(3)}ms exceeds 1ms`);
  assert.ok(measured.p99 <= 2, `draft reducer p99 ${measured.p99.toFixed(3)}ms exceeds 2ms`);
  assert.ok(domain.getHistoryMetrics().retainedBytes <= 4 * 1024 * 1024);
  process.stdout.write(`cell-edit.reducer ${JSON.stringify(measured)} history=${JSON.stringify(domain.getHistoryMetrics())}\n`);
});

test('32 KiB draft edits stay inside the long-text budget', () => {
  const text = 'A'.repeat(32 * 1024);
  const domain = new CellEditDomain();
  domain.dispatch({ type: 'begin', entry: entry(text) });
  const durations: number[] = [];
  for (let index = 0; index < 100; index += 1) {
    domain.dispatch({ type: 'caret.set', caret: { start: text.length + index, end: text.length + index } });
    const started = performance.now();
    domain.dispatch({ type: 'text.insert', text: 'x' });
    durations.push(performance.now() - started);
  }
  const measured = metrics(durations);
  assert.ok(measured.p95 <= 4, `32 KiB reducer p95 ${measured.p95.toFixed(3)}ms exceeds 4ms`);
  assert.ok(measured.p99 <= 8, `32 KiB reducer p99 ${measured.p99.toFixed(3)}ms exceeds 8ms`);
  process.stdout.write(`cell-edit.long-text ${JSON.stringify(measured)}\n`);
});

test('100k same-column AutoComplete builds incrementally and queries below 4ms p95', async () => {
  const index = new ColumnValueAutocompleteIndex();
  const heapBefore = process.memoryUsage().heapUsed;
  const entries = Array.from({ length: 100_000 }, (_, row) => ({ row, cell: { value: `Item ${String(row).padStart(6, '0')}` } }));
  const started = performance.now();
  await index.rebuild({ key: 'sheet-1:0', revision: 1, entries, excludeRow: -1, cultureId: 'en-US' }, new AbortController().signal);
  const buildMs = performance.now() - started;
  const heapAfter = process.memoryUsage().heapUsed;
  const durations: number[] = [];
  for (let value = 0; value < 1_000; value += 1) {
    const queryStarted = performance.now();
    const candidate = index.query('sheet-1:0', 1, `Item ${String(value % 100).padStart(3, '0')}`);
    durations.push(performance.now() - queryStarted);
    assert.ok(candidate);
  }
  const measured = metrics(durations);
  assert.ok(measured.p95 <= 4, `value AutoComplete p95 ${measured.p95.toFixed(3)}ms exceeds 4ms`);
  assert.ok(measured.p99 <= 8, `value AutoComplete p99 ${measured.p99.toFixed(3)}ms exceeds 8ms`);
  assert.ok(heapAfter - heapBefore <= 80 * 1024 * 1024, `value AutoComplete retained ${(heapAfter - heapBefore) / 1024 / 1024}MiB exceeds 80MiB`);
  process.stdout.write(`cell-edit.value-autocomplete build=${buildMs.toFixed(3)}ms heapDelta=${((heapAfter - heapBefore) / 1024 / 1024).toFixed(3)}MiB query=${JSON.stringify(measured)}\n`);
});
