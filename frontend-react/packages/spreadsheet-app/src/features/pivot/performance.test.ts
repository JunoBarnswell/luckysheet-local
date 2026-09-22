import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { buildPivotGridProjection, computePivotResult, evaluatePivotTask, findPivotProjectionCellAt, getPivotFieldCatalog, normalizePivotDefinition, preparePivotTaskInput } from './engine';
import { buildPivotModel } from './helpers';
import { createPivotSourceIndex, estimatePivotSourceIndexBytes } from './source-index';

const ROW_COUNT = 4_058;
const COLUMN_COUNT = 23;

function attachmentScaleWorkbook(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-performance', 'Pivot performance');
  const sheet = workbook.getSheet('sheet-1');
  const headers = ['页码', 'BOM1', 'BOM2', ...Array.from({ length: COLUMN_COUNT - 3 }, (_, index) => `文本${index + 1}`)];
  headers.forEach((value, column) => sheet.cells.set(0, column, { value }));
  for (let row = 1; row <= ROW_COUNT; row += 1) {
    const values = [String(row % 192), String(row % 30), `BOM-${row % 408}`, ...Array.from({ length: COLUMN_COUNT - 3 }, (_, column) => `T${column}-${row % Math.max(1, 14 - column % 10)}`)];
    values.forEach((value, column) => sheet.cells.set(row, column, { value }));
  }
  return workbook;
}

it('keeps attachment-scale Pivot calculation, projection, and lookup inside the interactive budget', () => {
  const workbook = attachmentScaleWorkbook();
  workbook.addSheet('pivot-target', 'Pivot target', 1_000, 512);
  const startedAt = performance.now();
  const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-performance', { sheetId: 'sheet-1', startRow: 0, endRow: ROW_COUNT, startColumn: 0, endColumn: COLUMN_COUNT - 1 });
  assert.ok(pivot);
  pivot.target = { sheetId: 'pivot-target', anchor: { row: 0, column: 0 } };
  const page = pivot.fieldCatalog.fields.find((field) => field.name === '页码')!;
  const bom1 = pivot.fieldCatalog.fields.find((field) => field.name === 'BOM1')!;
  const bom2 = pivot.fieldCatalog.fields.find((field) => field.name === 'BOM2')!;
  pivot.layout.rows = [{ fieldId: bom2.fieldId }];
  pivot.layout.columns = [{ fieldId: page.fieldId }];
  pivot.layout.values = [{ valueId: 'bom1:count', fieldId: bom1.fieldId, summarizeBy: 'count' }];
  const result = computePivotResult(workbook, pivot);
  const projection = buildPivotGridProjection(workbook, pivot, result);
  const projectedAt = performance.now();
  let hits = 0;
  for (let index = 0; index < 10_000; index += 1) {
    const row = 2 + index % result.rows.length;
    const column = 1 + index % result.columnPaths.length;
    if (findPivotProjectionCellAt(projection, row, column)) hits += 1;
  }
  const finishedAt = performance.now();
  assert.equal(result.rows.length, 408);
  assert.equal(result.columnPaths.length, 192);
  assert.equal(hits, 10_000);
  assert.ok(projectedAt - startedAt < 1_000, `attachment-scale Pivot exceeded 1000ms: ${Math.round(projectedAt - startedAt)}ms`);
  assert.ok(finishedAt - projectedAt < 100, `10,000 Pivot cell lookups exceeded 100ms: ${Math.round(finishedAt - projectedAt)}ms`);
});

it('rejects an oversized attachment-scale Pivot from its footprint before materializing projection cells', () => {
  const workbook = attachmentScaleWorkbook();
  workbook.addSheet('pivot-target', 'Pivot target');
  const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-oversized', { sheetId: 'sheet-1', startRow: 0, endRow: ROW_COUNT, startColumn: 0, endColumn: COLUMN_COUNT - 1 });
  assert.ok(pivot);
  const page = pivot.fieldCatalog.fields.find((field) => field.name === '页码')!;
  const bom1 = pivot.fieldCatalog.fields.find((field) => field.name === 'BOM1')!;
  const bom2 = pivot.fieldCatalog.fields.find((field) => field.name === 'BOM2')!;
  pivot.target = { sheetId: 'pivot-target', anchor: { row: 0, column: 0 } };
  pivot.layout.rows = [{ fieldId: bom2.fieldId }, { fieldId: page.fieldId }];
  pivot.layout.values = [{ valueId: 'bom1:count', fieldId: bom1.fieldId, summarizeBy: 'count' }];
  const startedAt = performance.now();
  const definition = normalizePivotDefinition(workbook, pivot);
  assert.throws(() => evaluatePivotTask(preparePivotTaskInput(workbook, definition)), /exceeds the destination worksheet boundary/);
  const finishedAt = performance.now();
  assert.ok(finishedAt - startedAt < 1_000, `oversized Pivot rejection exceeded 1000ms: ${Math.round(finishedAt - startedAt)}ms`);
});

it('reuses one revision-owned source index and field catalog across repeated Field List edits', () => {
  const workbook = attachmentScaleWorkbook();
  const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-source-cache', { sheetId: 'sheet-1', startRow: 0, endRow: ROW_COUNT, startColumn: 0, endColumn: COLUMN_COUNT - 1 });
  assert.ok(pivot);
  const startedAt = performance.now();
  for (let iteration = 0; iteration < 20; iteration += 1) {
    const catalog = getPivotFieldCatalog(workbook, pivot);
    assert.equal(catalog.fields.length, COLUMN_COUNT);
  }
  const finishedAt = performance.now();
  assert.ok(finishedAt - startedAt < 300, `20 cached field-catalog reads exceeded 300ms: ${Math.round(finishedAt - startedAt)}ms`);
});

it('keeps a 100k x 20 columnar Pivot task within the worker and memory budgets', () => {
  const rowCount = 100_000;
  const columnCount = 20;
  const columns = Array.from({ length: columnCount }, (_, column) => ({
    field: { fieldId: `field:${column}`, name: `Field ${column}`, ordinal: column, dataType: column === 2 ? 'number' as const : 'text' as const },
    values: Array.from({ length: rowCount }, (_, row) => column === 2 ? row % 1_000 : `V${column}:${row % (column === 0 ? 100 : column === 1 ? 20 : 32)}`),
  }));
  const source = createPivotSourceIndex({
    columns,
    rowPaths: Array.from({ length: rowCount }, (_, row) => [{ sheetId: 'source-sheet', row }]),
  });
  const definition = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-100k',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'source-sheet', startRow: 0, endRow: rowCount, startColumn: 0, endColumn: columnCount - 1 } },
    target: { sheetId: 'target-sheet', anchor: { row: 0, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: source.fields.map((field) => ({ fieldId: field.fieldId, name: field.name, dataType: field.dataType ?? 'mixed' as const, ordinal: field.ordinal })) },
    refreshPolicy: { mode: 'manual' as const, preserveFormatting: true, refreshOnLoad: false },
    layout: {
      rows: [{ fieldId: 'field:0' }], columns: [{ fieldId: 'field:1' }], filters: [], allowMultipleFiltersPerField: true,
      collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const },
      values: [{ valueId: 'field:2:sum', fieldId: 'field:2', summarizeBy: 'sum' as const }],
      calculatedFields: [], calculatedItems: [], subtotalLocation: 'bottom' as const,
      showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const,
      expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
    },
    presentation: {
      styleOptions: { showRowHeaders: true, showColumnHeaders: true, showRowStripes: false, showColumnStripes: false, showLastColumn: false },
    },
  };
  const startedAt = performance.now();
  const result = evaluatePivotTask({
    definition,
    source,
    controls: [],
    revisions: { pivotId: definition.id, sourceRevision: 'source-1', layoutRevision: 'layout-1', filterRevision: 'filter-1' },
    targetBounds: { rowCount: 2_000, columnCount: 512 },
  });
  const finishedAt = performance.now();
  const legacyRowObjectEstimate = rowCount * columnCount * 56 + rowCount * 48;
  assert.equal(result.rows.length, 100);
  assert.equal(result.columnPaths.length, 20);
  assert.ok(finishedAt - startedAt < 2_000, `100k x 20 Pivot task exceeded 2000ms: ${Math.round(finishedAt - startedAt)}ms`);
  assert.ok(estimatePivotSourceIndexBytes(source) < legacyRowObjectEstimate * 0.35);
});

it('rejects a sparse high-cardinality source before allocating a dense 300k-cell result', () => {
  const rowCount = 600;
  const source = createPivotSourceIndex({
    columns: [
      { field: { fieldId: 'row', name: 'Row', ordinal: 0, dataType: 'text' }, values: Array.from({ length: rowCount }, (_, row) => `R${row}`) },
      { field: { fieldId: 'column', name: 'Column', ordinal: 1, dataType: 'text' }, values: Array.from({ length: rowCount }, (_, row) => `C${row % 500}`) },
      { field: { fieldId: 'value', name: 'Value', ordinal: 2, dataType: 'number' }, values: Array.from({ length: rowCount }, () => 1) },
    ],
    rowPaths: Array.from({ length: rowCount }, (_, row) => [{ sheetId: 'source-sheet', row }]),
  });
  const definition = {
    schema: 'PivotDefinition' as const,
    id: 'pivot-result-limit',
    source: { kind: 'worksheet-range' as const, range: { sheetId: 'source-sheet', startRow: 0, endRow: rowCount, startColumn: 0, endColumn: 2 } },
    target: { sheetId: 'target-sheet', anchor: { row: 0, column: 0 } },
    fieldCatalog: { schema: 'PivotFieldCatalog' as const, fields: source.fields.map((field) => ({ fieldId: field.fieldId, name: field.name, dataType: field.dataType ?? 'mixed' as const, ordinal: field.ordinal })) },
    refreshPolicy: { mode: 'manual' as const, preserveFormatting: true, refreshOnLoad: false },
    layout: {
      rows: [{ fieldId: 'row' }], columns: [{ fieldId: 'column' }], filters: [], allowMultipleFiltersPerField: true,
      collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const },
      values: [{ valueId: 'value:sum', fieldId: 'value', summarizeBy: 'sum' as const }],
      calculatedFields: [], calculatedItems: [], subtotalLocation: 'bottom' as const,
      showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const,
      expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
    },
  };
  assert.throws(() => evaluatePivotTask({
    definition,
    source,
    controls: [],
    revisions: { pivotId: definition.id, sourceRevision: 'source-1', layoutRevision: 'layout-1', filterRevision: 'filter-1' },
    targetBounds: { rowCount: 1_000, columnCount: 1_000 },
  }), /result cell limit exceeded/);
});
