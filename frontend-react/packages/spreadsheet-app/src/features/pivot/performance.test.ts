import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { buildPivotGridProjection, computePivotResult, detectPivotCollision, findPivotProjectionCellAt, getPivotFieldCatalog, getPivotOccupiedRange, normalizePivotDefinition } from './engine';
import { buildPivotModel } from './helpers';

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
  const startedAt = performance.now();
  const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-performance', { sheetId: 'sheet-1', startRow: 0, endRow: ROW_COUNT, startColumn: 0, endColumn: COLUMN_COUNT - 1 });
  assert.ok(pivot);
  const page = pivot.fieldCatalog.fields.find((field) => field.name === '页码')!;
  const bom1 = pivot.fieldCatalog.fields.find((field) => field.name === 'BOM1')!;
  const bom2 = pivot.fieldCatalog.fields.find((field) => field.name === 'BOM2')!;
  pivot.layout.rows = [{ fieldId: bom2.fieldId }];
  pivot.layout.columns = [{ fieldId: page.fieldId }];
  pivot.layout.values = [{ valueId: 'bom1:count', fieldId: bom1.fieldId, summarizeBy: 'count' }];
  const result = computePivotResult(workbook, pivot);
  const projection = buildPivotGridProjection(workbook, pivot, result, { refreshAuthorized: true });
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
  const result = computePivotResult(workbook, definition);
  const occupied = getPivotOccupiedRange(definition, result);
  const collision = detectPivotCollision(workbook, definition, occupied);
  const finishedAt = performance.now();
  assert.equal(collision.reasons.includes('worksheet-bounds'), true);
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
