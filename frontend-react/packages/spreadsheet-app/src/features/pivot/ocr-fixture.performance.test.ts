import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { performance } from 'node:perf_hooks';
import { it } from 'node:test';
import { pivotSourceIdentity, type PivotModel } from '@react-sheets/core-model';
import { importXlsx } from '@react-sheets/exchange-excel-ooxml';
import { createSpreadsheetRuntime, disposeSpreadsheetRuntime, hydrateRuntime } from '../../runtime';
import { buildPivotGridProjection, preparePivotTaskDescriptor, preparePivotTaskInputAsync } from './engine';
import { InlinePivotTaskPort } from './task-port';
import { createPivotCalculateRequest, createPivotSourceRegisterRequest } from './task-protocol';

const fixturePath = process.env.OCR_XLSX_FIXTURE ?? 'C:\\Users\\kuo13\\Downloads\\OCR结果.xlsx';

it('meets the real OCR workbook import, source-index, worker, and projection budgets', { skip: !existsSync(fixturePath) }, async () => {
  const importStartedAt = performance.now();
  const bytes = readFileSync(fixturePath);
  const imported = await importXlsx({ fileName: 'OCR结果.xlsx', buffer: bytes, options: { compatibilityTarget: 'B' } });
  const importedAt = performance.now();
  const runtime = createSpreadsheetRuntime({ unitId: imported.snapshot.unitId });
  hydrateRuntime(runtime, { snapshot: imported.snapshot, revision: 0 });
  await runtime.formulaCalculation;
  const hydratedAt = performance.now();
  const taskPort = new InlinePivotTaskPort();
  try {
    const sheet = runtime.model.getSheets()[0]!;
    assert.equal(sheet.usedRange.endRow, 4_058);
    assert.equal(sheet.usedRange.endColumn, 22);
    assert.equal(typeof sheet.cells.get(1, 8)?.value, 'string');
    assert.equal(typeof sheet.cells.get(1, 9)?.value, 'string');
    assert.equal(typeof sheet.cells.get(1, 21)?.value, 'string');
    runtime.model.addSheet('pivot-benchmark-target', 'Pivot Benchmark', 2_000, 512);
    const source = { kind: 'worksheet-range' as const, range: { sheetId: sheet.id, startRow: 0, endRow: 4_058, startColumn: 0, endColumn: 22 } };
    const draft: PivotModel = {
      schema: 'PivotDefinition',
      id: 'pivot-ocr-benchmark',
      source,
      target: { sheetId: 'pivot-benchmark-target', anchor: { row: 0, column: 0 } },
      fieldCatalog: { schema: 'PivotFieldCatalog', fields: [] },
      refreshPolicy: { mode: 'manual', preserveFormatting: true, refreshOnLoad: false },
      layout: {
        rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true,
        collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' },
        values: [], calculatedFields: [], calculatedItems: [], subtotalLocation: 'bottom',
        showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact',
        expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
      },
    };
    const chunkDurations: number[] = [];
    const sourceStartedAt = performance.now();
    const prepared = await preparePivotTaskInputAsync(runtime.model, draft, runtime.formula, { onChunk: ({ durationMs }) => chunkDurations.push(durationMs) });
    const sourcePreparedAt = performance.now();
    assert.equal(prepared.definition.fieldCatalog.fields.length, 23);
    assert.equal(prepared.definition.fieldCatalog.fields.every((field) => field.dataType === 'text' || field.dataType === 'mixed'), true, JSON.stringify(prepared.definition.fieldCatalog.fields.map((field) => [field.name, field.dataType])));
    const page = prepared.definition.fieldCatalog.fields.find((field) => field.name === '页码')!;
    const bom1 = prepared.definition.fieldCatalog.fields.find((field) => field.name === 'BOM1')!;
    const pivot: PivotModel = {
      ...prepared.definition,
      layout: {
        ...prepared.definition.layout,
        rows: [{ fieldId: page.fieldId }],
        values: [{ valueId: 'bom1:count', fieldId: bom1.fieldId, summarizeBy: 'count' }],
      },
    };
    const descriptor = preparePivotTaskDescriptor(runtime.model, pivot, runtime.formula);
    const sourceIdentity = `${runtime.model.unitId}:${pivotSourceIdentity(source)}`;
    const registered = await taskPort.submit(createPivotSourceRegisterRequest('ocr-register', 1, sourceIdentity, descriptor.revisions.sourceRevision, prepared.source));
    assert.equal(registered.status, 'accepted');
    const workerStartedAt = performance.now();
    const calculated = await taskPort.submit(createPivotCalculateRequest('ocr-calculate', 1, sourceIdentity, descriptor.definition, descriptor.controls, descriptor.revisions, descriptor.targetBounds));
    const workerFinishedAt = performance.now();
    assert.equal(calculated.status, 'completed');
    if (calculated.status !== 'completed') return;
    const projectionStartedAt = performance.now();
    const projection = buildPivotGridProjection(runtime.model, pivot, calculated.result);
    const projectionFinishedAt = performance.now();
    const cachedStartedAt = performance.now();
    preparePivotTaskDescriptor(runtime.model, pivot, runtime.formula);
    const singleCachedFinishedAt = performance.now();
    for (let iteration = 1; iteration < 20; iteration += 1) preparePivotTaskDescriptor(runtime.model, pivot, runtime.formula);
    const cachedFinishedAt = performance.now();

    assert.ok(importedAt - importStartedAt < 2_000, `OCR XLSX import exceeded 2000ms: ${Math.round(importedAt - importStartedAt)}ms`);
    assert.ok(hydratedAt - importedAt < 800, `OCR workbook hydration exceeded 800ms: ${Math.round(hydratedAt - importedAt)}ms`);
    assert.ok(sourcePreparedAt - sourceStartedAt < 300, `OCR source index exceeded 300ms: ${Math.round(sourcePreparedAt - sourceStartedAt)}ms`);
    assert.ok(Math.max(...chunkDurations) < 50, `OCR source chunk exceeded 50ms: ${Math.round(Math.max(...chunkDurations))}ms`);
    assert.ok(workerFinishedAt - workerStartedAt < 500, `OCR Pivot worker exceeded 500ms: ${Math.round(workerFinishedAt - workerStartedAt)}ms`);
    assert.ok(projectionFinishedAt - projectionStartedAt < 100, `OCR Pivot projection exceeded 100ms: ${Math.round(projectionFinishedAt - projectionStartedAt)}ms`);
    assert.ok(singleCachedFinishedAt - cachedStartedAt < 10, `Cached Pivot descriptor exceeded 10ms: ${Math.round(singleCachedFinishedAt - cachedStartedAt)}ms`);
    assert.ok(cachedFinishedAt - cachedStartedAt < 30, `20 cached Pivot descriptors exceeded 30ms: ${Math.round(cachedFinishedAt - cachedStartedAt)}ms`);
    assert.equal(projection.cells.some((cell) => cell.kind === 'value'), true);
    if (process.env.PIVOT_BENCHMARK_VERBOSE === '1') console.log(JSON.stringify({
      importMs: Math.round(importedAt - importStartedAt),
      hydrationMs: Math.round(hydratedAt - importedAt),
      sourceIndexMs: Math.round(sourcePreparedAt - sourceStartedAt),
      maxSourceChunkMs: Math.round(Math.max(...chunkDurations) * 10) / 10,
      workerMs: Math.round(workerFinishedAt - workerStartedAt),
      projectionMs: Math.round(projectionFinishedAt - projectionStartedAt),
      cachedDescriptors20Ms: Math.round((cachedFinishedAt - cachedStartedAt) * 10) / 10,
    }));
  } finally {
    taskPort.dispose();
    disposeSpreadsheetRuntime(runtime);
  }
});
