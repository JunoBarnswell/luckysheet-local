import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { createPivotMemberKey, PIVOT_MAX_MEMBER_COUNT, PIVOT_MEMBER_DISPLAY_LIMIT, WorkbookModel, type PivotModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import {
  aggregatePivotValues,
  buildPivotGroupedFilterMembers,
  buildPivotGridProjection,
  canonicalPivotMembers,
  computePivotResult,
  computePivotResultFromBlockSource,
  detectPivotCollision,
  getPivotFieldCatalog,
  getPivotSourceRanges,
  hitTestPivotProjection,
  summarizePivotReportFilters,
} from './engine';
import { buildPivotModel } from './helpers';
import { patchPivotValueField } from './panel-state';
import { buildPivotSlicerDrawing } from '../pivot-controls/helpers';
import { createSpillEnvironment } from '../../formula-spill-sync';

function workbookWithData(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-projection', 'Pivot Projection');
  const sheet = workbook.getSheet('sheet-1');
  [['Region', 'Amount'], ['East', 10], ['West', 20], ['East', 5]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  return workbook;
}

it('keeps repeated source fields independent through Values placement identities', () => {
  const workbook = workbookWithData();
  const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-repeated-values', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
  assert.ok(pivot);
  const region = pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!;
  const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
  pivot.layout.rows = [{ fieldId: region.fieldId }];
  pivot.layout.values = [
    { valueId: 'amount:sum', fieldId: amount.fieldId, summarizeBy: 'sum' },
    { valueId: 'amount:count', fieldId: amount.fieldId, summarizeBy: 'count' },
  ];

  const result = computePivotResult(workbook, pivot);
  assert.deepEqual(result.valueFields?.map((field) => [field.valueId, field.fieldId, field.summarizeBy]), [
    ['amount:sum', amount.fieldId, 'sum'],
    ['amount:count', amount.fieldId, 'count'],
  ]);
  assert.deepEqual(result.grandTotal?.values, [35, 3]);
  const projection = buildPivotGridProjection(workbook, pivot);
  assert.ok(projection.cells.some((cell) => cell.valueId === 'amount:sum'));
  assert.ok(projection.cells.some((cell) => cell.valueId === 'amount:count'));

  const edited = patchPivotValueField(pivot.layout, 'amount:count', { summarizeBy: 'average' });
  assert.equal(edited.values[0]?.summarizeBy, 'sum');
  assert.equal(edited.values[1]?.summarizeBy, 'average');
  assert.throws(() => patchPivotValueField(pivot.layout, amount.fieldId, { summarizeBy: 'max' }), /Unknown Pivot Values placement/);
  assert.throws(() => computePivotResult(workbook, { ...pivot, layout: { ...pivot.layout, values: [{ ...pivot.layout.values[0]!, valueId: 'amount:sum' }, { ...pivot.layout.values[1]!, valueId: 'amount:sum' }] } }), /Duplicate Pivot Values placement identity/);
  assert.throws(() => computePivotResult(workbook, { ...pivot, layout: { ...pivot.layout, rows: [{ fieldId: region.fieldId, sort: { direction: 'ascending', by: 'value', valueId: amount.fieldId } }] } }), /value sort placement is not in Values/);
});

function relationalWorkbook(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-relations', 'Pivot Relations');
  const sheet = workbook.getSheet('sheet-1');
  const ranges = [
    [['CustomerId', 'ProductId', 'Amount'], ['c1', 'p1', 100], ['c2', 'p2', 200], ['c1', 'p2', 50]],
    [['CustomerId', 'Region'], ['c1', 'East'], ['c2', 'West']],
    [['ProductId', 'Category'], ['p1', 'Widget'], ['p2', 'Gadget']],
  ];
  const offsets = [0, 4, 7];
  ranges.forEach((rows, rangeIndex) => rows.forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, offsets[rangeIndex]! + columnIndex, { value }))));
  // The Products range is intentionally on the same worksheet as Orders and
  // Customers; sourceId, rather than sheetId, is the logical identity.
  return workbook;
}

function relationalPivot(workbook: WorkbookModel, order: string[]): PivotModel {
  const ranges = {
    orders: { sourceId: 'orders', range: { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 } },
    customers: { sourceId: 'customers', range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 4, endColumn: 5 } },
    products: { sourceId: 'products', range: { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 7, endColumn: 8 } },
  } as const;
  const source = {
    kind: 'worksheet-ranges' as const,
    ranges: order.map((sourceId) => ranges[sourceId as keyof typeof ranges]),
    relationships: [
      { id: 'orders-customers', left: { sourceId: 'orders', fieldId: 'source:orders:column:0' }, right: { sourceId: 'customers', fieldId: 'source:customers:column:0' }, join: 'left' as const },
      { id: 'orders-products', left: { sourceId: 'orders', fieldId: 'source:orders:column:1' }, right: { sourceId: 'products', fieldId: 'source:products:column:0' }, join: 'left' as const },
    ],
  };
  const pivot: PivotModel = {
    schema: 'PivotDefinition' as const,
    id: `pivot-relational-${order.join('-')}`,
    source,
    target: { sheetId: 'sheet-1', anchor: { row: 10, column: 0 } },
    fieldCatalog: { fields: [] },
    refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
    layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [], subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const },
  };
  const catalog = getPivotFieldCatalog(workbook, pivot);
  const region = catalog.fields.find((field) => field.name === 'Region')!;
  const amount = catalog.fields.find((field) => field.name === 'Amount')!;
  pivot.fieldCatalog = catalog;
  pivot.layout.rows = [{ fieldId: region.fieldId }];
  pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
  return pivot;
}

describe('native PivotGridProjection contract', () => {
  it('keeps the complete typed member domain separate from the bounded display limit', () => {
    const values = Array.from({ length: PIVOT_MEMBER_DISPLAY_LIMIT + 1 }, (_, index) => `Member ${index}`);
    const members = canonicalPivotMembers(values);
    assert.equal(members.length, PIVOT_MEMBER_DISPLAY_LIMIT + 1);
    assert.equal(members.at(-1), `Member ${PIVOT_MEMBER_DISPLAY_LIMIT}`);

    assert.throws(
      () => canonicalPivotMembers(Array.from({ length: PIVOT_MAX_MEMBER_COUNT + 1 }, (_, index) => index)),
      /member domain exceeds/,
    );
  });

  it('renders compact, outline, and tabular report layouts with distinct canonical row semantics', () => {
    const workbook = new WorkbookModel('pivot-report-layouts', 'Pivot Report Layouts');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Category', 'Amount'], ['East', 'Widget', 10], ['East', 'Gadget', 20], ['West', 'Widget', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-report-layouts', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    pivot.fieldCatalog = catalog;
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const category = catalog.fields.find((field) => field.name === 'Category')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }, { fieldId: category.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.subtotalLocation = 'bottom';
    pivot.layout.showRowGrandTotals = false;
    const result = computePivotResult(workbook, pivot);
    const projectionFor = (reportLayout: 'compact' | 'outline' | 'tabular') => {
      const candidate = structuredClone(pivot);
      candidate.layout.reportLayout = reportLayout;
      return buildPivotGridProjection(workbook, candidate, result);
    };
    const compact = projectionFor('compact');
    const outline = projectionFor('outline');
    const tabular = projectionFor('tabular');
    const headers = (projection: ReturnType<typeof buildPivotGridProjection>) => projection.cells.filter((cell) => cell.kind === 'column-header').map((cell) => cell.text);
    assert.deepEqual(headers(compact).slice(0, 2), ['Row Labels', 'Amount']);
    assert.deepEqual(headers(outline).slice(0, 3), ['Region', 'Category', 'Amount']);
    assert.deepEqual(headers(tabular).slice(0, 3), ['Region', 'Category', 'Amount']);
    assert.equal(compact.occupiedRange.endColumn - compact.occupiedRange.startColumn, 1);
    assert.equal(outline.occupiedRange.endColumn - outline.occupiedRange.startColumn, 2);
    assert.equal(tabular.occupiedRange.endColumn - tabular.occupiedRange.startColumn, 2);
    const rowHeaders = (projection: ReturnType<typeof buildPivotGridProjection>) => projection.cells.filter((cell) => cell.kind === 'row-header' || cell.kind === 'expand-toggle').map((cell) => [cell.row, cell.column, cell.text]);
    assert.ok(rowHeaders(compact).some(([, column, text]) => column === 0 && text === 'East / Gadget'));
    assert.ok(rowHeaders(tabular).some(([, column, text]) => column === 0 && text === 'East'));
    assert.ok(rowHeaders(tabular).some(([, column, text]) => column === 1 && text === 'Gadget'));
    assert.ok(rowHeaders(outline).some(([, column, text]) => column === 0 && text === 'East'));
    assert.ok(rowHeaders(outline).some(([, column, text]) => column === 1 && text === 'Gadget'));
  });

  it('projects selected and has-data states independently after another Slicer filters the source', () => {
    const workbook = new WorkbookModel('pivot-slicer-projection', 'Pivot Slicer Projection');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Category', 'Amount'], ['East', 'Widget', 10], ['East', 'Gadget', 20], ['West', 'Gadget', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-slicer-projection', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    pivot.fieldCatalog = catalog;
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const category = catalog.fields.find((field) => field.name === 'Category')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    sheet.pivots.push(pivot);
    const regionSlicer = buildPivotSlicerDrawing({ drawingId: 'region-slicer', payloadId: 'region-slicer-payload', sheetId: sheet.id, pivotId: pivot.id, fieldId: region.fieldId, transform: { x: 0, y: 0, width: 200, height: 120 }, zIndex: 1 });
    const categorySlicer = buildPivotSlicerDrawing({ drawingId: 'category-slicer', payloadId: 'category-slicer-payload', sheetId: sheet.id, pivotId: pivot.id, fieldId: category.fieldId, filter: { mode: 'include', memberKeys: [createPivotMemberKey('Widget')] }, transform: { x: 0, y: 140, width: 200, height: 120 }, zIndex: 2 });
    sheet.drawings.push(regionSlicer.drawing, categorySlicer.drawing);
    sheet.drawingPayloads.set(regionSlicer.drawing.payloadId, regionSlicer.payload);
    sheet.drawingPayloads.set(categorySlicer.drawing.payloadId, categorySlicer.payload);
    const result = computePivotResult(workbook, pivot);
    const items = result.slicerItems?.[regionSlicer.drawing.id] ?? [];
    assert.deepEqual(items.map((item) => [item.value, item.selected, item.hasData]), [['East', true, true], ['West', true, false]]);
    assert.deepEqual(result.slicerItems?.[categorySlicer.drawing.id]?.map((item) => [item.value, item.selected, item.hasData]), [['Gadget', false, true], ['Widget', true, true]]);
  });

  it('reads worksheet-range Pivot values through the FormulaEngine spill authority', () => {
    const workbook = new WorkbookModel('pivot-spill', 'Pivot Spill');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 'Member' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    [10, 20, 30].forEach((value, row) => sheet.cells.set(row + 1, 1, { value }));
    const formula = new FormulaEngine({ defaultSheetId: sheet.id });
    formula.setSpillEnvironment(sheet.id, createSpillEnvironment(sheet));
    formula.setFormula({ sheetId: sheet.id, row: 1, column: 0 }, '=SEQUENCE(3,1,1,1)');
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-spill', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot, formula);
    const member = catalog.fields.find((field) => field.name === 'Member')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.fieldCatalog = catalog;
    pivot.layout.rows = [{ fieldId: member.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const result = computePivotResult(workbook, pivot, formula);
    assert.deepEqual(result.rows.map((node) => node.key), [1, 2, 3]);
    assert.equal(result.grandTotal?.values[0], 60);
  });

  it('resolves named-range spill sources and recomputes their current extent', () => {
    const workbook = new WorkbookModel('pivot-named-spill', 'Pivot Named Spill');
    const sheet = workbook.getSheet('sheet-1');
    workbook.setDefinedName({ name: 'DynamicSource', formula: "='Sheet1'!A1#", scope: 'workbook' });
    const formula = new FormulaEngine({ defaultSheetId: sheet.id });
    formula.setSpillEnvironment(sheet.id, createSpillEnvironment(sheet));
    formula.setFormula({ sheetId: sheet.id, row: 0, column: 0 }, '=SEQUENCE(3,1,0,1)');
    const pivot: PivotModel = {
      schema: 'PivotDefinition',
      id: 'pivot-named-spill',
      source: { kind: 'named-range', name: 'DynamicSource' },
      target: { sheetId: sheet.id, anchor: { row: 5, column: 0 } },
      fieldCatalog: { fields: [] },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
    };
    const catalog = getPivotFieldCatalog(workbook, pivot, formula);
    pivot.fieldCatalog = catalog;
    pivot.layout.rows = [{ fieldId: catalog.fields[0]!.fieldId }];
    pivot.layout.values = [{ valueId: `value:${catalog.fields[0]!.fieldId}`, fieldId: catalog.fields[0]!.fieldId, summarizeBy: 'count' }];
    assert.equal(getPivotSourceRanges(workbook, pivot, formula)[0]?.endRow, 2);
    assert.deepEqual(computePivotResult(workbook, pivot, formula).rows.map((node) => node.key), [1, 2]);
    formula.setFormula({ sheetId: sheet.id, row: 0, column: 0 }, '=SEQUENCE(4,1,0,1)');
    assert.equal(getPivotSourceRanges(workbook, pivot, formula)[0]?.endRow, 3);
    assert.deepEqual(computePivotResult(workbook, pivot, formula).rows.map((node) => node.key), [1, 2, 3]);
  });

  it('resolves colliding workbook and worksheet names through the source sheet scope', () => {
    const workbook = new WorkbookModel('pivot-named-scope', 'Pivot Named Scope');
    const workbookSheet = workbook.getSheet('sheet-1');
    const localSheet = workbook.addSheet('sheet-2', 'Sheet2');
    [['Region', 'Amount'], ['East', 10], ['West', 20]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => workbookSheet.cells.set(rowIndex, columnIndex, { value })));
    [['Region', 'Amount'], ['North', 7], ['South', 8]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => localSheet.cells.set(rowIndex, columnIndex, { value })));
    workbook.setDefinedName({ name: 'SharedSource', formula: "='Sheet1'!A1:B3", scope: 'workbook' });
    workbook.setDefinedName({ name: 'SharedSource', formula: "='Sheet2'!A1:B3", scope: 'sheet', sheetId: localSheet.id });

    const createPivot = (id: string, source: PivotModel['source']): PivotModel => ({
      schema: 'PivotDefinition',
      id,
      source,
      target: { sheetId: workbookSheet.id, anchor: { row: 8, column: 0 } },
      fieldCatalog: { fields: [] },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
    });
    const globalPivot = createPivot('pivot-global-scope', { kind: 'named-range', name: 'SharedSource' });
    const localPivot = createPivot('pivot-local-scope', { kind: 'named-range', name: 'SharedSource', sheetId: localSheet.id });
    for (const pivot of [globalPivot, localPivot]) {
      pivot.fieldCatalog = getPivotFieldCatalog(workbook, pivot);
      const region = pivot.fieldCatalog.fields.find((field) => field.name === 'Region')!;
      const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
      pivot.layout.rows = [{ fieldId: region.fieldId }];
      pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    }

    assert.deepEqual(computePivotResult(workbook, globalPivot).rows.map((node) => node.key), ['East', 'West']);
    assert.deepEqual(computePivotResult(workbook, localPivot).rows.map((node) => node.key), ['North', 'South']);
  });

  it('requires explicit worksheet scope for local named-range Pivot sources', () => {
    const workbook = new WorkbookModel('pivot-local-name', 'Pivot Local Name');
    const local = workbook.addSheet('sheet-2', 'Sheet 2');
    workbook.setDefinedName({ name: 'SalesData', formula: '=Sheet1!A1:B2', scope: 'sheet', sheetId: 'sheet-1' });
    workbook.setDefinedName({ name: 'SalesData', formula: '=C1:D2', scope: 'sheet', sheetId: local.id });
    workbook.setDefinedName({ name: 'WorkbookOnly', formula: '=Sheet1!A1:B2', scope: 'workbook' });
    const implicitLocal: PivotModel = {
      schema: 'PivotDefinition',
      id: 'pivot-implicit-local',
      source: { kind: 'named-range', name: 'SalesData' },
      target: { sheetId: 'sheet-1', anchor: { row: 5, column: 0 } },
      fieldCatalog: { fields: [] },
      refreshPolicy: { mode: 'on-change', preserveFormatting: true, refreshOnLoad: true },
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [], subtotalLocation: 'bottom', showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' },
    };
    assert.throws(() => getPivotSourceRanges(workbook, implicitLocal), /Unknown named range/);

    const scoped: PivotModel = { ...implicitLocal, id: 'pivot-scoped-local', source: { kind: 'named-range', name: 'SalesData', sheetId: local.id } };
    assert.deepEqual(getPivotSourceRanges(workbook, scoped)[0], { sheetId: local.id, startRow: 0, endRow: 1, startColumn: 2, endColumn: 3 });

    const incorrectlyScopedGlobal: PivotModel = { ...implicitLocal, id: 'pivot-incorrectly-scoped-global', source: { kind: 'named-range', name: 'WorkbookOnly', sheetId: local.id } };
    assert.throws(() => getPivotSourceRanges(workbook, incorrectlyScopedGlobal), /Unknown named range/);
  });

  it('fails closed when a Pivot source intersects a blocked spill', () => {
    const workbook = new WorkbookModel('pivot-blocked-spill', 'Pivot Blocked Spill');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 'Member' });
    sheet.cells.set(1, 0, { value: 'occupied' });
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-blocked-spill', { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 });
    assert.ok(pivot);
    sheet.cells.set(0, 0, { value: null, formula: '=SEQUENCE(3,1,1,1)' });
    const formula = new FormulaEngine({ defaultSheetId: sheet.id });
    formula.setSpillEnvironment(sheet.id, createSpillEnvironment(sheet));
    formula.setFormula({ sheetId: sheet.id, row: 0, column: 0 }, '=SEQUENCE(3,1,1,1)');
    assert.throws(() => computePivotResult(workbook, pivot, formula), /blocked spill/i);
  });

  it('keeps the canonical blank member in the field catalogue and manual filters', () => {
    const workbook = new WorkbookModel('pivot-blank-member', 'Pivot Blank Member');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 'Region' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: 'North' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.cells.set(2, 0, { value: null });
    sheet.cells.set(2, 1, { value: 20 });
    sheet.cells.set(3, 0, { value: '' });
    sheet.cells.set(3, 1, { value: 30 });
    sheet.cells.set(4, 0, { value: 'South' });
    sheet.cells.set(4, 1, { value: 40 });
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-blank-member', { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    assert.deepEqual(region.values, ['North', null, 'South']);
    pivot.fieldCatalog = catalog;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const all = computePivotResult(workbook, pivot);
    assert.deepEqual(all.rows.map((node) => node.label), ['(blank)', 'North', 'South']);
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', fieldId: region.fieldId, mode: 'include', memberKeys: [{ type: 'blank', value: null }] }];
    const blankOnly = computePivotResult(workbook, pivot);
    assert.deepEqual(blankOnly.rows.map((node) => node.label), ['(blank)']);
    assert.equal(blankOnly.rows[0]?.values[0]?.values[0], 50);
  });

  it('keeps distinct FormulaEngine errors as Pivot members and aggregates them explicitly', () => {
    const workbook = new WorkbookModel('pivot-errors', 'Pivot Errors');
    const sheet = workbook.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 'Member' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(1, 0, { value: null, formula: '=1/0' });
    sheet.cells.set(1, 1, { value: 10 });
    sheet.cells.set(2, 0, { value: null, formula: '=VALUE("bad")' });
    sheet.cells.set(2, 1, { value: 20 });
    const formula = new FormulaEngine({ defaultSheetId: sheet.id });
    formula.setSpillEnvironment(sheet.id, createSpillEnvironment(sheet));
    formula.setFormula({ sheetId: sheet.id, row: 1, column: 0 }, '=1/0');
    formula.setFormula({ sheetId: sheet.id, row: 2, column: 0 }, '=VALUE("bad")');
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-errors', { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot, formula);
    const member = catalog.fields.find((field) => field.name === 'Member')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.fieldCatalog = catalog;
    pivot.layout.rows = [{ fieldId: member.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const result = computePivotResult(workbook, pivot, formula);
    assert.deepEqual(result.rows.map((node) => node.label), ['#DIV/0!', '#VALUE!']);
    assert.equal(result.grandTotal?.values[0], 30);
    const errorRows = [
      { values: { value: { kind: 'error' as const, code: '#DIV/0!' as const } } },
      { values: { value: { kind: 'error' as const, code: '#N/A' as const } } },
      { values: { value: 2 } },
    ];
    assert.equal(aggregatePivotValues(errorRows, 'value', 'count'), 3);
    assert.equal(aggregatePivotValues(errorRows, 'value', 'distinct-count'), 3);
    assert.deepEqual(aggregatePivotValues(errorRows, 'value', 'sum'), { kind: 'error', code: '#DIV/0!' });
  });

  it('orders text members from persisted collation, independent of host defaults', () => {
    const workbook = new WorkbookModel('pivot-collation', 'Pivot Collation');
    const sheet = workbook.getSheet('sheet-1');
    [['Member', 'Amount'], ['é', 1], ['e', 2], ['中', 3], ['甲', 4], ['A', 5], ['a', 6]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-collation', { sheetId: 'sheet-1', startRow: 0, endRow: 6, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const member = catalog.fields.find((field) => field.name === 'Member')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: member.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const first = computePivotResult(workbook, pivot);
    const second = computePivotResult(workbook, pivot);
    assert.deepEqual(first.rows.map((node) => node.label), second.rows.map((node) => node.label));
    const alternate = structuredClone(pivot);
    alternate.layout.collation = { locale: 'zh-CN', sensitivity: 'variant', numeric: false, caseFirst: 'false' };
    assert.notDeepEqual(first.rows.map((node) => node.label), computePivotResult(workbook, alternate).rows.map((node) => node.label));
  });

  it('rejects an unsupported persisted collation before calculation', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-invalid-collation', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    pivot.layout.collation = { locale: '***', sensitivity: 'variant', numeric: false, caseFirst: 'false' };
    assert.throws(() => computePivotResult(workbook, pivot), /unsupported|invalid/i);
  });

  it('builds a complete canonical definition with stable field IDs', () => {
    const workbook = workbookWithData();
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    const pivot = buildPivotModel(workbook, 'sheet-1', 'canonical-pivot', sourceRange);
    assert.ok(pivot);
    assert.equal(pivot.schema, 'PivotDefinition');
    assert.deepEqual(pivot.source, { kind: 'worksheet-range', range: sourceRange });
    assert.equal(pivot.target.sheetId, 'sheet-1');
    assert.ok(pivot.fieldCatalog.fields.every((field) => field.fieldId.length > 0));
  });

  it('evaluates calculated fields against each Pivot summary bucket instead of source records', () => {
    const workbook = new WorkbookModel('pivot-calculated-summary', 'Pivot Calculated Summary');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Sales', 'Units'], ['East', 10, 2], ['East', 20, 3], ['West', 5, 4]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-calculated-summary', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const calculatedId = 'calculated:revenue';
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ valueId: `value:${calculatedId}`, fieldId: calculatedId, summarizeBy: 'sum' }];
    pivot.layout.calculatedFields = [{ fieldId: calculatedId, name: 'Revenue', formula: '=Sales*Units' }];

    const result = computePivotResult(workbook, pivot);
    assert.deepEqual(result.rows.map((node) => [node.label, node.values[0]?.values[0]]), [['East', 150], ['West', 20]]);
    assert.equal(result.grandTotal?.values[0], 315);
  });

  it('resolves calculated-field dependencies in summary order and supports bracketed field references', () => {
    const workbook = new WorkbookModel('pivot-calculated-dependencies', 'Pivot Calculated Dependencies');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Sales', 'Units'], ['East', 10, 2], ['East', 20, 3]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-calculated-dependencies', { sheetId: 'sheet-1', startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const revenue = 'calculated:revenue';
    const adjusted = 'calculated:adjusted';
    pivot.layout.values = [{ valueId: `value:${adjusted}`, fieldId: adjusted, summarizeBy: 'sum' }, { valueId: `value:${revenue}`, fieldId: revenue, summarizeBy: 'sum' }];
    pivot.layout.calculatedFields = [
      { fieldId: adjusted, name: 'Adjusted', formula: '=[Revenue]+10' },
      { fieldId: revenue, name: 'Revenue', formula: '=[Sales]*[Units]' },
    ];
    const result = computePivotResult(workbook, pivot);
    assert.deepEqual(result.grandTotal?.values, [160, 150]);
  });

  it('rejects unknown, cyclic, and structural calculated-field references before producing a result', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-invalid-calculated', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    const calculated = 'calculated:invalid';
    pivot.layout.values = [{ valueId: `value:${calculated}`, fieldId: calculated, summarizeBy: 'sum' }];
    pivot.layout.calculatedFields = [{ fieldId: calculated, name: 'Invalid', formula: '=Missing+1' }];
    assert.throws(() => computePivotResult(workbook, pivot), /unknown field/i);

    pivot.layout.calculatedFields = [
      { fieldId: 'calculated:a', name: 'A', formula: '=B+1' },
      { fieldId: 'calculated:b', name: 'B', formula: '=A+1' },
    ];
    pivot.layout.values = [{ valueId: `value:${'calculated:a'}`, fieldId: 'calculated:a', summarizeBy: 'sum' }];
    assert.throws(() => computePivotResult(workbook, pivot), /dependency cycle/i);

    pivot.layout.calculatedFields = [{ fieldId: calculated, name: 'Invalid', formula: `=${amount.fieldId}+1` }];
    pivot.layout.rows = [{ fieldId: calculated }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    assert.throws(() => computePivotResult(workbook, pivot), /only valid in Values/i);
  });

  it('plans star joins by source identity, is invariant to source order, and preserves provenance', () => {
    const workbook = relationalWorkbook();
    const first = relationalPivot(workbook, ['orders', 'customers', 'products']);
    const reordered = relationalPivot(workbook, ['products', 'orders', 'customers']);
    const firstResult = computePivotResult(workbook, first);
    const reorderedResult = computePivotResult(workbook, reordered);
    assert.deepEqual(firstResult.rows.map((node) => [node.label, node.values[0]?.values]), reorderedResult.rows.map((node) => [node.label, node.values[0]?.values]));
    const paths = firstResult.rows[0]?.values[0]?.sourceRowPaths ?? [];
    assert.deepEqual([...new Set(paths.map((path) => path.sourceId))].sort(), ['customers', 'orders', 'products']);
  });

  it('rejects duplicate lookup keys and incompatible relationship key types before aggregation', () => {
    const duplicate = relationalWorkbook();
    duplicate.getSheet('sheet-1').cells.set(2, 4, { value: 'c1' });
    duplicate.getSheet('sheet-1').cells.set(2, 5, { value: 'Duplicate' });
    assert.throws(() => relationalPivot(duplicate, ['orders', 'customers', 'products']), /lookup key is not unique/);
    const incompatible = relationalWorkbook();
    incompatible.getSheet('sheet-1').cells.set(1, 4, { value: 1 });
    assert.throws(() => relationalPivot(incompatible, ['orders', 'customers', 'products']), /key types are incompatible/);
  });

  it('keeps typed members distinct and treats manual all as no filter', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-typed', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const region = getPivotFieldCatalog(workbook, pivot).fields.find((field) => field.name === 'Region')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    const amountFieldId = getPivotFieldCatalog(workbook, pivot).fields.find((field) => field.name === 'Amount')!.fieldId;
    pivot.layout.values = [{ valueId: `value:${amountFieldId}`, fieldId: amountFieldId, summarizeBy: 'count' }];
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', fieldId: region.fieldId, mode: 'all', memberKeys: [] }];
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 3);
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', fieldId: region.fieldId, mode: 'include', memberKeys: [{ type: 'text', value: 'East' }] }];
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 2);
    assert.notDeepEqual({ type: 'text', value: '1' }, { type: 'number', value: 1 });
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', scope: 'field', fieldId: region.fieldId, mode: 'include', memberKeys: [{ type: 'text', value: 'East' }] }];
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 8, column: 0 } };
    const fieldFiltered = buildPivotGridProjection(workbook, pivot);
    assert.equal(fieldFiltered.cells.some((cell) => cell.kind === 'filter'), false);
    assert.equal(computePivotResult(workbook, pivot).grandTotal?.values[0], 2);
  });

  it('projects one typed report summary per field and never hides active families as All', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-report-summary', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.allowMultipleFiltersPerField = true;
    pivot.layout.filters = [
      { kind: 'manual', family: 'manual', scope: 'report', fieldId: region.fieldId, mode: 'all', memberKeys: [] },
      { kind: 'condition', family: 'label', fieldId: region.fieldId, scope: 'report', operator: 'begins-with', value: 'E' },
      { kind: 'top-items', family: 'top-items', scope: 'report', fieldId: region.fieldId, valueId: `value:${amount.fieldId}`, direction: 'top', mode: 'items', threshold: 1 },
    ];
    const summary = summarizePivotReportFilters(pivot.layout.filters, catalog, region.fieldId, pivot.layout.values);
    assert.equal(summary.fieldName, 'Region');
    assert.equal(summary.active, true);
    assert.equal(summary.entries.length, 3);
    assert.equal(summary.entries[0]?.kind, 'manual');
    assert.equal(summary.entries[1]?.kind, 'condition');
    assert.equal(summary.entries[2]?.kind, 'top-items');
    const projection = buildPivotGridProjection(workbook, pivot);
    const filterCells = projection.cells.filter((cell) => cell.kind === 'filter');
    assert.equal(filterCells.length, 1);
    assert.equal(filterCells[0]?.filterSummary?.active, true);
    assert.equal(filterCells[0]?.filterSummary?.entries.length, 3);
  });

  it('ranks top/bottom filters by the selected Values placement and explicit mode', () => {
    const workbook = new WorkbookModel('pivot-top-bottom', 'Pivot Top Bottom');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Amount'], ['East', 100], ['East', 100], ['West', 150], ['South', 40]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-top-bottom', { sheetId: 'sheet-1', startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    const filter = (valueId: string, mode: 'items' | 'percent' | 'sum', threshold: number, direction: 'top' | 'bottom' = 'top') => ({
      kind: 'top-items' as const, family: 'top-items' as const, fieldId: region.fieldId, scope: 'field' as const, valueId, mode, threshold, direction,
    });
    const names = (summarizeBy: 'sum' | 'count' | 'average' | 'min' | 'max', valueId = 'amount:value') => {
      pivot.layout.values = [{ valueId, fieldId: amount.fieldId, summarizeBy }];
      pivot.layout.filters = [filter(valueId, 'items', 1)];
      return computePivotResult(workbook, pivot).rows.map((node) => node.label);
    };
    assert.deepEqual(names('sum'), ['East']);
    assert.deepEqual(names('average'), ['West']);
    assert.deepEqual(names('count', 'amount:count'), ['East']);
    assert.deepEqual(names('min'), ['West']);
    assert.deepEqual(names('max'), ['West']);

    pivot.layout.values = [{ valueId: 'amount:value', fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.filters = [filter('amount:value', 'percent', 50)];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['East']);
    pivot.layout.filters = [filter('amount:value', 'sum', 300)];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['East', 'West']);
    pivot.layout.filters = [filter('amount:value', 'sum', 100, 'bottom')];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['South', 'West']);

    pivot.layout.values = [
      { valueId: 'amount:sum', fieldId: amount.fieldId, summarizeBy: 'sum' },
      { valueId: 'amount:count', fieldId: amount.fieldId, summarizeBy: 'count' },
    ];
    pivot.layout.filters = [filter('amount:count', 'items', 1)];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['East']);

    const invalidLegacy = { ...filter('amount:sum', 'items', 1), count: 1, valueFieldId: amount.fieldId } as unknown as PivotModel['layout']['filters'][number];
    pivot.layout.filters = [invalidLegacy];
    assert.throws(() => computePivotResult(workbook, pivot), /removed top-items count|unknown Values placement|threshold/i);
  });

  it('projects each value field with its canonical number format', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-number-format', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum', numberFormat: '#,##0.00' }];
    const projection = buildPivotGridProjection(workbook, pivot);
    const values = projection.cells.filter((cell) => cell.kind === 'value' || cell.kind === 'grand-total');
    assert.ok(values.length > 0);
    assert.equal(values[0]?.text, '15.00');
    assert.equal(values[0]?.numberFormat, '#,##0.00');
    assert.equal(values.at(-1)?.text, '35.00');
    assert.throws(() => {
      pivot.layout.values[0]!.numberFormat = '[Red';
      buildPivotGridProjection(workbook, pivot);
    }, /unterminated/);
  });

  it('applies canonical date, numeric and manual grouping before axis aggregation', () => {
    const workbook = new WorkbookModel('pivot-grouping', 'Pivot Grouping');
    const sheet = workbook.getSheet('sheet-1');
    [['Date', 'Amount', 'Category'], [45292, 10, 'A'], [45323, 20, 'B'], [45657, 30, 'C']].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-grouping', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const date = catalog.fields.find((field) => field.name === 'Date')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    const category = catalog.fields.find((field) => field.name === 'Category')!;
    pivot.fieldCatalog.fields[date.ordinal]!.dataType = 'date';
    pivot.layout.rows = [{ fieldId: date.fieldId, group: { kind: 'date', unit: 'year' } }, { fieldId: category.fieldId, group: { kind: 'manual', groups: [{ groupId: 'ab', name: 'AB', items: [{ type: 'text', value: 'A' }, { type: 'text', value: 'B' }] }] } }];
    pivot.layout.columns = [{ fieldId: amount.fieldId, group: { kind: 'number', interval: 10, start: 0 } }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const result = computePivotResult(workbook, pivot);
    assert.equal(result.rows[0]?.label, '2024');
    assert.equal(result.rows[0]?.children[0]?.label, 'AB');
    assert.equal(result.columnPaths[0]?.[0], 10);
    assert.equal(result.rows[0]?.children[0]?.values[0]?.values[0], 10);
  });

  it('keeps multi-level date grouping and custom numeric bounds in the canonical result', () => {
    const workbook = new WorkbookModel('pivot-grouping-levels', 'Pivot Grouping Levels');
    const sheet = workbook.getSheet('sheet-1');
    [['Date', 'Amount'], [45292, 25], [45323, 75], [45657, 125]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-grouping-levels', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const date = pivot.fieldCatalog.fields.find((field) => field.name === 'Date')!;
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
    pivot.fieldCatalog.fields[date.ordinal]!.dataType = 'date';
    pivot.layout.rows = [{ fieldId: date.fieldId, group: { kind: 'date', unit: 'year', units: ['year', 'quarter', 'month'] } }];
    pivot.layout.columns = [{ fieldId: amount.fieldId, group: { kind: 'number', interval: 50, start: 0, end: 1000 } }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const result = computePivotResult(workbook, pivot);
    assert.equal(result.rows[0]?.label, '2024 / 2024 Q1 / 2024-01');
    assert.deepEqual(result.columnPaths, [[0], [50], [100]]);
  });

  it('uses grouped members as the field-filter domain and keeps manual group identities stable', () => {
    const workbook = new WorkbookModel('pivot-grouped-filter', 'Pivot Grouped Filter');
    const sheet = workbook.getSheet('sheet-1');
    [['Date', 'Category', 'Amount'], [45292, 'A', 10], [45323, 'B', 20], [45658, 'C', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-grouped-filter', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const date = pivot.fieldCatalog.fields.find((field) => field.name === 'Date')!;
    const category = pivot.fieldCatalog.fields.find((field) => field.name === 'Category')!;
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
    pivot.fieldCatalog.fields[date.ordinal]!.dataType = 'date';
    const yearGroup = { kind: 'date' as const, unit: 'year' as const };
    const yearMembers = buildPivotGroupedFilterMembers(date.values ?? [], yearGroup);
    assert.deepEqual(yearMembers.map((member) => member.value), [2024, 2025]);
    pivot.layout.rows = [{ fieldId: date.fieldId, group: yearGroup }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', scope: 'field', fieldId: date.fieldId, mode: 'include', memberKeys: [yearMembers[0]!.key] }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2024']);
    pivot.layout.filters = [{ kind: 'condition', family: 'date', scope: 'field', fieldId: date.fieldId, operator: 'equals', value: '2024-01-01' }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2024']);

    const manualGroup = { kind: 'manual' as const, groups: [{ groupId: 'category-ab', name: 'AB', items: [{ type: 'text' as const, value: 'A' }, { type: 'text' as const, value: 'B' }] }] };
    const categoryMembers = buildPivotGroupedFilterMembers(category.values ?? [], manualGroup);
    assert.equal(categoryMembers.find((member) => member.label === 'AB')?.key.value, '__pivot_group__:category-ab');
    pivot.layout.rows = [{ fieldId: category.fieldId, group: manualGroup }];
    pivot.layout.filters = [{ kind: 'manual', family: 'manual', scope: 'field', fieldId: category.fieldId, mode: 'include', memberKeys: [categoryMembers.find((member) => member.label === 'AB')!.key] }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['AB']);
    pivot.layout.rows[0] = { fieldId: category.fieldId, group: { kind: 'manual', groups: [{ ...manualGroup.groups[0]!, name: 'Renamed' }] } };
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['Renamed']);
    pivot.layout.rows[0] = { fieldId: category.fieldId, group: { kind: 'manual', groups: [{ groupId: 'category-new', name: 'New', items: [{ type: 'text', value: 'A' }, { type: 'text', value: 'B' }] }] } };
    assert.throws(() => computePivotResult(workbook, pivot), /incompatible with grouping/);
  });

  it('keeps numeric label sorting separate from explicit Values sorting', () => {
    const workbook = new WorkbookModel('pivot-numeric-sort', 'Pivot Numeric Sort');
    const sheet = workbook.getSheet('sheet-1');
    [['Year', 'Sales'], [2024, 500], [2025, 100], [2026, 900]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-numeric-sort', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    const year = pivot.fieldCatalog.fields.find((field) => field.name === 'Year')!;
    const sales = pivot.fieldCatalog.fields.find((field) => field.name === 'Sales')!;
    pivot.layout.rows = [{ fieldId: year.fieldId, sort: { direction: 'ascending', by: 'label' } }];
    pivot.layout.values = [{ valueId: `value:${sales.fieldId}`, fieldId: sales.fieldId, summarizeBy: 'sum' }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2024', '2025', '2026']);

    pivot.layout.rows[0] = { fieldId: year.fieldId, sort: { direction: 'ascending', by: 'value', valueId: `value:${sales.fieldId}` } };
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2025', '2024', '2026']);
    pivot.layout.rows[0] = { fieldId: year.fieldId, sort: { direction: 'ascending', by: 'value' } };
    assert.throws(() => computePivotResult(workbook, pivot), /requires a valueId/);
  });

  it('evaluates typed label and date filter families with range predicates', () => {
    const workbook = new WorkbookModel('pivot-filter-families', 'Pivot Filter Families');
    const sheet = workbook.getSheet('sheet-1');
    [['Name', 'Date', 'Amount'], ['Alice', '2024-01-10', 10], ['Bob', '2024-06-10', 20], ['Avery', '2025-01-10', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-filter-families', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const name = pivot.fieldCatalog.fields.find((field) => field.name === 'Name')!;
    const date = pivot.fieldCatalog.fields.find((field) => field.name === 'Date')!;
    const amount = pivot.fieldCatalog.fields.find((field) => field.name === 'Amount')!;
    pivot.fieldCatalog.fields[date.ordinal]!.dataType = 'date';
    pivot.layout.rows = [{ fieldId: name.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.filters = [{ kind: 'condition', family: 'label', fieldId: name.fieldId, operator: 'begins-with', value: 'A' }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['Alice', 'Avery']);
    pivot.layout.rows = [{ fieldId: date.fieldId }];
    pivot.layout.filters = [{ kind: 'condition', family: 'date', fieldId: date.fieldId, operator: 'between', value: '2024-01-01', value2: '2024-12-31' }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2024-01-10', '2024-06-10']);
  });

  it('keeps subtotal ownership per row field and expands custom subtotal functions', () => {
    const workbook = new WorkbookModel('pivot-subtotals', 'Pivot Subtotals');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'City', 'Amount'], ['East', 'Boston', 10], ['East', 'Austin', 20], ['West', 'Seattle', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-subtotals', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const city = catalog.fields.find((field) => field.name === 'City')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [
      { fieldId: region.fieldId, subtotal: { mode: 'none' } },
      { fieldId: city.fieldId, subtotal: { mode: 'custom', functions: ['sum', 'average'] } },
    ];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.subtotalLocation = 'bottom';
    const result = computePivotResult(workbook, pivot);
    const east = result.rows.find((node) => node.label === 'East')!;
    assert.equal(east.subtotal, false);
    assert.equal(east.children[0]?.subtotal, false);
    assert.equal(result.valueFields?.length, 2);
    assert.deepEqual(east.children.find((node) => node.label === 'Boston')?.values[0]?.values, [10, 10]);
    const projection = buildPivotGridProjection(workbook, pivot, result);
    const visibleSubtotalCells = projection.cells.filter((cell) => cell.kind === 'subtotal');
    assert.equal(visibleSubtotalCells.length, 0);
    pivot.layout.rows[0] = { fieldId: region.fieldId, subtotal: { mode: 'custom', functions: ['sum', 'average'] } };
    const withOuterSubtotal = computePivotResult(workbook, pivot);
    const eastSubtotal = withOuterSubtotal.rows.find((node) => node.label === 'East');
    assert.equal(eastSubtotal?.subtotal, true);
    assert.equal(withOuterSubtotal.valueFields?.length, 3);
    assert.deepEqual(eastSubtotal?.values[0]?.values, [30, 15, 30]);
  });

  it('applies grand percentages to detail, subtotal, and grand-total cells from one raw matrix', () => {
    const workbook = new WorkbookModel('pivot-show-as-grand', 'Pivot Show As Grand');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'City', 'Amount'], ['East', 'Boston', 10], ['East', 'Austin', 20], ['West', 'Seattle', 40]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-show-as-grand', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const city = catalog.fields.find((field) => field.name === 'City')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }, { fieldId: city.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum', showAs: { kind: 'grand-percentage' } }];
    const result = computePivotResult(workbook, pivot);
    const east = result.rows.find((node) => node.label === 'East')!;
    const west = result.rows.find((node) => node.label === 'West')!;
    assert.equal(east.values[0]?.values[0], 30 / 70);
    assert.equal(east.children.find((node) => node.label === 'Boston')?.values[0]?.values[0], 10 / 70);
    assert.equal(east.children.find((node) => node.label === 'Austin')?.values[0]?.values[0], 20 / 70);
    assert.equal(west.values[0]?.values[0], 40 / 70);
    assert.equal(result.grandTotal?.values[0], 1);
  });

  it('uses subtotal peers for running total and rank instead of indexing a leaf-only series', () => {
    const workbook = new WorkbookModel('pivot-show-as-subtotals', 'Pivot Show As Subtotals');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'City', 'Amount'], ['East', 'Boston', 10], ['East', 'Austin', 20], ['West', 'Seattle', 40]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-show-as-subtotals', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const city = catalog.fields.find((field) => field.name === 'City')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }, { fieldId: city.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum', showAs: { kind: 'running-total', axis: 'row' } }];
    const running = computePivotResult(workbook, pivot);
    const eastRunning = running.rows.find((node) => node.label === 'East')!;
    const westRunning = running.rows.find((node) => node.label === 'West')!;
    assert.equal(eastRunning.values[0]?.values[0], 30);
    assert.equal(westRunning.values[0]?.values[0], 70);
    assert.equal(eastRunning.children.find((node) => node.label === 'Austin')?.values[0]?.values[0], 20);
    assert.equal(eastRunning.children.find((node) => node.label === 'Boston')?.values[0]?.values[0], 30);
    assert.equal(running.grandTotal?.values[0], 70);

    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum', showAs: { kind: 'rank', axis: 'row', direction: 'descending' } }];
    const ranked = computePivotResult(workbook, pivot);
    const eastRank = ranked.rows.find((node) => node.label === 'East')!;
    const westRank = ranked.rows.find((node) => node.label === 'West')!;
    assert.equal(eastRank.values[0]?.values[0], 2);
    assert.equal(westRank.values[0]?.values[0], 1);
    assert.equal(eastRank.children.find((node) => node.label === 'Austin')?.values[0]?.values[0], 2);
    assert.equal(eastRank.children.find((node) => node.label === 'Boston')?.values[0]?.values[0], 3);
    assert.equal(ranked.grandTotal?.values[0], 1);
    assert.equal(ranked.rows.flatMap((node) => [node, ...node.children]).flatMap((node) => node.values.flatMap((cell) => cell.values)).some((value) => value === 0), false);
  });

  it('keeps a root data row for Columns plus Values when Rows is empty', () => {
    const workbook = new WorkbookModel('pivot-columns-only', 'Pivot Columns Only');
    const sheet = workbook.getSheet('sheet-1');
    [['Month', 'Sales'], ['Jan', 10], ['Feb', 20], ['Mar', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-columns-only', sourceRange);
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const month = catalog.fields.find((field) => field.name === 'Month')!;
    const sales = catalog.fields.find((field) => field.name === 'Sales')!;
    pivot.layout.rows = [];
    pivot.layout.columns = [{ fieldId: month.fieldId }];
    pivot.layout.values = [{ valueId: `value:${sales.fieldId}`, fieldId: sales.fieldId, summarizeBy: 'sum' }];
    pivot.layout.showRowGrandTotals = true;
    pivot.layout.showColumnGrandTotals = true;
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 6, column: 0 } };

    const tree = computePivotResult(workbook, pivot);
    assert.equal(tree.rows.length, 1);
    assert.equal(tree.rows[0]?.nodeId, '__root__');
    assert.deepEqual(Object.fromEntries(tree.columnPaths.map((path, index) => [path[0], tree.rows[0]?.values[index]?.values[0]])), { Jan: 10, Feb: 20, Mar: 30 });
    assert.equal(tree.grandTotal?.values[0], 60);

    const projection = buildPivotGridProjection(workbook, pivot, tree);
    const values = projection.cells.filter((cell) => cell.kind === 'value');
    assert.deepEqual(Object.fromEntries(values.map((cell) => [cell.columnPath?.[0], cell.value])), { Jan: 10, Feb: 20, Mar: 30 });
    assert.equal(projection.cells.some((cell) => cell.text === 'Jan Sales'), true);
    assert.equal(projection.cells.some((cell) => cell.text === 'Grand Total'), true);
  });

  it('renders row and column grand totals independently and rejects the removed shared flag', () => {
    const workbook = new WorkbookModel('pivot-grand-total-axes', 'Pivot Grand Total Axes');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Month', 'Amount'], ['East', 'Jan', 10], ['East', 'Feb', 20], ['West', 'Jan', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 };
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-grand-total-axes', sourceRange);
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const month = catalog.fields.find((field) => field.name === 'Month')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.columns = [{ fieldId: month.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.showRowGrandTotals = true;
    pivot.layout.showColumnGrandTotals = false;
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 6, column: 0 } };

    const rowOnly = buildPivotGridProjection(workbook, pivot, computePivotResult(workbook, pivot));
    assert.equal(rowOnly.cells.some((cell) => cell.kind === 'grand-total' && cell.row === 0), false);
    assert.equal(rowOnly.cells.filter((cell) => cell.kind === 'grand-total').length, 2);
    assert.equal(rowOnly.cells.some((cell) => cell.resultCellId?.endsWith('|grand-total:row')), true);

    pivot.layout.showRowGrandTotals = false;
    pivot.layout.showColumnGrandTotals = true;
    const columnOnly = buildPivotGridProjection(workbook, pivot, computePivotResult(workbook, pivot));
    assert.equal(columnOnly.cells.filter((cell) => cell.kind === 'grand-total').length, 3);
    assert.equal(columnOnly.cells.filter((cell) => cell.kind === 'grand-total').some((cell) => cell.nodeId), false);

    pivot.layout.showColumnGrandTotals = false;
    const neither = buildPivotGridProjection(workbook, pivot, computePivotResult(workbook, pivot));
    assert.equal(neither.cells.some((cell) => cell.kind === 'grand-total'), false);
  });

  it('keeps a collapsed parent visible while hiding only its descendants', () => {
    const workbook = new WorkbookModel('pivot-expansion', 'Pivot Expansion');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'City', 'Amount'], ['East', 'Boston', 10], ['East', 'Austin', 20], ['West', 'Seattle', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const sourceRange = { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 };
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-expansion', sourceRange);
    assert.ok(pivot);
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const city = catalog.fields.find((field) => field.name === 'City')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }, { fieldId: city.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 6, column: 0 } };
    const tree = computePivotResult(workbook, pivot);
    const parent = tree.rows.find((node) => node.label === 'East')!;
    const child = parent.children.find((node) => node.label === 'Boston')!;
    pivot.layout.expansion = { expandedNodeIds: [], collapsedNodeIds: [parent.nodeId!], showButtons: true };
    const collapsed = buildPivotGridProjection(workbook, pivot, tree);
    assert.equal(collapsed.cells.some((cell) => cell.nodeId === parent.nodeId && cell.kind === 'expand-toggle' && cell.expanded === false), true);
    assert.equal(collapsed.cells.some((cell) => cell.nodeId === child.nodeId), false);
    pivot.layout.expansion = { expandedNodeIds: [parent.nodeId!], collapsedNodeIds: [], showButtons: true };
    const expanded = buildPivotGridProjection(workbook, pivot, tree);
    assert.equal(expanded.cells.some((cell) => cell.nodeId === child.nodeId), true);
  });

  it('implements each aggregate independently', () => {
    const rows = [{ values: { value: 2 } }, { values: { value: 4 } }, { values: { value: 4 } }, { values: { value: null } }];
    assert.equal(aggregatePivotValues(rows, 'value', 'sum'), 10);
    assert.equal(aggregatePivotValues(rows, 'value', 'count'), 3);
    assert.equal(aggregatePivotValues(rows, 'value', 'count-numbers'), 3);
    assert.equal(aggregatePivotValues(rows, 'value', 'average'), 10 / 3);
    assert.equal(aggregatePivotValues(rows, 'value', 'min'), 2);
    assert.equal(aggregatePivotValues(rows, 'value', 'max'), 4);
    assert.equal(aggregatePivotValues(rows, 'value', 'product'), 32);
    assert.equal(aggregatePivotValues(rows, 'value', 'distinct-count'), 2);
    assert.equal(aggregatePivotValues(rows, 'value', 'varp'), 8 / 9);
  });

  it('keeps numeric-looking text out of every numeric summary', () => {
    const rows = [
      { values: { value: 10 } },
      { values: { value: '10' } },
      { values: { value: '$100' } },
      { values: { value: '50%' } },
      { values: { value: 0.5 } },
      { values: { value: true } },
      { values: { value: null } },
    ];
    assert.equal(aggregatePivotValues(rows, 'value', 'count'), 6);
    assert.equal(aggregatePivotValues(rows, 'value', 'count-numbers'), 2);
    assert.equal(aggregatePivotValues(rows, 'value', 'sum'), 10.5);
    assert.equal(aggregatePivotValues(rows, 'value', 'average'), 5.25);
    assert.equal(aggregatePivotValues(rows, 'value', 'min'), 0.5);
    assert.equal(aggregatePivotValues(rows, 'value', 'max'), 10);
    assert.equal(aggregatePivotValues(rows, 'value', 'product'), 5);
    assert.equal(aggregatePivotValues(rows, 'value', 'stdev'), Math.sqrt(45.125));
    assert.equal(aggregatePivotValues(rows, 'value', 'stdevp'), Math.sqrt(22.5625));
    assert.equal(aggregatePivotValues(rows, 'value', 'var'), 45.125);
    assert.equal(aggregatePivotValues(rows, 'value', 'varp'), 22.5625);
    assert.equal(aggregatePivotValues(rows, 'value', 'distinct-count'), 6);
    assert.equal(aggregatePivotValues([{ values: { value: '$100' } }], 'value', 'sum'), 0);
    assert.equal(aggregatePivotValues([{ values: { value: '$100' } }], 'value', 'average'), null);
  });

  it('returns a derived overlay, reports collisions, and supports hit testing without cell writeback', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-overlay', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 5, column: 0 } };
    const before = workbook.getSheet('sheet-1').cells.count();
    const first = buildPivotGridProjection(workbook, pivot);
    assert.equal(first.collision.status, 'clear');
    workbook.getSheet('sheet-1').cells.set(5, 0, { value: 'ordinary cell' });
    const projection = buildPivotGridProjection(workbook, pivot);
    assert.equal(workbook.getSheet('sheet-1').cells.count(), before + 1);
    assert.equal(projection.collision.status, 'collision');
    assert.deepEqual(projection.cells, first.cells);
    assert.deepEqual(projection.occupiedRange, first.occupiedRange);
    assert.equal(projection.schema, 'PivotGridProjection');
    const hit = hitTestPivotProjection(projection, 0, 0);
    assert.equal(hit.pivotId, pivot.id);
    assert.equal(hit.kind, 'cell');
  });

  it('reserves the complete derived footprint of every Pivot instead of only its anchor', () => {
    const workbook = workbookWithData();
    const sheet = workbook.getSheet('sheet-1');
    const first = buildPivotModel(workbook, sheet.id, 'pivot-footprint-a', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    const second = buildPivotModel(workbook, sheet.id, 'pivot-footprint-b', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(first && second);
    first.target = { sheetId: sheet.id, anchor: { row: 5, column: 0 } };
    second.target = { sheetId: sheet.id, anchor: { row: 7, column: 0 } };
    sheet.pivots.push(first);
    const firstProjection = buildPivotGridProjection(workbook, first);
    assert.equal(firstProjection.collision.status, 'clear');
    sheet.pivots.push(second);
    const collision = buildPivotGridProjection(workbook, second);
    assert.equal(collision.collision.status, 'collision');
    assert.equal(collision.collision.reasons.includes('pivot'), true);
    assert.deepEqual(collision.collision.conflicts.find((entry) => entry.participantId === first.id)?.range, firstProjection.occupiedRange);
    assert.equal(collision.collision.conflictingRanges.some((range) => range.startRow === firstProjection.occupiedRange.startRow && range.endRow === firstProjection.occupiedRange.endRow), true);
  });

  it('keeps the last-valid projection when a layout expansion would overlap another Pivot, then allows a free move', () => {
    const workbook = new WorkbookModel('pivot-footprint-refresh', 'Pivot Footprint Refresh');
    const sheet = workbook.getSheet('sheet-1');
    [['Region', 'Category', 'Amount'], ['East', 'Widget', 10], ['East', 'Gadget', 20], ['West', 'Widget', 30]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
    const first = buildPivotModel(workbook, sheet.id, 'pivot-footprint-expand', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    const second = buildPivotModel(workbook, sheet.id, 'pivot-footprint-stable', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 });
    assert.ok(first && second);
    const catalog = getPivotFieldCatalog(workbook, first);
    first.fieldCatalog = catalog;
    second.fieldCatalog = structuredClone(catalog);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const category = catalog.fields.find((field) => field.name === 'Category')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    first.layout.rows = [{ fieldId: region.fieldId }];
    first.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    second.layout.rows = [{ fieldId: region.fieldId }];
    second.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    first.target = { sheetId: sheet.id, anchor: { row: 5, column: 0 } };
    sheet.pivots.push(first);
    const initial = buildPivotGridProjection(workbook, first);
    assert.equal(initial.collision.status, 'clear');
    second.target = { sheetId: sheet.id, anchor: { row: initial.occupiedRange.endRow + 1, column: 0 } };
    sheet.pivots.push(second);
    assert.equal(buildPivotGridProjection(workbook, second).collision.status, 'clear');
    first.layout.rows = [{ fieldId: region.fieldId }, { fieldId: category.fieldId }];
    const expanded = buildPivotGridProjection(workbook, first);
    assert.equal(expanded.collision.status, 'collision');
    assert.equal(expanded.collision.reasons.includes('pivot'), true);
    assert.deepEqual(expanded.cells, initial.cells);
    assert.deepEqual(expanded.occupiedRange, initial.occupiedRange);
    first.target = { sheetId: sheet.id, anchor: { row: 5, column: 10 } };
    const moved = buildPivotGridProjection(workbook, first);
    assert.equal(moved.collision.status, 'clear');
    assert.equal(moved.occupiedRange.startColumn, 10);
  });

  it('fails closed for occupied structures, anchored objects, protection, and worksheet bounds', () => {
    const workbook = workbookWithData();
    const sheet = workbook.getSheet('sheet-1');
    const pivot = buildPivotModel(workbook, sheet.id, 'pivot-structure-collision', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    pivot.target = { sheetId: sheet.id, anchor: { row: 5, column: 0 } };
    const clear = buildPivotGridProjection(workbook, pivot);
    assert.equal(clear.collision.status, 'clear');
    sheet.sheetTables.push({ id: 'table-collision', sheetId: sheet.id, name: 'TableCollision', range: structuredClone(clear.occupiedRange), hasHeaderRow: true, hasTotalRow: false, showBandedRows: true, showBandedColumns: false, showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'none', columns: [] });
    sheet.drawings.push({ id: 'drawing-collision', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'two-cell', row: clear.occupiedRange.startRow, column: clear.occupiedRange.startColumn, endRow: clear.occupiedRange.startRow + 1, endColumn: clear.occupiedRange.startColumn + 1 }, transform: { x: 0, y: 0, width: 100, height: 40 }, zIndex: 1, payloadId: 'drawing-collision-payload' });
    sheet.drawingPayloads.set('drawing-collision-payload', { kind: 'shape', type: 'rectangle', fill: '#fff', stroke: '#000' });
    sheet.protectionRules.push({ id: 'protection-collision', scope: 'range', sheetId: sheet.id, range: structuredClone(clear.occupiedRange), locked: true, allow: {} });
    const blocked = buildPivotGridProjection(workbook, pivot);
    assert.equal(blocked.collision.status, 'collision');
    assert.equal(blocked.collision.reasons.includes('sheet-table'), true);
    assert.equal(blocked.collision.reasons.includes('drawing'), true);
    assert.equal(blocked.collision.reasons.includes('protection'), true);
    assert.equal(blocked.collision.conflicts.some((entry) => entry.participantId === 'table-collision'), true);
    pivot.target = { sheetId: sheet.id, anchor: { row: sheet.rowCount - 1, column: 0 } };
    const outOfBounds = detectPivotCollision(workbook, pivot, { sheetId: sheet.id, startRow: sheet.rowCount - 1, endRow: sheet.rowCount, startColumn: 0, endColumn: 1 });
    assert.equal(outOfBounds.reasons.includes('worksheet-bounds'), true);
  });

  it('rejects an existing Pivot with no verifiable current or last-valid footprint', () => {
    const workbook = workbookWithData();
    const sheet = workbook.getSheet('sheet-1');
    const candidate = buildPivotModel(workbook, sheet.id, 'pivot-known', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    const unresolved = buildPivotModel(workbook, sheet.id, 'pivot-unresolved', { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(candidate && unresolved);
    unresolved.source = { kind: 'data-source', dataSourceId: 'missing-source' };
    candidate.target = { sheetId: sheet.id, anchor: { row: 5, column: 10 } };
    sheet.pivots.push(unresolved);
    const collision = buildPivotGridProjection(workbook, candidate);
    assert.equal(collision.collision.status, 'collision');
    assert.equal(collision.collision.reasons.includes('unresolved-pivot'), true);
    assert.equal(collision.collision.conflicts.some((entry) => entry.participantId === unresolved.id), true);
  });

  it('rejects a stale layout result, localizes field captions, and preserves source-stale manual results', () => {
    const workbook = workbookWithData();
    const pivot = buildPivotModel(workbook, 'sheet-1', 'pivot-revision', { sheetId: 'sheet-1', startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 });
    assert.ok(pivot);
    pivot.target = { sheetId: 'sheet-1', anchor: { row: 8, column: 0 } };
    pivot.refreshPolicy = { ...pivot.refreshPolicy, mode: 'manual', refreshOnLoad: false };
    const catalog = getPivotFieldCatalog(workbook, pivot);
    const region = catalog.fields.find((field) => field.name === 'Region')!;
    const amount = catalog.fields.find((field) => field.name === 'Amount')!;
    pivot.layout.rows = [{ fieldId: region.fieldId }];
    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const firstResult = computePivotResult(workbook, pivot);
    const firstProjection = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.equal(firstProjection.cells.some((cell) => cell.text === 'Amount'), true);

    workbook.getSheet('sheet-1').cells.set(1, 1, { value: 100 });
    const stale = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.equal(stale.refresh.status, 'stale');
    assert.equal(stale.cells.some((cell) => cell.value === 35), true);

    pivot.layout.values = [{ valueId: `value:${amount.fieldId}`, fieldId: amount.fieldId, summarizeBy: 'count' }];
    const refreshedLayout = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.notEqual(refreshedLayout.refresh.status, 'stale');
    assert.equal(refreshedLayout.cells.some((cell) => cell.value === 3), true);
  });

  it('retains a cached block Pivot result across loading and source failure states', () => {
    const workbook = new WorkbookModel('pivot-block', 'Pivot Block');
    workbook.addDataSource({
      schema: 'DataSourceManifest',
      version: 1,
      id: 'source-block',
      name: 'Block Source',
      kind: 'chunked-table',
      sourceSheetId: 'sheet-1',
      rowCount: 2,
      fields: [
        { id: 'region', name: 'Region', ordinal: 0, type: 'text' },
        { id: 'amount', name: 'Amount', ordinal: 1, type: 'number' },
      ],
      blockRowCount: 65_536,
      blocks: [],
      revision: 1,
    });
    const pivot = {
      schema: 'PivotDefinition' as const,
      id: 'pivot-block',
      source: { kind: 'data-source' as const, dataSourceId: 'source-block' },
      target: { sheetId: 'sheet-1', anchor: { row: 5, column: 0 } },
      fieldCatalog: {
        schema: 'PivotFieldCatalog' as const,
        fields: [
          { fieldId: 'region', name: 'Region', dataType: 'text' as const, ordinal: 0 },
          { fieldId: 'amount', name: 'Amount', dataType: 'number' as const, ordinal: 1 },
        ],
      },
      refreshPolicy: { mode: 'on-change' as const, preserveFormatting: true, refreshOnLoad: true },
      layout: {
        rows: [{ fieldId: 'region' }], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const },
        values: [{ valueId: `value:${'amount'}`, fieldId: 'amount', summarizeBy: 'sum' as const }],
        subtotalLocation: 'bottom' as const, showRowGrandTotals: true, showColumnGrandTotals: true, reportLayout: 'compact' as const,
        expansion: { expandedNodeIds: [], collapsedNodeIds: [], showButtons: true },
      },
    };
    const loading = buildPivotGridProjection(workbook, pivot);
    assert.equal(loading.refresh.status, 'refreshing');
    assert.equal(loading.cells.some((cell) => cell.kind === 'loading'), true);
    const result = computePivotResultFromBlockSource(workbook, pivot, {
      fields: [
        { fieldId: 'region', name: 'Region', ordinal: 0, dataType: 'text' },
        { fieldId: 'amount', name: 'Amount', ordinal: 1, dataType: 'number' },
      ],
      rows: [
        { values: { region: 'East', amount: 10 }, paths: [{ sheetId: 'sheet-1', row: 1 }] },
        { values: { region: 'West', amount: 20 }, paths: [{ sheetId: 'sheet-1', row: 2 }] },
      ],
    }, 'source-block:1');
    const ready = buildPivotGridProjection(workbook, pivot, result, { sourceState: { availability: 'ready' } });
    assert.equal(ready.refresh.status, 'ready');
    assert.equal(ready.cells.some((cell) => cell.value === 30), true);
    const failed = buildPivotGridProjection(workbook, pivot, undefined, { sourceState: { availability: 'error', error: 'offline' } });
    assert.equal(failed.refresh.status, 'error');
    assert.equal(failed.cells.some((cell) => cell.value === 30), true);
  });
});
