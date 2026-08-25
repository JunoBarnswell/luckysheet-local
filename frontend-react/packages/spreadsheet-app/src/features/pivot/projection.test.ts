import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel, type PivotModel } from '@react-sheets/core-model';
import { FormulaEngine } from '@react-sheets/formula-engine';
import {
  aggregatePivotValues,
  buildPivotGroupedFilterMembers,
  buildPivotGridProjection,
  computePivotResult,
  computePivotResultFromBlockSource,
  getPivotFieldCatalog,
  getPivotSourceRanges,
  hitTestPivotProjection,
  summarizePivotReportFilters,
} from './engine';
import { buildPivotModel } from './helpers';
import { createSpillEnvironment } from '../../formula-spill-sync';

function workbookWithData(): WorkbookModel {
  const workbook = new WorkbookModel('pivot-projection', 'Pivot Projection');
  const sheet = workbook.getSheet('sheet-1');
  [['Region', 'Amount'], ['East', 10], ['West', 20], ['East', 5]].forEach((row, rowIndex) => row.forEach((value, columnIndex) => sheet.cells.set(rowIndex, columnIndex, { value })));
  return workbook;
}

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
    layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant' as const, numeric: false, caseFirst: 'false' as const }, values: [], subtotalLocation: 'bottom' as const, showGrandTotals: true, compact: true, repeatLabels: false },
  };
  const catalog = getPivotFieldCatalog(workbook, pivot);
  const region = catalog.fields.find((field) => field.name === 'Region')!;
  const amount = catalog.fields.find((field) => field.name === 'Amount')!;
  pivot.fieldCatalog = catalog;
  pivot.layout.rows = [{ fieldId: region.fieldId }];
  pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
  return pivot;
}

describe('native PivotGridProjection contract', () => {
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
      layout: { rows: [], columns: [], filters: [], allowMultipleFiltersPerField: true, collation: { locale: 'en-US', sensitivity: 'variant', numeric: false, caseFirst: 'false' }, values: [], subtotalLocation: 'bottom', showGrandTotals: true, compact: true, repeatLabels: false },
    };
    const catalog = getPivotFieldCatalog(workbook, pivot, formula);
    pivot.fieldCatalog = catalog;
    pivot.layout.rows = [{ fieldId: catalog.fields[0]!.fieldId }];
    pivot.layout.values = [{ fieldId: catalog.fields[0]!.fieldId, summarizeBy: 'count' }];
    assert.equal(getPivotSourceRanges(workbook, pivot, formula)[0]?.endRow, 2);
    assert.deepEqual(computePivotResult(workbook, pivot, formula).rows.map((node) => node.key), [1, 2]);
    formula.setFormula({ sheetId: sheet.id, row: 0, column: 0 }, '=SEQUENCE(4,1,0,1)');
    assert.equal(getPivotSourceRanges(workbook, pivot, formula)[0]?.endRow, 3);
    assert.deepEqual(computePivotResult(workbook, pivot, formula).rows.map((node) => node.key), [1, 2, 3]);
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: getPivotFieldCatalog(workbook, pivot).fields.find((field) => field.name === 'Amount')!.fieldId, summarizeBy: 'count' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
    pivot.layout.allowMultipleFiltersPerField = true;
    pivot.layout.filters = [
      { kind: 'manual', family: 'manual', fieldId: region.fieldId, mode: 'all', memberKeys: [] },
      { kind: 'condition', family: 'label', fieldId: region.fieldId, operator: 'begins-with', value: 'E' },
      { kind: 'top-items', family: 'top-items', fieldId: region.fieldId, count: 1, valueFieldId: amount.fieldId, direction: 'top' },
    ];
    const summary = summarizePivotReportFilters(pivot.layout.filters, catalog, region.fieldId);
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: sales.fieldId, summarizeBy: 'sum' }];
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2024', '2025', '2026']);

    pivot.layout.rows[0] = { fieldId: year.fieldId, sort: { direction: 'ascending', by: 'value', valueFieldId: sales.fieldId } };
    assert.deepEqual(computePivotResult(workbook, pivot).rows.map((node) => node.label), ['2025', '2024', '2026']);
    pivot.layout.rows[0] = { fieldId: year.fieldId, sort: { direction: 'ascending', by: 'value' } };
    assert.throws(() => computePivotResult(workbook, pivot), /requires a valueFieldId/);
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: sales.fieldId, summarizeBy: 'sum' }];
    pivot.layout.showGrandTotals = true;
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
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
    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'sum' }];
    const firstResult = computePivotResult(workbook, pivot);
    const firstProjection = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.equal(firstProjection.cells.some((cell) => cell.text === 'Amount'), true);

    workbook.getSheet('sheet-1').cells.set(1, 1, { value: 100 });
    const stale = buildPivotGridProjection(workbook, pivot, firstResult);
    assert.equal(stale.refresh.status, 'stale');
    assert.equal(stale.cells.some((cell) => cell.value === 35), true);

    pivot.layout.values = [{ fieldId: amount.fieldId, summarizeBy: 'count' }];
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
        values: [{ fieldId: 'amount', summarizeBy: 'sum' as const }],
        subtotalLocation: 'bottom' as const, showGrandTotals: true, compact: true, repeatLabels: false,
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
