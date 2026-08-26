import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel } from '@react-sheets/core-model';
import { createFormulaError } from '@react-sheets/formula-engine';
import {
  computeConditionalOverlays,
  computeFilterHiddenRows,
  createEffectiveFilterVisualResolver,
  getAutoFilterValueDomain,
  getAutoFilterDateDomain,
  getAutoFilterDomainDescriptor,
  validateFilterCriterionAgainstDomain,
  registerSheetCommands,
  normalizeAutoFilterModel,
  normalizeDataValidationRule,
  validationList,
  validateDataInput,
} from './index';
import { compareSortValues, resolveSortCellValue } from './data-features';

function runtime(): { workbook: WorkbookModel; commands: CommandRuntime } {
  const workbook = new WorkbookModel('m3-m4', 'M3/M4');
  const commands = new CommandRuntime(workbook);
  registerSheetCommands(commands);
  return { workbook, commands };
}

test('sorting uses resolved formula results, keeps stable ties, and replays/undoes as one permutation', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  commands.execute('sheet.range.set', {
    sheetId: sheet.id,
    startRow: 0,
    startColumn: 0,
    values: [
      [{ value: 'Calculated' }, { value: 'Stable row' }],
      [{ formula: '=B2+10', value: null }, { value: 'first' }],
      [{ formula: '=B3+10', value: null }, { value: 'second' }],
      [{ formula: '=B4+10', value: null }, { value: 'third' }],
    ],
  });
  const beforeSort = workbook.snapshot();
  const formulaResults = new Map([[1, 20], [2, 5], [3, 5]]);
  commands.setCellValueResolver((_currentSheet, row, column) => column === 0 ? formulaResults.get(row) ?? null : undefined);

  commands.execute('data.sort.rows', {
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    criteria: [{ column: 0, ascending: true }],
    hasHeader: true,
  });

  assert.equal(commands.getHistoryDepth().undo, 2, 'the setup write and sort are separate commands');
  const sortEntry = commands.getUndoEntries().at(-1)!;
  assert.equal(sortEntry.redo.length, 1);
  assert.deepEqual((sortEntry.redo[0]?.params as { sourceRows: number[] }).sourceRows, [2, 3, 1]);
  assert.equal(sheet.cells.get(1, 1)?.value, 'second');
  assert.equal(sheet.cells.get(2, 1)?.value, 'third');
  assert.equal(sheet.cells.get(3, 1)?.value, 'first');

  const remoteWorkbook = WorkbookModel.fromSnapshot(beforeSort);
  const remoteCommands = new CommandRuntime(remoteWorkbook);
  registerSheetCommands(remoteCommands);
  remoteCommands.applyRemoteMutations(sortEntry.redo);
  const currentSheetSnapshot = workbook.snapshot().sheets.find((candidate) => candidate.id === sheet.id);
  const remoteSheetSnapshot = remoteWorkbook.snapshot().sheets.find((candidate) => candidate.id === sheet.id);
  assert.deepEqual(remoteSheetSnapshot?.cells, currentSheetSnapshot?.cells);

  assert.equal(commands.undo(), true);
  assert.equal(sheet.cells.get(1, 1)?.value, 'first');
  assert.equal(sheet.cells.get(2, 1)?.value, 'second');
  assert.equal(sheet.cells.get(3, 1)?.value, 'third');
  assert.equal(commands.redo(), true);
  assert.equal(sheet.cells.get(1, 1)?.value, 'second');
  assert.equal(sheet.cells.get(3, 1)?.value, 'first');
});

test('sort keys retain canonical typed formula results and reject unresolved values', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(1, 0, { formula: '=1+1', value: null, formulaValue: 2 });
  assert.equal(resolveSortCellValue(sheet, 1, 0), 2);
  sheet.cells.set(2, 0, { formula: '=A1', value: null });
  assert.throws(() => resolveSortCellValue(sheet, 2, 0), /formula result unavailable/);
  assert.equal(compareSortValues(2, 10) < 0, true);
  assert.equal(compareSortValues(true, 'true') > 0, true);
  assert.equal(compareSortValues(createFormulaError('#N/A', 'missing'), null) < 0, true);
  assert.equal(compareSortValues(createFormulaError('#DIV/0!', 'zero'), createFormulaError('#N/A', 'missing')) < 0, true);
  assert.equal(compareSortValues(5, 5), 0);
  assert.throws(
    () => resolveSortCellValue(sheet, 1, 0, () => [[2, 3]]),
    /unresolved array result/,
  );
  assert.throws(
    () => resolveSortCellValue(sheet, 1, 0, () => Number.POSITIVE_INFINITY),
    /non-finite numeric result/,
  );
});

test('rows.permuted rejects a tampered duplicate source order before changing cells', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(1, 0, { value: 'A' });
  sheet.cells.set(2, 0, { value: 'B' });
  const before = workbook.snapshot();
  assert.throws(() => commands.applyRemoteMutations([{
    id: 'rows.permuted',
    unitId: workbook.unitId,
    sheetId: sheet.id,
    params: {
      sheetId: sheet.id,
      range: { sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 },
      sourceRows: [1, 1],
    },
    affectedRanges: [{ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 16_383 }],
  }]), /Invalid mutation history/);
  assert.deepEqual(workbook.snapshot().sheets.find((candidate) => candidate.id === sheet.id)?.cells,
    before.sheets.find((candidate) => candidate.id === sheet.id)?.cells);
});

test('Home, worksheet AutoFilter, and table sorting share the resolved-value owner', () => {
  const seed = (): { workbook: WorkbookModel; commands: CommandRuntime } => {
    const next = runtime();
    const sheet = next.workbook.getSheet(next.workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Calculated' });
    sheet.cells.set(0, 1, { value: 'Row' });
    sheet.cells.set(1, 0, { formula: '=B2', value: null });
    sheet.cells.set(1, 1, { value: 'twenty' });
    sheet.cells.set(2, 0, { formula: '=B3', value: null });
    sheet.cells.set(2, 1, { value: 'five' });
    sheet.cells.set(3, 0, { formula: '=B4', value: null });
    sheet.cells.set(3, 1, { value: 'ten' });
    const formulaResults = new Map([[1, 20], [2, 5], [3, 10]]);
    next.commands.setCellValueResolver((_sheet, row, column) => column === 0 ? formulaResults.get(row) ?? null : undefined);
    return next;
  };
  const assertSorted = (commands: CommandRuntime, workbook: WorkbookModel): void => {
    const sheet = workbook.getSheet(workbook.primarySheetId);
    assert.deepEqual([1, 2, 3].map((row) => sheet.cells.get(row, 1)?.value), ['five', 'ten', 'twenty']);
    assert.equal(commands.getHistoryDepth().undo > 0, true);
  };

  for (const command of ['data.sort.quick', 'sheet.sort.multi'] as const) {
    const { workbook, commands } = seed();
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const range = { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    if (command === 'data.sort.quick') commands.execute(command, { sheetId: sheet.id, range, sortColumn: 0, ascending: true, hasHeader: true });
    else commands.execute(command, { sheetId: sheet.id, range, criteria: [{ column: 0, ascending: true }], hasHeader: true });
    assertSorted(commands, workbook);
  }

  for (const owner of ['worksheet', 'table'] as const) {
    const { workbook, commands } = seed();
    const sheet = workbook.getSheet(workbook.primarySheetId);
    const range = { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 };
    const filter = normalizeAutoFilterModel({ sheetId: sheet.id, range, columns: {} });
    if (owner === 'worksheet') {
      sheet.autoFilter = filter;
    } else {
      commands.execute('sheetTable.add', {
        id: 'table-sort', sheetId: sheet.id, name: 'SortTable', range, hasHeaderRow: true, hasTotalRow: false,
        showBandedRows: false, showBandedColumns: false, showFirstColumn: false, showLastColumn: false,
        showFilterButton: true, autoExpand: 'both', columns: [{ id: 'calculated', name: 'Calculated' }, { id: 'row', name: 'Row' }],
      });
      commands.execute('sheetTable.autoFilter.set', { sheetId: sheet.id, tableId: 'table-sort', autoFilter: filter });
    }
    commands.execute('sheet.autoFilter.sort', { sheetId: sheet.id, column: 0, ascending: true });
    assertSorted(commands, workbook);
  }
});

test('scoped defined names survive command undo and snapshot round-trip', () => {
  const { workbook, commands } = runtime();
  const local = workbook.addSheet('sheet-2', 'Local');
  commands.execute('workbook.name.set', { name: 'Rate', value: '0.1' });
  commands.execute('workbook.name.set', { name: 'Rate', formula: '0.2', scope: 'sheet', sheetId: local.id });
  assert.equal(workbook.getDefinedName('Rate', local.id)?.formula, '0.2');
  commands.undo();
  assert.equal(workbook.getDefinedNameExact('Rate', 'sheet', local.id), undefined);
  const restored = WorkbookModel.fromSnapshot(workbook.snapshot());
  assert.equal(restored.getDefinedName('Rate')?.formula, '0.1');
});

test('total row inserts a real row and does not overwrite data below table', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Name' });
  sheet.cells.set(1, 0, { value: 'A' });
  sheet.cells.set(2, 0, { value: 'B' });
  sheet.cells.set(3, 0, { value: 'outside' });
  commands.execute('sheetTable.add', {
    id: 'table-1', sheetId: sheet.id, name: 'Sales',
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: false,
    showBandedColumns: false, showFirstColumn: false, showLastColumn: false, showFilterButton: false, autoExpand: 'both',
    columns: [{ id: 'name', name: 'Name', totalsFunction: 'count' }],
  });
  commands.execute('sheetTable.toggleTotalRow', { sheetId: sheet.id, tableId: 'table-1', enabled: true });
  assert.equal(sheet.cells.get(4, 0)?.value, 'outside');
  assert.match(sheet.cells.get(3, 0)?.formula ?? '', /^=SUBTOTAL\(/);
  commands.undo();
  assert.equal(sheet.cells.get(3, 0)?.value, 'outside');
});

test('sort and remove duplicates preserve formulas and use structural row deletion', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  commands.execute('sheet.range.set', {
    sheetId: sheet.id, startRow: 0, startColumn: 0,
    values: [
      [{ value: 'Key' }, { value: 'Value' }, { value: 'Formula' }],
      [{ value: 'B' }, { value: 2 }, { formula: '=B2*2', value: null }],
      [{ value: 'A' }, { value: 1 }, { formula: '=B3*2', value: null }],
      [{ value: 'A' }, { value: 1 }, { formula: '=B4*2', value: null }],
    ],
  });
  commands.execute('sheet.sort', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }, sortColumn: 0, ascending: true, hasHeader: true });
  const permutation = commands.getUndoEntries().at(-1)?.redo[0];
  assert.equal(permutation?.id, 'rows.permuted');
  assert.deepEqual(permutation?.affectedRanges, [{
    sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: sheet.columnCount - 1,
  }]);
  assert.equal(sheet.cells.get(1, 0)?.value, 'A');
  assert.equal(sheet.cells.get(1, 2)?.formula, '=B3*2');
  commands.execute('data.removeDuplicates', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 }, columns: [0], hasHeader: true });
  assert.equal(sheet.cells.get(3, 0), undefined);
  assert.equal(sheet.cells.get(2, 2)?.formula, '=B2*2');
});

test('conditional format priority/stop and validation alert style are represented', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  commands.execute('sheet.range.set', { sheetId: sheet.id, startRow: 0, startColumn: 0, values: [[{ value: 10 }], [{ value: 1 }]] });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: { id: 'top', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }], type: 'topBottom', topBottom: { direction: 'top', rank: 1 }, priority: 1, stopIfTrue: true, style: { bold: true } },
  });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: { id: 'lower', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }], type: 'highlight', operator: 'lessThan', value1: 20, priority: 2, style: { italic: true } },
  });
  assert.equal(computeConditionalOverlays(sheet).get('0:0')?.style?.bold, true);
  commands.execute('sheet.dv.add', {
    sheetId: sheet.id,
    rule: { id: 'list', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }], type: 'list', listSource: { kind: 'values', values: ['A', 'B'] }, multiSelect: true, alertStyle: 'warning' },
  });
  const result = validateDataInput(sheet, 0, 1, 'A,B');
  assert.equal(result.valid, true);
  assert.equal(result.blocking, false);
});

test('transpose fails closed when a selected range contains a drawing', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.drawings.push({ id: 'd1', sheetId: sheet.id, kind: 'shape', anchor: { kind: 'one-cell', row: 0, column: 0 }, transform: { x: 0, y: 0, width: 10, height: 10 }, zIndex: 0, payloadId: 'p1' });
  assert.throws(() => commands.execute('matrix.transpose', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 } }), /drawing anchors/);
});

test('Filter supports compound text/blank/date conditions and rejects out-of-range criteria', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Value' });
  sheet.cells.set(1, 0, { value: 'Alpha' });
  sheet.cells.set(2, 0, { value: '' });
  sheet.cells.set(3, 0, { value: 'Beta' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'custom', join: 'and', conditions: [{ operator: 'contains', value: 'a' }] } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((a, b) => a - b), [2]);
  assert.throws(() => normalizeAutoFilterModel({
    ...sheet.autoFilter!,
    columns: { 1: { column: 1, showButton: true, hiddenButton: false, criterion: { kind: 'custom', join: 'and', conditions: [{ operator: 'equals', value: 'x' }] } } },
  }), /outside/);
});

test('AutoFilter value domain is complete and ignores the current column criterion', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  for (let row = 0; row <= 250; row += 1) {
    sheet.cells.set(row, 0, { value: row === 0 ? 'Name' : `Value-${row}` });
    sheet.cells.set(row, 1, { value: row === 0 ? 'Group' : row % 2 === 0 ? 'keep' : 'drop' });
  }
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 250, startColumn: 0, endColumn: 1 },
    columns: {
      0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['Value-2'], includeBlank: false } },
      1: { column: 1, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: ['keep'], includeBlank: false } },
    },
  });
  const domain = getAutoFilterValueDomain(sheet, 0);
  assert.equal(domain.length, 125);
  assert.equal(domain.includes('Value-2'), true);
  assert.equal(domain.includes('Value-3'), false);
});

test('AutoFilter date domain remains typed instead of parsing display strings', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Date' });
  sheet.cells.set(1, 0, { value: 'not-a-date' });
  sheet.cells.set(2, 0, { value: '2026-08-15T13:14:15Z' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false } },
  });
  const domain = getAutoFilterDateDomain(sheet, 0);
  assert.equal(domain.length, 0, 'text that resembles an ISO date remains text');
  sheet.cells.set(1, 0, { value: 46249.5515625, numberFormat: 'yyyy-mm-dd hh:mm:ss' });
  const numericDomain = getAutoFilterDateDomain(sheet, 0);
  assert.equal(numericDomain.length, 1);
  assert.equal(numericDomain[0]?.value, 46249.5515625);
  assert.deepEqual(numericDomain[0]?.group, { year: 2026, month: 8, day: 15, hour: 13, minute: 14, second: 15 });
});

test('FilterDomainDescriptor uses resolved formula values and exposes only compatible families', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Amount' });
  sheet.cells.set(1, 0, { value: '=A1+1', formula: '=A1+1', formulaValue: 42 });
  sheet.cells.set(2, 0, { value: 7 });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false } },
  });
  const descriptor = getAutoFilterDomainDescriptor(sheet, 0);
  assert.deepEqual(descriptor.values, [7, 42]);
  assert.equal(descriptor.dominantType, 'number');
  assert.deepEqual(descriptor.supportedFamilies, ['values', 'number']);
  assert.equal(descriptor.values.map(String).includes('=A1+1'), false);
});

test('AutoFilter does not expose an authored formula string when no result is available', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Amount' });
  sheet.cells.set(1, 0, { value: '=A1+1', formula: '=A1+1' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false } },
  });
  assert.deepEqual(getAutoFilterValueDomain(sheet, 0), ['']);
  assert.deepEqual([...computeFilterHiddenRows(sheet)], []);
});

test('worksheet and table filters share the typed formula scalar for value, custom, date and Top10 criteria', () => {
  const create = (owner: 'worksheet' | 'table'): { workbook: WorkbookModel; sheet: ReturnType<WorkbookModel['getSheet']> } => {
    const { workbook } = runtime();
    const sheet = workbook.getSheet(workbook.primarySheetId);
    sheet.cells.set(0, 0, { value: 'Status' });
    sheet.cells.set(0, 1, { value: 'Amount' });
    sheet.cells.set(0, 2, { value: 'Date' });
    sheet.cells.set(1, 0, { value: '=\"Open\"', formula: '=\"Open\"', formulaValue: 'Open' });
    sheet.cells.set(2, 0, { value: '=\"Closed\"', formula: '=\"Closed\"', formulaValue: 'Closed' });
    sheet.cells.set(3, 0, { value: '=\"Open\"', formula: '=\"Open\"', formulaValue: 'Open' });
    sheet.cells.set(1, 1, { value: '=20', formula: '=20', formulaValue: 20 });
    sheet.cells.set(2, 1, { value: '=5', formula: '=5', formulaValue: 5 });
    sheet.cells.set(3, 1, { value: '=10', formula: '=10', formulaValue: 10 });
    sheet.cells.set(1, 2, { value: '=46260', formula: '=46260', formulaValue: 46260, numberFormat: 'yyyy-mm-dd' });
    sheet.cells.set(2, 2, { value: '=46259', formula: '=46259', formulaValue: 46259, numberFormat: 'yyyy-mm-dd' });
    sheet.cells.set(3, 2, { value: '=46260', formula: '=46260', formulaValue: 46260, numberFormat: 'yyyy-mm-dd' });
    const range = { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 };
    const filter = normalizeAutoFilterModel({
      sheetId: sheet.id,
      range,
      columns: {
        0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'custom', join: 'and', conditions: [{ operator: 'equals', value: 'Open' }] } },
        1: { column: 1, showButton: true, hiddenButton: false, criterion: { kind: 'top10', top: true, percent: false, rank: 1 } },
        2: { column: 2, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: [], includeBlank: false, dateGroups: [{ year: 2026, month: 8, day: 26 }] } },
      },
    });
    if (owner === 'worksheet') sheet.autoFilter = filter;
    else sheet.sheetTables.push({
      id: 'formula-filter-table', sheetId: sheet.id, name: 'FormulaFilterTable', range,
      hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
      showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'both',
      columns: [{ id: 'status', name: 'Status' }, { id: 'amount', name: 'Amount' }, { id: 'date', name: 'Date' }],
      autoFilter: filter,
    });
    return { workbook, sheet };
  };

  for (const owner of ['worksheet', 'table'] as const) {
    const { sheet } = create(owner);
    assert.deepEqual(getAutoFilterValueDomain(sheet, 0), ['Open']);
    assert.deepEqual(getAutoFilterDateDomain(sheet, 2).map((entry) => entry.group), [{ year: 2026, month: 8, day: 26, hour: 0, minute: 0, second: 0 }]);
    assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((left, right) => left - right), [2, 3]);
  }
});

test('FilterDomainDescriptor date hierarchy requires numeric values with canonical date format', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Date' });
  sheet.cells.set(1, 0, { value: '2026-08-15' });
  sheet.cells.set(2, 0, { value: 46249.5, numberFormat: 'yyyy-mm-dd hh:mm:ss' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false } },
  });
  const descriptor = getAutoFilterDomainDescriptor(sheet, 0);
  assert.equal(descriptor.dominantType, 'mixed');
  assert.equal(descriptor.dateDomain.length, 1);
  assert.equal(descriptor.dateHierarchy.includes('year'), true);
  assert.equal(descriptor.supportedFamilies.includes('date'), false, 'mixed text/date values do not expose date operators');
  assert.throws(() => validateFilterCriterionAgainstDomain(descriptor, { kind: 'dynamic', type: 'nextQuarter' }), /FILTER_DOMAIN_MISMATCH/);
  assert.throws(() => validateFilterCriterionAgainstDomain(descriptor, { kind: 'custom', join: 'and', conditions: [{ operator: 'contains', value: '2026' }] }), /FILTER_OPERATOR_MISMATCH|FILTER_DOMAIN_MISMATCH/);
});

test('FilterDomainDescriptor rejects visual criteria without a resolved visual domain', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Value' });
  sheet.cells.set(1, 0, { value: 1 });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false } },
  });
  const descriptor = getAutoFilterDomainDescriptor(sheet, 0);
  assert.deepEqual(descriptor.colorDomain, []);
  assert.deepEqual(descriptor.iconDomain, []);
  assert.throws(() => validateFilterCriterionAgainstDomain(descriptor, { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#fff' } }), /FILTER_DOMAIN_MISMATCH/);
  assert.throws(() => validateFilterCriterionAgainstDomain(descriptor, { kind: 'icon', iconSet: '3TrafficLights1', iconId: 1 }), /FILTER_DOMAIN_MISMATCH/);
});

test('AutoFilter rejects malformed date-group shapes instead of guessing', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  assert.throws(() => normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'values', values: [], includeBlank: false, dateGroups: [{ year: 2026, month: 8, second: 1 }] } } },
  }), /requires minute/);
});

test('AutoFilter evaluates Top10 and dynamic date criteria against canonical rows', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Amount' });
  [10, 5, 20, 1].forEach((value, index) => sheet.cells.set(index + 1, 0, { value }));
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'top10', top: true, percent: false, rank: 2 } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((a, b) => a - b), [2, 4]);

  const referenceDate = { year: 2026, month: 8, day: 26, hour: 12, minute: 0, second: 0, millisecond: 0 };
  const serialToday = 46260;
  sheet.cells.set(0, 0, { value: 'Date' });
  sheet.cells.set(1, 0, { value: serialToday, numberFormat: 'yyyy-mm-dd' });
  sheet.cells.set(2, 0, { value: 36526, numberFormat: 'yyyy-mm-dd' });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'dynamic', type: 'today' } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet, undefined, '1900', undefined, { referenceDate })], [2]);

  assert.throws(() => normalizeAutoFilterModel({
    ...sheet.autoFilter!,
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'dynamic', type: 'attackerUnknown' as 'today' } } },
  }), /UNSUPPORTED_FEATURE: dynamic AutoFilter type "attackerUnknown" is not supported/);
});

test('AutoFilter color and icon criteria use native cell metadata or imported differential style', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Status' });
  sheet.cells.set(1, 0, { value: 'red', style: { background: '#ff0000' } });
  sheet.cells.set(2, 0, { value: 'blue', style: { background: '#0000ff' } });
  sheet.cells.set(3, 0, { value: 'icon', filterMetadata: { icon: { iconSet: '3TrafficLights1', iconId: 2 } } });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#ff0000' } } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [2, 3]);
  sheet.autoFilter.columns[0]!.criterion = { kind: 'icon', iconSet: '3TrafficLights1', iconId: 2 };
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [1, 2]);
});

test('AutoFilter color criteria consume the effective conditional-format fill/font without materializing CellData styles', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Fill' });
  sheet.cells.set(0, 1, { value: 'Font' });
  sheet.cells.set(1, 0, { value: 10, style: { background: '#0000ff' } });
  sheet.cells.set(1, 1, { value: 10 });
  sheet.cells.set(2, 0, { value: 1 });
  sheet.cells.set(2, 1, { value: 1 });
  sheet.cells.set(3, 0, { value: 20 });
  sheet.cells.set(3, 1, { value: 20 });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: {
      id: 'effective-filter-color',
      sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 3, startColumn: 0, endColumn: 1 }],
      type: 'highlight',
      operator: 'greaterThan',
      value1: 5,
      priority: 1,
      style: { background: '#ff0000', textColor: '#00ff00' },
    },
  });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
    columns: {
      0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#ff0000' } } },
      1: { column: 1, showButton: true, hiddenButton: false, criterion: { kind: 'color', target: 'font', dxfId: -1, style: { textColor: '#00FF00' } } },
    },
  });

  assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((a, b) => a - b), [2]);
  assert.equal(sheet.cells.get(1, 0)?.style?.background, '#0000ff', 'the conditional result must remain a projection');
  assert.equal(sheet.cells.get(2, 0)?.style, undefined, 'CF-only color must not be materialized into CellData');
  const visual = createEffectiveFilterVisualResolver(computeConditionalOverlays(sheet))(1, 0, sheet.cells.get(1, 0));
  assert.equal(visual.style.background, '#ff0000', 'CF fill overrides direct fill in the effective visual');
  assert.equal(visual.style.textColor, '#00ff00');
});

test('AutoFilter color matching follows conditional-format priority, stopIfTrue, color scales, and value changes', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Status' });
  sheet.cells.set(1, 0, { value: 10 });
  sheet.cells.set(2, 0, { value: 1 });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: {
      id: 'winner', sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }],
      type: 'highlight', operator: 'greaterThan', value1: 5, priority: 1, stopIfTrue: true,
      style: { background: '#00ff00' },
    },
  });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: {
      id: 'loser', sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }],
      type: 'highlight', operator: 'greaterThan', value1: 0, priority: 2,
      style: { background: '#ff0000' },
    },
  });
  sheet.autoFilter = normalizeAutoFilterModel({
    sheetId: sheet.id,
    range: { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 },
    columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#00ff00' } } } },
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [2]);
  sheet.cells.set(1, 0, { value: 0 });
  assert.deepEqual([...computeFilterHiddenRows(sheet)].sort((a, b) => a - b), [1, 2]);

  sheet.cells.set(1, 0, { value: 0 });
  sheet.cells.set(2, 0, { value: 10 });
  sheet.conditionalFormats.length = 0;
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: {
      id: 'scale', sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }],
      type: 'colorScale', priority: 1, minColor: '#000000', maxColor: '#ffffff',
    },
  });
  sheet.autoFilter.columns[0]!.criterion = { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#ffffff' } };
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [1]);
});

test('Worksheet and Sheet Table AutoFilter color criteria share the effective visual resolver', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Color' });
  sheet.cells.set(1, 0, { value: 10 });
  sheet.cells.set(2, 0, { value: 1 });
  commands.execute('sheet.cf.add', {
    sheetId: sheet.id,
    rule: {
      id: 'table-color', sheetId: sheet.id,
      ranges: [{ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 0 }],
      type: 'highlight', operator: 'greaterThan', value1: 5, priority: 1, style: { background: '#abcdef' },
    },
  });
  const range = { sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 };
  sheet.sheetTables.push({
    id: 'table-color', sheetId: sheet.id, name: 'ColorTable', range,
    hasHeaderRow: true, hasTotalRow: false, showBandedRows: false, showBandedColumns: false,
    showFirstColumn: false, showLastColumn: false, showFilterButton: true, autoExpand: 'none',
    columns: [{ id: 'color', name: 'Color' }],
    autoFilter: normalizeAutoFilterModel({
      sheetId: sheet.id, range,
      columns: { 0: { column: 0, showButton: true, hiddenButton: false, criterion: { kind: 'color', target: 'cell', dxfId: -1, style: { background: '#ABCDEF' } } } },
    }),
  });
  assert.deepEqual([...computeFilterHiddenRows(sheet)], [2]);
  assert.equal(sheet.cells.get(1, 0)?.style, undefined);
});

test('Validation supports custom AST, formula-backed list, time/date, multi-select and non-blocking alerts', () => {
  const { workbook } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'Allowed' });
  sheet.cells.set(1, 0, { value: 'Other' });
  const custom = normalizeDataValidationRule({ id: 'custom', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }], type: 'custom', formula1: '=B1="Allowed"', alertStyle: 'stop' });
  sheet.dataValidations.push(custom);
  assert.equal(validateDataInput(sheet, 0, 1, 'x').blocking, true);
  const list = normalizeDataValidationRule({ id: 'list', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 }], type: 'list', listSource: { kind: 'formula', formula: '=A1:A2' }, multiSelect: true, alertStyle: 'warning' });
  assert.deepEqual(validationList(list, sheet), ['Allowed', 'Other']);
  sheet.dataValidations.push(list);
  assert.equal(validateDataInput(sheet, 0, 2, 'Allowed,Other').blocking, false);
  const time = normalizeDataValidationRule({ id: 'time', sheetId: sheet.id, ranges: [{ sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 3, endColumn: 3 }], type: 'time', alertStyle: 'stop' });
  sheet.dataValidations.push(time);
  assert.equal(validateDataInput(sheet, 0, 3, '12:30').valid, true);
  assert.equal(validateDataInput(sheet, 0, 3, '25:30').blocking, true);
});

test('Text Columns, Split and Flip are one undoable transaction and clear stale output', () => {
  const { workbook, commands } = runtime();
  const sheet = workbook.getSheet(workbook.primarySheetId);
  sheet.cells.set(0, 0, { value: 'a,b' });
  sheet.cells.set(0, 2, { value: 'stale' });
  let commandEvents = 0;
  commands.onCommand(() => { commandEvents += 1; });
  const textResult = commands.execute('data.textToColumns', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }, delimiter: ',', maxColumns: 3 });
  assert.equal(commandEvents, 1);
  assert.equal(textResult.mutationCount, 2);
  assert.equal(sheet.cells.get(0, 2)?.value, null);
  commands.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'a,b');
  commands.execute('data.splitColumn', { sheetId: sheet.id, row: 0, column: 0, delimiter: ',', maxColumns: 2 });
  assert.equal(sheet.cells.get(0, 1)?.value, 'b');
  commands.undo();
  assert.equal(sheet.cells.get(0, 0)?.value, 'a,b');
  sheet.cells.set(0, 0, { value: 'a' });
  sheet.cells.set(0, 1, { value: 'b' });
  commands.execute('matrix.flip', { sheetId: sheet.id, range: { sheetId: sheet.id, startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 }, direction: 'horizontal' });
  assert.equal(sheet.cells.get(0, 0)?.value, 'b');
});
