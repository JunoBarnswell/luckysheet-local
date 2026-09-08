import assert from 'node:assert/strict';
import test from 'node:test';
import { CommandRuntime, type CommandContext, type MutationInfo } from '@react-sheets/command-runtime';
import { WorkbookModel, type SheetSnapshot } from '@react-sheets/core-model';
import { registerInsertCommands } from './commands';

function sheet(overrides: Partial<SheetSnapshot> = {}): SheetSnapshot {
  return {
    kind: 'table-sheet',
    id: 'table-sheet-1',
    name: 'Table Sheet',
    rowCount: 20,
    columnCount: 8,
    cells: {
      0: {
        0: { value: 'Name' },
        1: { value: null, formula: '=1+1', formulaMetadata: { kind: 'shared', sharedIndex: 4 } },
        3: { value: 'Amount' },
      },
    },
    merges: [],
    pane: { kind: 'none' },
    pivots: [],
    sparklines: [],
    drawings: [],
    drawingPayloads: {},
    review: { threads: [], notes: [] },
    defaultRowHeightPx: 20,
    defaultColumnWidthPx: 64,
    tableSheet: { viewId: 'view-1', columns: [], grouping: [] },
    ...overrides,
  } as SheetSnapshot;
}

function fixture() {
  const workbook = new WorkbookModel('insert-test', 'Insert test');
  const runtime = new CommandRuntime(workbook);
  registerInsertCommands(runtime);
  const mutations: MutationInfo[] = [];
  const context = {
    workbook,
    operationId: 'advanced-sheet-operation',
    executeCommand: () => ({ operationId: 'advanced-sheet-operation', mutationCount: 0, affectedRanges: [] }),
    applyMutation: (mutation: MutationInfo) => mutations.push(structuredClone(mutation)),
    recordOperation: () => ({ operationId: 'advanced-sheet-operation' }),
  } as CommandContext;
  return { workbook, runtime, mutations, context };
}

test('advanced sheet creation plans one canonical transaction with definition and contiguous cell runs', () => {
  const { workbook, runtime, mutations, context } = fixture();
  const result = runtime.registry
    .getCommand('sheet.create.advanced')
    .execute({ sheet: sheet(), index: 0 }, context);

  assert.equal(result.mutationCount, 5);
  assert.deepEqual(
    mutations.map(mutation => mutation.id),
    ['sheet.add', 'tableSheet.update', 'range.set', 'range.set', 'sheet.reordered'],
  );
  assert.deepEqual(mutations[2]!.params, {
    sheetId: 'table-sheet-1',
    startRow: 0,
    startColumn: 0,
    values: [[
      { value: 'Name' },
      { value: null, formula: '=1+1' },
    ]],
  });
  assert.deepEqual(mutations[3]!.params, {
    sheetId: 'table-sheet-1',
    startRow: 0,
    startColumn: 3,
    values: [[{ value: 'Amount' }]],
  });
  assert.equal(workbook.sheets.has('table-sheet-1'), false);
  assert.throws(() => workbook.manifest(), { code: 'WORKBOOK_NOT_OPEN' });
});

test('invalid advanced sheet creation rejects before any workbook revision is published', () => {
  const { workbook, runtime, context } = fixture();
  const invalid = sheet({ cells: { 20: { 0: { value: 'outside' } } } });
  assert.throws(
    () => runtime.registry.getCommand('sheet.create.advanced').execute({ sheet: invalid }, context),
    /cell row is invalid/,
  );
  assert.equal(workbook.sheets.has('table-sheet-1'), false);
  assert.deepEqual(workbook.sheetOrder, ['sheet-1']);
  assert.throws(() => workbook.manifest(), { code: 'WORKBOOK_NOT_OPEN' });
});
