import test from 'node:test';
import assert from 'node:assert/strict';
import { WorkbookModel } from '@react-sheets/core-model';
import { CommandRegistry, CommandRuntime, type MutationInfo } from './index';

const cellRange = (params: { row: number; column: number; sheetId?: string }) => [{
  sheetId: params.sheetId ?? 'sheet-1',
  startRow: params.row,
  endRow: params.row,
  startColumn: params.column,
  endColumn: params.column,
}];

const cellSetMetadata = {
  schema: {
    name: 'CellSet',
    validate: (value: unknown) => {
      if (!value || typeof value !== 'object') return false;
      const params = value as Record<string, unknown>;
      return Number.isInteger(params.row) && Number.isInteger(params.column) && 'value' in params;
    },
  },
  permission: { capability: 'test.cell.write' },
  affectedRanges: { resolve: cellRange },
  inversePolicy: { allowedMutationIds: ['cell.restore'], minCount: 1 },
} as const;

const cellRestoreMetadata = {
  schema: {
    name: 'CellRestore',
    validate: (value: unknown) => {
      if (!value || typeof value !== 'object') return false;
      const params = value as Record<string, unknown>;
      return Number.isInteger(params.row) && Number.isInteger(params.column);
    },
  },
  permission: { capability: 'test.cell.write' },
  affectedRanges: { resolve: cellRange },
  inversePolicy: { allowedMutationIds: ['cell.set'], minCount: 1 },
} as const;

test('CommandRuntime executes a registered command and tracks history', () => {
  const workbook = new WorkbookModel('unit-1', 'Runtime');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation({
    id: 'cell.set',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; value: string };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
    },
    metadata: cellSetMetadata,
  });
  runtime.registry.registerMutation({
    id: 'cell.restore',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; previous?: { value: string } };
      if (params.previous) context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, params.previous);
      else context.workbook.getSheet(item.sheetId).cells.delete(params.row, params.column);
    },
    metadata: cellRestoreMetadata,
  });
  runtime.registry.registerCommand({
    id: 'cell.set',
    execute: (params: { row: number; column: number; value: string }, context) => {
      const sheet = context.workbook.getSheet('sheet-1');
      const previous = sheet.cells.get(params.row, params.column);
      const range = [{ sheetId: 'sheet-1', startRow: params.row, endRow: params.row, startColumn: params.column, endColumn: params.column }];
      context.applyMutation({
        id: 'cell.set',
        unitId: context.workbook.unitId,
        sheetId: 'sheet-1',
        params,
        affectedRanges: range,
        inverse: [{ id: 'cell.restore', unitId: context.workbook.unitId, sheetId: 'sheet-1', params: { row: params.row, column: params.column, previous }, affectedRanges: range }],
        apply: () => sheet.cells.set(params.row, params.column, { value: params.value }),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: range };
    },
  });

  const listenedMutations: MutationInfo[] = [];
  const unsubscribe = runtime.onMutation((m) => listenedMutations.push(m));

  const result = runtime.execute('cell.set', { row: 1, column: 1, value: 'A' });
  assert.equal(result.mutationCount, 1);
  assert.equal(listenedMutations.length, 1);
  assert.equal(runtime.getHistoryDepth().undo, 1);
  assert.equal(runtime.undo(), true);
  assert.equal(workbook.getSheet('sheet-1').cells.get(1, 1), undefined);
  assert.equal(runtime.redo(), true);
  assert.equal(workbook.getSheet('sheet-1').cells.get(1, 1)?.value, 'A');

  unsubscribe();
});

test('CommandRuntime rolls back applied mutations if a command throws mid-execution', () => {
  const workbook = new WorkbookModel('unit-rollback', 'Rollback');
  const runtime = new CommandRuntime(workbook);

  runtime.registry.registerMutation({
    id: 'val.set',
    handler: (item, context) => {
      const params = item.params as { row: number; value: number };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, 0, { value: params.value });
    },
    metadata: {
      schema: { name: 'ValueSet', validate: (value: unknown) => !!value && typeof value === 'object' && Number.isInteger((value as { row?: unknown }).row) },
      permission: { capability: 'test.value.write' },
      affectedRanges: { resolve: () => [] },
      inversePolicy: { allowedMutationIds: ['val.restore'], minCount: 1 },
    },
  });
  runtime.registry.registerMutation({
    id: 'val.restore',
    handler: (item, context) => {
      const params = item.params as { row: number };
      context.workbook.getSheet(item.sheetId).cells.delete(params.row, 0);
    },
    metadata: {
      schema: { name: 'ValueRestore', validate: (value: unknown) => !!value && typeof value === 'object' && Number.isInteger((value as { row?: unknown }).row) },
      permission: { capability: 'test.value.write' },
      affectedRanges: { resolve: () => [] },
      inversePolicy: { allowedMutationIds: ['val.set'], minCount: 1 },
    },
  });

  runtime.registry.registerCommand({
    id: 'failing.transaction',
    execute: (_params: unknown, context) => {
      const sheet = context.workbook.getSheet('sheet-1');
      context.applyMutation({
        id: 'val.set',
        unitId: context.workbook.unitId,
        sheetId: 'sheet-1',
        params: { row: 0, value: 100 },
        affectedRanges: [],
        inverse: [{ id: 'val.restore', unitId: context.workbook.unitId, sheetId: 'sheet-1', params: { row: 0 }, affectedRanges: [] }],
        apply: () => sheet.cells.set(0, 0, { value: 100 }),
      });

      // Now throw an error intentionally
      throw new Error('Simulated failure during multi-mutation command');
    },
  });

  assert.throws(() => runtime.execute('failing.transaction', {}), /Simulated failure/);
  // The first mutation should have been rolled back
  assert.equal(workbook.getSheet('sheet-1').cells.get(0, 0), undefined);
  assert.equal(runtime.getHistoryDepth().undo, 0);
});

test('CommandRegistry guards against duplicate IDs and unknown lookups', () => {
  const workbook = new WorkbookModel('unit-guard', 'Guards');
  const runtime = new CommandRuntime(workbook);

  runtime.registry.registerCommand({ id: 'cmd.1', execute: () => ({ operationId: '1', mutationCount: 0, affectedRanges: [] }) });
  assert.throws(() => runtime.registry.registerCommand({ id: 'cmd.1', execute: () => ({ operationId: '1', mutationCount: 0, affectedRanges: [] }) }), /Duplicate command/);
  assert.throws(() => runtime.execute('non.existent', {}), /Unknown command/);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('remote mutations reject a different workbook unit', () => {
  const workbook = new WorkbookModel('unit-remote', 'Remote');
  const runtime = new CommandRuntime(workbook);
  runtime.registry.registerMutation({
    id: 'cell.set',
    handler: (item, context) => {
      const params = item.params as { row: number; column: number; value: string };
      context.workbook.getSheet(item.sheetId).cells.set(params.row, params.column, { value: params.value });
    },
    metadata: cellSetMetadata,
  });
  runtime.registry.registerMutation({
    id: 'cell.restore',
    handler: () => undefined,
    metadata: cellRestoreMetadata,
  });

  assert.throws(() => runtime.applyRemoteMutations([{
    id: 'cell.set',
    unitId: 'other-unit',
    sheetId: 'sheet-1',
    params: { row: 0, column: 0, value: 'invalid' },
    affectedRanges: [],
  }]), /Mutation unit mismatch/);
});

test('CommandRuntime rejects an unregistered mutation before touching the workbook', () => {
  const workbook = new WorkbookModel('unit-unregistered', 'Unregistered');
  const runtime = new CommandRuntime(workbook);
  let applyCalled = false;
  runtime.registry.registerCommand({
    id: 'invalid.mutation',
    execute: (_params: unknown, context) => {
      const affectedRanges = [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
      context.applyMutation({
        id: 'mutation.not.registered',
        unitId: workbook.unitId,
        sheetId: 'sheet-1',
        params: {},
        affectedRanges,
        inverse: [],
        apply: () => {
          applyCalled = true;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  assert.throws(() => runtime.execute('invalid.mutation', {}), /Unknown mutation: mutation\.not\.registered/);
  assert.equal(applyCalled, false);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('CommandRuntime rejects an inverse that is not registered before applying the mutation', () => {
  const workbook = new WorkbookModel('unit-invalid-inverse', 'Invalid inverse');
  const runtime = new CommandRuntime(workbook);
  let applyCalled = false;
  runtime.registry.registerMutation({
    id: 'primary.set',
    handler: () => undefined,
    metadata: {
      schema: { name: 'PrimarySet', validate: (value: unknown) => !!value && typeof value === 'object' },
      permission: { capability: 'test.write' },
      affectedRanges: { resolve: () => [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }] },
      inversePolicy: { allowedMutationIds: ['known.inverse'], minCount: 1 },
    },
  });
  runtime.registry.registerMutation({
    id: 'known.inverse',
    handler: () => undefined,
    metadata: {
      schema: { name: 'KnownInverse', validate: (value: unknown) => !!value && typeof value === 'object' },
      permission: { capability: 'test.write' },
      affectedRanges: { resolve: () => [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }] },
      inversePolicy: { allowedMutationIds: ['primary.set'], minCount: 1 },
    },
  });
  runtime.registry.registerCommand({
    id: 'invalid.inverse',
    execute: (_params: unknown, context) => {
      const affectedRanges = [{ sheetId: 'sheet-1', startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }];
      context.applyMutation({
        id: 'primary.set',
        unitId: workbook.unitId,
        sheetId: 'sheet-1',
        params: {},
        affectedRanges,
        inverse: [{
          id: 'inverse.not.registered',
          unitId: workbook.unitId,
          sheetId: 'sheet-1',
          params: {},
          affectedRanges,
        }],
        apply: () => {
          applyCalled = true;
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });

  assert.throws(() => runtime.execute('invalid.inverse', {}), /unknown inverse inverse\.not\.registered/);
  assert.equal(applyCalled, false);
  assert.deepEqual(runtime.getHistoryDepth(), { undo: 0, redo: 0 });
});

test('CommandRegistry validates schema, permission, affected ranges, and declared inverses', () => {
  const registry = new CommandRegistry();
  const range = { sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 };
  const metadata = {
    schema: { name: 'EmptyParams', validate: (value: unknown) => value !== null && typeof value === 'object' },
    permission: { capability: 'sheet.write' },
    affectedRanges: { resolve: () => [range] },
    inverseIds: ['cell.restore'],
  } as const;
  registry.registerMutation('cell.set', () => undefined, metadata);
  registry.registerMutation({
    id: 'cell.restore',
    handler: () => undefined,
    metadata: {
      schema: { name: 'EmptyParams', validate: (value: unknown) => value !== null && typeof value === 'object' },
      permission: { capability: 'sheet.write' },
      affectedRanges: { resolve: () => [range] },
      inverseIds: ['cell.set'],
    },
  });

  const result = registry.validateCompleteness();
  assert.equal(result.ok, true);
  assert.deepEqual(result.issues, []);

  const invalid = registry.validateMutation({
    id: 'cell.set',
    unitId: 'unit-1',
    sheetId: 'sheet-1',
    params: {},
    affectedRanges: [],
    inverse: [{ id: 'cell.restore', unitId: 'unit-1', sheetId: 'sheet-1', params: {}, affectedRanges: [range] }],
    apply: () => undefined,
  });
  assert.equal(invalid.some((entry) => entry.code === 'invalid-affected-ranges'), true);
});

test('CommandRegistry rejects incomplete metadata and declared inverse drift', () => {
  const registry = new CommandRegistry();
  assert.throws(() => registry.registerMutation('missing.contract', () => undefined), /requires canonical metadata/);
  assert.throws(() => registry.registerMutation({
    id: 'broken.registration',
    handler: () => undefined,
    metadata: {} as never,
  }), /must declare a parameter schema/);

  registry.registerMutation({
    id: 'declared.drift',
    handler: () => undefined,
    metadata: {
      schema: { name: 'DeclaredDrift', validate: () => true },
      permission: { capability: 'test.write' },
      affectedRanges: { resolve: () => [] },
      inverseIds: ['missing.inverse'],
    },
  });
  const result = registry.validateCompleteness();
  assert.equal(result.ok, false);
  assert.equal(result.issues.some((entry) => entry.code === 'unknown-inverse' && entry.inverseId === 'missing.inverse'), true);
  assert.throws(() => registry.assertComplete(), /Mutation registry is incomplete/);
});
