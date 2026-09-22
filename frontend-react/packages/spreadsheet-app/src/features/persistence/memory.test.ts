import assert from 'node:assert/strict';
import test from 'node:test';
import { WorkspaceMemoryCoordinator, type WorkspaceMemoryTransaction } from './memory';

test('writes do not clone untouched block bytes or native documents', async (context) => {
  const memory = new WorkspaceMemoryCoordinator();
  await memory.transaction((transaction) => {
    transaction.set('dataBlocks', 'cold', { id: 'cold-block', bytes: new ArrayBuffer(1024 * 1024) });
    transaction.set('nativeDocuments', 'cold', { id: 'cold-document', bytes: new ArrayBuffer(1024 * 1024) });
  });
  const actualClone = globalThis.structuredClone;
  const copied: unknown[] = [];
  // Observe the real clone operation; do not replace its behavior.
  context.mock.method(globalThis, 'structuredClone', <T>(value: T): T => {
    copied.push(value);
    return actualClone(value);
  });
  await memory.transaction((transaction) => transaction.set('workspaceCatalog', 'unit', { name: 'renamed' }));
  await memory.transaction((transaction) => transaction.set('dataBlocks', 'new', { id: 'new-block', bytes: new ArrayBuffer(8) }));
  await memory.transaction((transaction) => transaction.delete('dataBlocks', 'new'));
  assert.equal(copied.some((value) => value !== null && typeof value === 'object' && 'id' in value
    && (value.id === 'cold-block' || value.id === 'cold-document')), false);
  assert.equal(await memory.read((reader) => reader.get<{ bytes: ArrayBuffer }>('dataBlocks', 'cold')?.bytes.byteLength), 1024 * 1024);
});

test('successful and failed reads cannot mutate stored values', async () => {
  const memory = new WorkspaceMemoryCoordinator();
  await memory.transaction((transaction) => transaction.set('workspaceCatalog', 'unit', { nested: { value: 1 } }));
  await memory.read((reader) => {
    reader.get<{ nested: { value: number } }>('workspaceCatalog', 'unit')!.nested.value = 2;
    reader.getAll<{ nested: { value: number } }>('workspaceCatalog')[0]!.nested.value = 3;
    assert.equal('set' in reader, false);
    assert.equal('delete' in reader, false);
  });
  await assert.rejects(memory.read((reader) => {
    reader.get<{ nested: { value: number } }>('workspaceCatalog', 'unit')!.nested.value = 4;
    throw new Error('read failed');
  }), { code: 'STORAGE_MEMORY_TRANSACTION_FAILED' });
  assert.deepEqual(await memory.read((reader) => reader.get('workspaceCatalog', 'unit')), { nested: { value: 1 } });
});

test('explicit multi-bucket writes commit atomically and do not retain caller aliases', async () => {
  const memory = new WorkspaceMemoryCoordinator();
  const input = { nested: { value: 1 } };
  const result = await memory.transaction((transaction) => {
    transaction.set('workspaceCatalog', 'unit', input);
    input.nested.value = 2;
    const value = transaction.get<typeof input>('workspaceCatalog', 'unit')!;
    value.nested.value = 3;
    assert.equal(transaction.get<typeof input>('workspaceCatalog', 'unit')!.nested.value, 1);
    transaction.set('workspaceCatalog', 'unit', value);
    transaction.set('workspaceHeads', 'unit', { revision: 1 });
    transaction.set('workspaceHeads', 'temporary', { revision: 0 });
    transaction.delete('workspaceHeads', 'temporary');
    return value;
  });
  result.nested.value = 4;
  input.nested.value = 5;
  assert.deepEqual(await memory.read((reader) => reader.get('workspaceCatalog', 'unit')), { nested: { value: 3 } });
  assert.deepEqual(await memory.read((reader) => reader.getAll('workspaceHeads')), [{ revision: 1 }]);
});

test('uncloneable results and values roll back all staged writes', async () => {
  const memory = new WorkspaceMemoryCoordinator();
  await memory.transaction((transaction) => transaction.set('workspaceCatalog', 'unit', { value: 'before' }));
  await assert.rejects(memory.transaction((transaction) => {
    transaction.set('workspaceCatalog', 'unit', { value: 'after' });
    transaction.set('workspaceHeads', 'unit', { revision: 1 });
    return () => 'cannot clone this result';
  }), { code: 'STORAGE_MEMORY_TRANSACTION_FAILED' });
  await assert.rejects(memory.transaction((transaction) => {
    transaction.delete('workspaceCatalog', 'unit');
    transaction.set('workspaceHeads', 'unit', { callback: () => undefined });
  }), { code: 'STORAGE_MEMORY_TRANSACTION_FAILED' });
  assert.deepEqual(await memory.read((reader) => reader.get('workspaceCatalog', 'unit')), { value: 'before' });
  assert.equal(await memory.read((reader) => reader.get('workspaceHeads', 'unit')), undefined);
  await memory.transaction((transaction) => transaction.set('workspaceHeads', 'unit', { revision: 2 }));
  assert.deepEqual(await memory.read((reader) => reader.get('workspaceHeads', 'unit')), { revision: 2 });
});

test('a retained transaction handle cannot mutate committed or rolled-back state', async () => {
  const memory = new WorkspaceMemoryCoordinator();
  let retained!: WorkspaceMemoryTransaction;
  await memory.transaction((transaction) => {
    retained = transaction;
    transaction.set('workspaceCatalog', 'unit', { value: 1 });
  });
  assert.throws(() => retained.set('workspaceCatalog', 'unit', { value: 2 }), { operation: 'closed-transaction' });
  assert.throws(() => retained.delete('workspaceCatalog', 'unit'), { operation: 'closed-transaction' });
  await assert.rejects(memory.transaction((transaction) => {
    retained = transaction;
    throw new Error('abort');
  }));
  assert.throws(() => retained.get('workspaceCatalog', 'unit'), { operation: 'closed-transaction' });
  assert.deepEqual(await memory.read((reader) => reader.get('workspaceCatalog', 'unit')), { value: 1 });
});

test('an async reader retains its original snapshot while a later write commits', async () => {
  const memory = new WorkspaceMemoryCoordinator();
  await memory.transaction((transaction) => transaction.set('workspaceCatalog', 'unit', { value: 1 }));
  let signalStarted!: () => void;
  const started = new Promise<void>((resolve) => { signalStarted = resolve; });
  let resume!: () => void;
  const resumed = new Promise<void>((resolve) => { resume = resolve; });
  const read = memory.read(async (reader) => {
    signalStarted();
    await resumed;
    return reader.get('workspaceCatalog', 'unit');
  });
  await started;
  await memory.transaction((transaction) => transaction.set('workspaceCatalog', 'unit', { value: 2 }));
  resume();
  assert.deepEqual(await read, { value: 1 });
  assert.deepEqual(await memory.read((reader) => reader.get('workspaceCatalog', 'unit')), { value: 2 });
});

test('an async read cannot report success after its memory session is disposed', async () => {
  const memory = new WorkspaceMemoryCoordinator();
  let signalStarted!: () => void;
  const started = new Promise<void>((resolve) => { signalStarted = resolve; });
  let resume!: () => void;
  const resumed = new Promise<void>((resolve) => { resume = resolve; });
  const read = memory.read(async () => { signalStarted(); await resumed; return 'stale'; });
  const rejected = assert.rejects(read, { code: 'STORAGE_MEMORY_DISPOSED' });
  await started;
  await memory.disposeAsync();
  resume();
  await rejected;
});
