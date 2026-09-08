import assert from 'node:assert/strict';
import test from 'node:test';
import { kernelInvoke } from '@react-sheets/kernel-client';
import { initializeNodeKernel } from '@react-sheets/kernel-client/node';
import type { KernelPagePayload, WorkbookManifest } from '@react-sheets/protocol';
import { WorkbookSession } from './workbook-session';

test('cloud session construction defers manifest-derived persistence metadata until open completes', async () => {
  await initializeNodeKernel();
  const manifest = {
    schema: 'WorkbookManifest' as const,
    version: 11 as const,
    unitId: 'unit-loading',
    name: 'Resolved workbook',
    revision: 4,
    sheets: [{ sheetId: 'imported-sheet', name: 'Imported', rowCount: 128, columnCount: 16, metadata: {} }],
    pages: [],
    metadata: {},
  };
  const session = new WorkbookSession({
    initialPhase: 'loading',
    resolution: {
      schema: 'WorkbookResolution',
      unitId: manifest.unitId,
      source: 'remote',
      mode: 'remote',
      lifecycle: 'active',
      manifest,
      revision: manifest.revision,
      access: { unitId: manifest.unitId, role: 'owner', nextClientSequence: 1 },
    },
  });
  try {
    const snapshot = session.getUiSnapshot();
    assert.equal(snapshot.unitId, 'unit-loading');
    assert.equal(snapshot.workbookName, 'Resolved workbook');
    assert.equal(snapshot.activeSheetId, 'imported-sheet');
    assert.equal(snapshot.phase, 'loading');
    assert.equal(snapshot.persistenceChecksum, '');
  } finally {
    session.dispose();
  }
});

test('collaboration socket closure does not revoke the authoritative HTTP edit session', async () => {
  await initializeNodeKernel();
  const manifest = {
    schema: 'WorkbookManifest' as const,
    version: 11 as const,
    unitId: 'unit-collaboration-status',
    name: 'Connected workbook',
    revision: 0,
    sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 128, columnCount: 16, metadata: {} }],
    pages: [],
    metadata: {},
  };
  const session = new WorkbookSession({
    initialPhase: 'ready',
    resolution: {
      schema: 'WorkbookResolution',
      unitId: manifest.unitId,
      source: 'remote',
      mode: 'remote',
      lifecycle: 'active',
      manifest,
      revision: manifest.revision,
      access: { unitId: manifest.unitId, role: 'owner', nextClientSequence: 1 },
    },
  });
  try {
    const runtime = session['runtime'];
    runtime.remoteConnected = true;
    runtime.handlers.onAccessRole?.('owner');
    assert.equal(session.getUiSnapshot().permissions.editCell, true);

    runtime.handlers.onCollabStatus?.('closed');

    assert.equal(runtime.remoteConnected, true);
    assert.equal(session.getUiSnapshot().collabStatus, 'closed');
    assert.equal(session.getUiSnapshot().permissions.editCell, true);
  } finally {
    session.dispose();
  }
});

test('cloud open loads only the seed page and fetches a distant viewport on demand', async () => {
  await initializeNodeKernel();
  const unitId = 'unit-lazy-open';
  kernelInvoke('create', {
    unitId,
    name: 'Lazy open',
    sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 4096, columnCount: 64, metadata: {} }],
  });
  const committed = kernelInvoke<{ manifest: WorkbookManifest; pages: KernelPagePayload[] }>('command', {
    unitId,
    baseRevision: 0,
    operationId: `${unitId}:seed`,
    commandId: 'operation.apply',
    accessRole: 'owner',
    params: { mutations: [
      { id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'first' } } },
      { id: 'cell.set', sheetId: 'sheet-1', params: { sheetId: 'sheet-1', row: 3072, column: 0, value: { value: 'distant' } } },
    ] },
  });
  const pages = new Map(committed.pages.map((page) => [`${page.pageRow}:${page.pageColumn}`, page]));
  const requested: string[] = [];
  kernelInvoke('close', { unitId });
  const session = new WorkbookSession({
    initialPhase: 'loading',
    api: {
      getPage: async (request: { pageRow: number; pageColumn: number }) => {
        requested.push(`${request.pageRow}:${request.pageColumn}`);
        const page = pages.get(`${request.pageRow}:${request.pageColumn}`);
        if (!page) throw new Error('PAGE_NOT_FOUND');
        return structuredClone(page);
      },
    } as never,
    resolution: {
      schema: 'WorkbookResolution',
      unitId,
      source: 'remote',
      mode: 'remote',
      lifecycle: 'active',
      manifest: committed.manifest,
      revision: committed.manifest.revision,
      access: { unitId, role: 'owner', nextClientSequence: 1 },
    },
  });
  try {
    assert.equal(session.getUiSnapshot().phase, 'loading', 'manifest-only construction must be renderable before the seed page arrives');
    session.start();
    await session['runtime'].persistenceReady;
    assert.deepEqual(requested, ['0:0']);
    assert.equal(session.getUiSnapshot().selectedSheet.getCell(0, 0)?.rawValue, 'first');

    await session.ensureVisibleRanges([{ sheetId: 'sheet-1', startRow: 3072, endRow: 3072, startColumn: 0, endColumn: 0 }]);
    assert.deepEqual(requested, ['0:0', '3:0']);
    assert.equal(session.getUiSnapshot().selectedSheet.getCell(3072, 0)?.rawValue, 'distant');
  } finally {
    session.dispose();
    kernelInvoke('close', { unitId });
  }
});

test('cloud hydrate preserves the session calculation date required by the cell-edit input contract', async () => {
  await initializeNodeKernel();
  const unitId = 'unit-edit-reference-date';
  const manifest: WorkbookManifest = {
    schema: 'WorkbookManifest',
    version: 11,
    unitId,
    name: 'Editable cloud workbook',
    revision: 0,
    sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 128, columnCount: 16, metadata: {} }],
    pages: [],
    metadata: {},
  };
  const referenceDate = { year: 2026, month: 9, day: 8, hour: 7, minute: 30, second: 0, millisecond: 0 };
  const session = new WorkbookSession({
    initialPhase: 'loading',
    canonicalReferenceDate: referenceDate,
    resolution: {
      schema: 'WorkbookResolution',
      unitId,
      source: 'remote',
      mode: 'remote',
      lifecycle: 'active',
      manifest,
      revision: manifest.revision,
      access: { unitId, role: 'owner', nextClientSequence: 1 },
    },
  });
  try {
    session.start();
    await session['runtime'].persistenceReady;

    assert.deepEqual(session['runtime'].canonicalReferenceDate, referenceDate);
    assert.doesNotThrow(() => session['createInputContext']('direct-entry'));
  } finally {
    session.dispose();
    kernelInvoke('close', { unitId });
  }
});
