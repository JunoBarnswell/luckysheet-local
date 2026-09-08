import assert from 'node:assert/strict';
import test from 'node:test';
import { initializeNodeKernel } from '@react-sheets/kernel-client/node';
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
      pages: [],
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
      pages: [],
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
