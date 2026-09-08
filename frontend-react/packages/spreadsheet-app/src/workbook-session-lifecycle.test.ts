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
      access: { unitId: manifest.unitId, role: 'owner' },
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
