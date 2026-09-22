import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import { ApiRequestError } from '@react-sheets/protocol';
import { WorkbookSession } from './workbook-session';

describe('WorkbookSession local workspace integration', () => {
  it('exposes a canonical persistence checksum in the UI projection', () => {
    const app = new WorkbookSession({ nativeDocumentExecution: 'inline-test' });
    const meta = app.getPersistenceSnapshot();
    assert.equal(meta.unitId, app['runtime'].model.unitId);
    assert.equal(meta.checksum.length, 64);
    assert.equal(app.getUiSnapshot().persistenceChecksum.length, 64);
  });

  it('checkpoints a local mutation as one canonical workspace record', async () => {
    const app = new WorkbookSession();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'local-value' },
    });
    await app['runtime'].checkpointWorkspace();
    const record = await app['runtime'].workspacePersistence.load(app['runtime'].model.unitId);
    assert.equal(record?.snapshot.schema, 'WorkbookSnapshot');
    assert.equal(record?.snapshot.sheets[0]?.cells['0']?.['0']?.value, 'local-value');
    assert.equal(record?.checksum.length, 64);
  });

  it('saves locally without depending on a server role projection', async () => {
    const app = new WorkbookSession({ nativeDocumentExecution: 'inline-test' });
    await app.saveWorkbook('local save');
    assert.match(app.getUiSnapshot().notice, /checkpoint saved/i);
    const artifact = await app['runtime'].workspacePersistence.nativeDocuments.load(app['runtime'].model.unitId);
    assert.equal(artifact?.format.family, 'ssjson');
    assert.equal(artifact?.fileName.endsWith('.ssjson'), true);
  });

  it('uses the session share token for remote workbook reads', async () => {
    const originalFetch = globalThis.fetch;
    const snapshot = new WorkbookModel('shared-workbook', 'Shared workbook').snapshot();
    let receivedShareToken: string | null = null;
    globalThis.fetch = async (_input, init) => {
      receivedShareToken = new Headers(init?.headers).get('x-workbook-share-token');
      return new Response(JSON.stringify({ snapshot, revision: 1 }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      });
    };
    try {
      const app = new WorkbookSession({ unitId: 'shared-workbook', shareTokenProvider: () => 'guest-share-token' });
      await app['runtime'].api.getSnapshot('shared-workbook');
      assert.equal(receivedShareToken, 'guest-share-token');
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it('surfaces a rejected session share token as an authoritative remote failure', async () => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = async () => new Response(JSON.stringify({ code: 'FORBIDDEN', message: 'Share access revoked' }), {
      status: 403,
      headers: { 'content-type': 'application/json' },
    });
    try {
      const app = new WorkbookSession({ unitId: 'revoked-workbook', shareTokenProvider: () => 'revoked-share-token' });
      await assert.rejects(
        () => app['runtime'].api.getSnapshot('revoked-workbook'),
        (error: unknown) => error instanceof ApiRequestError && error.status === 403,
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });
});
