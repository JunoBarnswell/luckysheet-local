import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';

describe('SpreadsheetApplication local workspace integration', () => {
  it('exposes a canonical persistence checksum in the UI projection', () => {
    const app = new SpreadsheetApplication();
    const meta = app.getPersistenceSnapshot();
    assert.equal(meta.unitId, app['runtime'].model.unitId);
    assert.equal(meta.checksum.length, 64);
    assert.equal(app.getUiSnapshot().persistenceChecksum.length, 64);
  });

  it('checkpoints a local mutation as one canonical workspace record', async () => {
    const app = new SpreadsheetApplication();
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
    const app = new SpreadsheetApplication();
    await app.saveWorkbook('local save');
    assert.match(app.getUiSnapshot().notice, /checkpoint saved/i);
  });
});
