import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { SpreadsheetApplication } from './application';
import { buildLocalDraftRecord } from './features/persistence';

describe('SpreadsheetApplication persistence integration', () => {
  it('exposes persistence metadata in ui snapshot', () => {
    const app = new SpreadsheetApplication();
    const meta = app.getPersistenceSnapshot();
    assert.equal(meta.unitId, app.getWorkbook().unitId);
    assert.equal(meta.checksum.length, 64);
    assert.equal(app.getUiSnapshot().persistenceChecksum.length, 64);
  });

  it('writes local drafts after workbook mutations', async () => {
    const app = new SpreadsheetApplication();
    app.runCommand('sheet.cell.set', {
      sheetId: app.getActiveSheetId(),
      row: 0,
      column: 0,
      value: { value: 'draft-value' },
    });
    await new Promise((resolve) => setTimeout(resolve, 900));
    const draft = app.getPersistenceSnapshot();
    assert.equal(draft.hasLocalDraft, true);
    assert.equal(app.getUiSnapshot().hasLocalDraft, true);
  });

  it('recovers a stored local draft into the runtime', () => {
    const app = new SpreadsheetApplication();
    const sheetId = app.getActiveSheetId();
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'recovered' },
    });
    const record = buildLocalDraftRecord(app.getWorkbook().snapshot(), 9);
    app['runtime'].draftStore.write(record);
    app.runCommand('sheet.cell.set', {
      sheetId,
      row: 0,
      column: 0,
      value: { value: 'lost' },
    });
    assert.equal(app.recoverLocalDraft(), true);
    assert.equal(app.getWorkbook().getSheet(sheetId).cells.get(0, 0)?.value, 'recovered');
  });

  it('blocks persistence.save for viewers', async () => {
    const app = new SpreadsheetApplication();
    app.setShareRole('viewer');
    await app.saveWorkbook('blocked');
    assert.match(app.getUiSnapshot().notice, /viewer|save|Permission/i);
  });
});
