import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import type { RevisionRecord } from '@react-sheets/protocol';
import { createLocalCalculationSessionPort } from '@react-sheets/formula-engine';
import {
  buildRestoreParams,
  describeRevisionMutations,
  revisionToHistoryMeta,
} from './replay';
import { HistoryPreviewSession } from './index';

describe('history replay', () => {
  it('maps revision records to history metadata', () => {
    const record: RevisionRecord = {
      operationId: 'op-1',
      revision: 2,
      createdAt: '2026-01-01T00:00:00.000Z',
      payload: {
        schema: 'OperationEnvelope',
        operationId: 'op-1',
        unitId: 'wb-1',
        actorId: 'actor-1',
        origin: 'client',
        clientSequence: 1,
        baseRevision: 1,
        revision: 2,
        committedAt: '2026-01-01T00:00:00.000Z',
        createdAt: '2026-01-01T00:00:00.000Z',
        mutations: [{
          id: 'cell.set',
          sheetId: 'sheet-1',
          params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'x' } },
          affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
        }],
      },
    };

    const meta = revisionToHistoryMeta(record);
    assert.equal(meta.revision, 2);
    assert.equal(meta.actorId, 'actor-1');
    assert.match(describeRevisionMutations(record), /cell\.set/);
    const restore = buildRestoreParams(2, 'restore test');
    assert.deepEqual(restore, { targetRevision: 2, reason: 'restore test' });
    assert.equal('snapshot' in restore, false);
  });

  it('builds an isolated formula-backed UI projection for preview', async () => {
    const source = new WorkbookModel('wb-preview', 'Preview');
    const sheet = source.getSheet('sheet-1');
    sheet.cells.set(0, 0, { value: 2 });
    sheet.cells.set(0, 1, { value: null, formula: '=A1+3' });
    const session = await HistoryPreviewSession.fromSnapshot({
      revision: 7,
      operationId: 'op-preview',
      createdAt: '2026-01-01T00:00:00.000Z',
      description: 'preview',
    }, source.snapshot(), undefined, createLocalCalculationSessionPort());

    assert.notEqual(session.workbook, source);
    assert.equal(session.formula.getCellValue({ sheetId: 'sheet-1', row: 0, column: 1 }), 5);
    assert.equal(session.getSheet('sheet-1')?.getCell(0, 1)?.value, '5');
    assert.equal(session.ui.activeSheetId, 'sheet-1');
    session.dispose();
    assert.throws(() => session.sheets, /disposed/);
  });
});
