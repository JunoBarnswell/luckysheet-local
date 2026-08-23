import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { WorkbookModel } from '@react-sheets/core-model';
import type { RevisionRecord } from '@react-sheets/protocol';
import {
  buildRestoreParams,
  describeRevisionMutations,
  replayRevisionsToSnapshot,
  revisionToHistoryMeta,
} from './history-bridge';

describe('history-bridge', () => {
  it('maps revision records to history metadata', () => {
    const record: RevisionRecord = {
      operationId: 'op-1',
      revision: 2,
      createdAt: '2026-01-01T00:00:00.000Z',
      payload: {
        schema: 'OperationEnvelopeV2',
        operationId: 'op-1',
        unitId: 'wb-1',
        actorId: 'actor-1',
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
    assert.equal(buildRestoreParams({ schema: 'WorkbookSnapshotV1', unitId: 'wb-1', name: 'Test', activeSheetId: 'sheet-1', sheets: [], tables: [], definedNames: {} }, 2).targetRevision, 2);
  });

  it('replays revision mutations onto a base snapshot', () => {
    const baseSnapshot = new WorkbookModel('wb-1', 'Replay').snapshot();
    const revisions: RevisionRecord[] = [{
      operationId: 'op-1',
      revision: 1,
      createdAt: '2026-01-01T00:00:00.000Z',
      payload: {
        schema: 'OperationEnvelopeV2',
        operationId: 'op-1',
        unitId: 'wb-1',
        actorId: 'actor-1',
        clientSequence: 1,
        baseRevision: 0,
        revision: 1,
        committedAt: '2026-01-01T00:00:00.000Z',
        createdAt: '2026-01-01T00:00:00.000Z',
        mutations: [{
          id: 'cell.set',
          sheetId: 'sheet-1',
          params: { sheetId: 'sheet-1', row: 0, column: 0, value: { value: 'restored' } },
          affectedRanges: [{ sheetId: 'sheet-1', startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
        }],
      },
    }];

    const replayed = replayRevisionsToSnapshot(baseSnapshot, revisions, 1);
    const sheet = replayed.sheets.find((entry) => entry.id === baseSnapshot.activeSheetId);
    assert.equal(sheet?.cells['0']?.['0']?.value, 'restored');
  });
});
