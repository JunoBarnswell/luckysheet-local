import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { CollaborationChangeSet, SnapshotResponse, WorkbookSummary } from '@react-sheets/protocol';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerProSheetCommands } from '@react-sheets/pro-features';

const databasePath = resolve(process.cwd(), 'data/react-sheets.sqlite');
mkdirSync(dirname(databasePath), { recursive: true });

function createMutationRuntime(workbook: WorkbookModel): CommandRuntime {
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  registerProSheetCommands(runtime);
  return runtime;
}

export class WorkbookStorage {
  private readonly database = new DatabaseSync(databasePath);

  constructor() {
    this.database.exec(`
      PRAGMA journal_mode = WAL;
      CREATE TABLE IF NOT EXISTS workbooks (
        unit_id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        snapshot_json TEXT NOT NULL,
        revision INTEGER NOT NULL DEFAULT 0,
        updated_at TEXT NOT NULL
      );
      CREATE TABLE IF NOT EXISTS changesets (
        operation_id TEXT PRIMARY KEY,
        unit_id TEXT NOT NULL,
        revision INTEGER NOT NULL,
        payload_json TEXT NOT NULL,
        created_at TEXT NOT NULL
      );
    `);
  }

  listWorkbooks(): WorkbookSummary[] {
    const rows = this.database
      .prepare('SELECT unit_id, name, revision, updated_at FROM workbooks ORDER BY updated_at DESC')
      .all() as Array<{ unit_id: string; name: string; revision: number; updated_at: string }>;
    return rows.map((row) => ({
      unitId: row.unit_id,
      name: row.name,
      revision: row.revision,
      updatedAt: row.updated_at,
    }));
  }

  createWorkbook(snapshot: WorkbookSnapshotV1): SnapshotResponse {
    const now = new Date().toISOString();
    this.database
      .prepare(
        `
      INSERT INTO workbooks (unit_id, name, snapshot_json, revision, updated_at)
      VALUES (?, ?, ?, 0, ?)
      ON CONFLICT(unit_id) DO NOTHING
    `,
      )
      .run(snapshot.unitId, snapshot.name, JSON.stringify(snapshot), now);
    return this.getSnapshot(snapshot.unitId);
  }

  getSnapshot(unitId: string): SnapshotResponse {
    const row = this.database
      .prepare('SELECT snapshot_json, revision FROM workbooks WHERE unit_id = ?')
      .get(unitId) as { snapshot_json?: string; revision?: number } | undefined;
    if (!row?.snapshot_json || row.revision == null) {
      throw new Error(`Workbook not found: ${unitId}`);
    }
    return {
      snapshot: JSON.parse(row.snapshot_json) as WorkbookSnapshotV1,
      revision: row.revision,
    };
  }

  appendChangeSet(changeSet: CollaborationChangeSet): number {
    const current = this.getSnapshot(changeSet.unitId);
    const existing = this.database
      .prepare('SELECT revision FROM changesets WHERE operation_id = ?')
      .get(changeSet.operationId) as { revision?: number } | undefined;
    if (existing?.revision != null) return existing.revision;
    if (current.revision !== changeSet.baseRevision) throw new Error('Revision conflict');

    const nextRevision = current.revision + 1;
    const workbook = WorkbookModel.fromSnapshot(current.snapshot);
    const runtime = createMutationRuntime(workbook);
    runtime.applyRemoteMutations(
      changeSet.mutations.map((mutation) => ({
        ...mutation,
        unitId: changeSet.unitId,
      })),
    );

    const nextSnapshot = workbook.snapshot();
    this.database
      .prepare(
        `
      INSERT INTO changesets (operation_id, unit_id, revision, payload_json, created_at)
      VALUES (?, ?, ?, ?, ?)
    `,
      )
      .run(
        changeSet.operationId,
        changeSet.unitId,
        nextRevision,
        JSON.stringify(changeSet),
        changeSet.createdAt,
      );
    this.database
      .prepare(
        'UPDATE workbooks SET snapshot_json = ?, revision = ?, updated_at = ? WHERE unit_id = ?',
      )
      .run(JSON.stringify(nextSnapshot), nextRevision, new Date().toISOString(), changeSet.unitId);
    return nextRevision;
  }

  listRevisions(
    unitId: string,
  ): Array<{
    operationId: string;
    revision: number;
    createdAt: string;
    payload: CollaborationChangeSet;
  }> {
    const rows = this.database
      .prepare(
        'SELECT operation_id, revision, created_at, payload_json FROM changesets WHERE unit_id = ? ORDER BY revision DESC',
      )
      .all(unitId) as Array<{
      operation_id: string;
      revision: number;
      created_at: string;
      payload_json: string;
    }>;
    return rows.map((row) => ({
      operationId: row.operation_id,
      revision: row.revision,
      createdAt: row.created_at,
      payload: JSON.parse(row.payload_json) as CollaborationChangeSet,
    }));
  }

  getRevision(
    unitId: string,
    revision: number,
  ):
    | {
        operationId: string;
        revision: number;
        createdAt: string;
        payload: CollaborationChangeSet;
      }
    | undefined {
    const row = this.database
      .prepare(
        'SELECT operation_id, revision, created_at, payload_json FROM changesets WHERE unit_id = ? AND revision = ?',
      )
      .get(unitId, revision) as
      | {
          operation_id?: string;
          revision?: number;
          created_at?: string;
          payload_json?: string;
        }
      | undefined;
    if (!row?.operation_id || row.revision == null || !row.created_at || !row.payload_json) {
      return undefined;
    }
    return {
      operationId: row.operation_id,
      revision: row.revision,
      createdAt: row.created_at,
      payload: JSON.parse(row.payload_json) as CollaborationChangeSet,
    };
  }
}
