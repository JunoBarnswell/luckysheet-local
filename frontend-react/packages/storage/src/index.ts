import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import type { CollaborationChangeSet, SnapshotResponse } from '@react-sheets/protocol';
import { WorkbookModel } from '@react-sheets/core-model';
import type { WorkbookSnapshotV1 } from '@react-sheets/core-model';

const databasePath = resolve(process.cwd(), 'data/react-sheets.sqlite');
mkdirSync(dirname(databasePath), { recursive: true });

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

  createWorkbook(snapshot: WorkbookSnapshotV1): SnapshotResponse {
    const now = new Date().toISOString();
    this.database.prepare(`
      INSERT INTO workbooks (unit_id, name, snapshot_json, revision, updated_at)
      VALUES (?, ?, ?, 0, ?)
      ON CONFLICT(unit_id) DO NOTHING
    `).run(snapshot.unitId, snapshot.name, JSON.stringify(snapshot), now);
    return this.getSnapshot(snapshot.unitId);
  }

  getSnapshot(unitId: string): SnapshotResponse {
    const row = this.database.prepare('SELECT snapshot_json, revision FROM workbooks WHERE unit_id = ?').get(unitId) as { snapshot_json?: string; revision?: number } | undefined;
    if (!row?.snapshot_json || row.revision == null) throw new Error(`Workbook not found: ${unitId}`);
    return { snapshot: JSON.parse(row.snapshot_json) as WorkbookSnapshotV1, revision: row.revision };
  }

  appendChangeSet(changeSet: CollaborationChangeSet): number {
    const current = this.getSnapshot(changeSet.unitId);
    const existing = this.database.prepare('SELECT revision FROM changesets WHERE operation_id = ?').get(changeSet.operationId) as { revision?: number } | undefined;
    if (existing?.revision != null) return existing.revision;
    if (current.revision !== changeSet.baseRevision) throw new Error('Revision conflict');
    const nextRevision = current.revision + 1;
    const workbook = WorkbookModel.fromSnapshot(current.snapshot);
    for (const mutation of changeSet.mutations) {
      if (mutation.id === 'cell.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as { row: number; column: number; value: unknown };
        sheet.cells.set(params.row, params.column, params.value as never);
      } else if (mutation.id === 'range.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as { startRow: number; startColumn: number; values: unknown[][] };
        params.values.forEach((row, rowOffset) => row.forEach((value, columnOffset) => {
          sheet.cells.set(params.startRow + rowOffset, params.startColumn + columnOffset, value as never);
        }));
      } else if (mutation.id === 'sheet.add') {
        const params = mutation.params as { id: string; name: string };
        workbook.addSheet(params.id, params.name);
      } else {
        throw new Error(`Unsupported mutation: ${mutation.id}`);
      }
    }
    const nextSnapshot = workbook.snapshot();
    this.database.prepare(`
      INSERT INTO changesets (operation_id, unit_id, revision, payload_json, created_at)
      VALUES (?, ?, ?, ?, ?)
    `).run(changeSet.operationId, changeSet.unitId, nextRevision, JSON.stringify(changeSet), changeSet.createdAt);
    this.database.prepare('UPDATE workbooks SET snapshot_json = ?, revision = ?, updated_at = ? WHERE unit_id = ?').run(JSON.stringify(nextSnapshot), nextRevision, new Date().toISOString(), changeSet.unitId);
    return nextRevision;
  }

  listRevisions(unitId: string): Array<{ operationId: string; revision: number; createdAt: string; payload: CollaborationChangeSet }> {
    const rows = this.database.prepare('SELECT operation_id, revision, created_at, payload_json FROM changesets WHERE unit_id = ? ORDER BY revision DESC').all(unitId) as Array<{ operation_id: string; revision: number; created_at: string; payload_json: string }>;
    return rows.map((row) => ({ operationId: row.operation_id, revision: row.revision, createdAt: row.created_at, payload: JSON.parse(row.payload_json) as CollaborationChangeSet }));
  }

  getRevision(unitId: string, revision: number): { operationId: string; revision: number; createdAt: string; payload: CollaborationChangeSet } | undefined {
    const row = this.database.prepare('SELECT operation_id, revision, created_at, payload_json FROM changesets WHERE unit_id = ? AND revision = ?').get(unitId, revision) as { operation_id?: string; revision?: number; created_at?: string; payload_json?: string } | undefined;
    if (!row?.operation_id || row.revision == null || !row.created_at || !row.payload_json) return undefined;
    return { operationId: row.operation_id, revision: row.revision, createdAt: row.created_at, payload: JSON.parse(row.payload_json) as CollaborationChangeSet };
  }
}
