import { randomUUID } from 'node:crypto';
import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { deflateRawSync, inflateRawSync } from 'node:zlib';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type PivotLayout, type TableScalar, type WorkbookTableBlock, type WorkbookTableModel, type WorkbookSnapshotV1 } from '@react-sheets/core-model';
import {
  validateOperationEnvelopeV2,
  type CommittedOperationEnvelopeV2,
  type HistoryAuditRecord,
  type HistoryRestoreResponse,
  type OperationEnvelopeV2,
  type SnapshotResponse,
  type WorkbookAclRecord,
  type WorkbookAclRole,
  type WorkbookSummary,
} from '@react-sheets/protocol';
import { registerSpreadsheetFeatures, DrawingRuntime } from '@react-sheets/spreadsheet-app';
import { computeSnapshotChecksum } from './persistence-session';

const defaultDatabasePath = resolve(process.cwd(), 'data/react-sheets.sqlite');

export interface WorkbookStorageOptions {
  /** Absolute or relative SQLite path. Useful for isolated integration tests. */
  databasePath?: string;
}

const ROLE_RANK: Record<WorkbookAclRole, number> = {
  viewer: 1,
  commenter: 2,
  editor: 3,
  owner: 4,
};

export class StorageAccessError extends Error {
  readonly code = 'FORBIDDEN' as const;
  readonly status = 403 as const;
}

export class StorageConflictError extends Error {
  readonly code = 'CONFLICT' as const;
  readonly status = 409 as const;
}

export class StorageValidationError extends Error {
  readonly code = 'VALIDATION_ERROR' as const;
  readonly status = 400 as const;
}

function requireSubject(subject: string): string {
  const normalized = subject.trim();
  if (!normalized) throw new StorageAccessError('Authenticated subject is required');
  return normalized;
}

function assertRole(role: string): asserts role is WorkbookAclRole {
  if (!(role in ROLE_RANK)) throw new Error(`Invalid workbook ACL role: ${role}`);
}

function createMutationRuntime(workbook: WorkbookModel): CommandRuntime {
  const runtime = new CommandRuntime(workbook);
  registerSpreadsheetFeatures(runtime, new DrawingRuntime());
  return runtime;
}

function migrateSnapshot(snapshot: WorkbookSnapshotV1): WorkbookSnapshotV1 {
  const migrated = structuredClone(snapshot) as WorkbookSnapshotV1;
  for (const sheet of migrated.sheets) {
    for (const pivot of sheet.pivots) {
      const legacy = pivot as unknown as Record<string, unknown>;
      if (!legacy.layout) {
        const rows = Array.isArray(legacy.rowFields) ? legacy.rowFields.map((field) => ({ field: String(field) })) : [];
        const columns = Array.isArray(legacy.columnFields) ? legacy.columnFields.map((field) => ({ field: String(field) })) : [];
        const values = Array.isArray(legacy.valueFields) ? legacy.valueFields as PivotLayout['values'] : [];
        legacy.layout = {
          rows,
          columns,
          filters: [],
          values,
          showSubtotals: true,
          showGrandTotals: true,
          compact: false,
          repeatLabels: false,
        } satisfies PivotLayout;
      }
      delete legacy.rowFields;
      delete legacy.columnFields;
      delete legacy.valueFields;
      delete legacy.filterFields;
    }
  }
  return migrated;
}

interface EncodedTableBlock {
  fields: string[];
  columns: Record<string, TableScalar[]>;
}

export class WorkbookStorage {
  private readonly database: DatabaseSync;

  constructor(options: WorkbookStorageOptions = {}) {
    const databasePath = resolve(options.databasePath ?? process.env.REACT_SHEETS_DATABASE_PATH ?? defaultDatabasePath);
    mkdirSync(dirname(databasePath), { recursive: true });
    this.database = new DatabaseSync(databasePath);
    this.database.exec(`
      PRAGMA journal_mode = WAL;
      CREATE TABLE IF NOT EXISTS workbooks (
        unit_id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        snapshot_json TEXT NOT NULL,
        snapshot_revision INTEGER NOT NULL DEFAULT 0,
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
      CREATE UNIQUE INDEX IF NOT EXISTS changesets_unit_revision
        ON changesets(unit_id, revision);
      CREATE TABLE IF NOT EXISTS data_tables (
        table_id TEXT PRIMARY KEY,
        unit_id TEXT NOT NULL,
        name TEXT NOT NULL,
        fields_json TEXT NOT NULL,
        row_count INTEGER NOT NULL DEFAULT 0,
        block_size INTEGER NOT NULL DEFAULT 4096,
        revision INTEGER NOT NULL DEFAULT 0
      );
      CREATE TABLE IF NOT EXISTS data_blocks (
        block_id TEXT PRIMARY KEY,
        table_id TEXT NOT NULL,
        start_row INTEGER NOT NULL,
        row_count INTEGER NOT NULL,
        payload BLOB NOT NULL,
        UNIQUE(table_id, start_row)
      );
      CREATE TABLE IF NOT EXISTS snapshots (
        unit_id TEXT NOT NULL,
        revision INTEGER NOT NULL,
        checksum TEXT NOT NULL,
        payload BLOB NOT NULL,
        created_at TEXT NOT NULL,
        PRIMARY KEY (unit_id, revision)
      );
      CREATE TABLE IF NOT EXISTS workbook_acl (
        unit_id TEXT NOT NULL,
        subject TEXT NOT NULL,
        role TEXT NOT NULL CHECK (role IN ('owner', 'editor', 'commenter', 'viewer')),
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        PRIMARY KEY (unit_id, subject)
      );
      CREATE INDEX IF NOT EXISTS workbook_acl_subject
        ON workbook_acl(subject, unit_id);
      CREATE TABLE IF NOT EXISTS audit_events (
        audit_id TEXT PRIMARY KEY,
        unit_id TEXT NOT NULL,
        operation_id TEXT NOT NULL,
        actor_subject TEXT NOT NULL,
        action TEXT NOT NULL,
        target_revision INTEGER NOT NULL,
        revision INTEGER NOT NULL,
        reason TEXT,
        created_at TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS audit_events_unit_revision
        ON audit_events(unit_id, revision DESC);
    `);
    // Existing local databases predate snapshot_revision. This is a one-way
    // storage migration; a baseline snapshot is always revision 0 for them.
    try {
      this.database.exec('ALTER TABLE workbooks ADD COLUMN snapshot_revision INTEGER NOT NULL DEFAULT 0');
    } catch {
      // The column already exists.
    }
  }

  close(): void {
    this.database.close();
  }

  getAcl(unitId: string, actorSubject: string): WorkbookAclRecord[] {
    this.requireRole(unitId, actorSubject, 'owner');
    return this.listAclUnchecked(unitId);
  }

  listAcl(unitId: string, actorSubject: string): WorkbookAclRecord[] {
    return this.getAcl(unitId, actorSubject);
  }

  grantAccess(unitId: string, actorSubject: string, subject: string, role: WorkbookAclRole): WorkbookAclRecord {
    this.requireRole(unitId, actorSubject, 'owner');
    const normalizedSubject = requireSubject(subject);
    assertRole(role);
    const now = new Date().toISOString();
    this.database
      .prepare(
        `INSERT INTO workbook_acl (unit_id, subject, role, created_at, updated_at)
         VALUES (?, ?, ?, ?, ?)
         ON CONFLICT(unit_id, subject) DO UPDATE SET role = excluded.role, updated_at = excluded.updated_at`,
      )
      .run(unitId, normalizedSubject, role, now, now);
    const record = this.database
      .prepare('SELECT unit_id, subject, role, created_at, updated_at FROM workbook_acl WHERE unit_id = ? AND subject = ?')
      .get(unitId, normalizedSubject) as {
      unit_id?: string;
      subject?: string;
      role?: string;
      created_at?: string;
      updated_at?: string;
    } | undefined;
    if (!record?.unit_id || !record.subject || !record.role || !record.created_at || !record.updated_at) {
      throw new Error('ACL record was not persisted');
    }
    assertRole(record.role);
    return {
      unitId: record.unit_id,
      subject: record.subject,
      role: record.role,
      createdAt: record.created_at,
      updatedAt: record.updated_at,
    };
  }

  revokeAccess(unitId: string, actorSubject: string, subject: string): void {
    this.requireRole(unitId, actorSubject, 'owner');
    const normalizedSubject = requireSubject(subject);
    const target = this.database
      .prepare('SELECT role FROM workbook_acl WHERE unit_id = ? AND subject = ?')
      .get(unitId, normalizedSubject) as { role?: string } | undefined;
    if (target?.role === 'owner') throw new StorageAccessError('The workbook owner cannot be revoked');
    this.database.prepare('DELETE FROM workbook_acl WHERE unit_id = ? AND subject = ?').run(unitId, normalizedSubject);
  }

  getRole(unitId: string, actorSubject: string): WorkbookAclRole | undefined {
    const subject = requireSubject(actorSubject);
    const row = this.database
      .prepare('SELECT role FROM workbook_acl WHERE unit_id = ? AND subject = ?')
      .get(unitId, subject) as { role?: string } | undefined;
    if (!row?.role) return undefined;
    assertRole(row.role);
    return row.role;
  }

  listWorkbooks(actorSubject: string): WorkbookSummary[] {
    const subject = requireSubject(actorSubject);
    const rows = this.database
      .prepare(
        `SELECT w.unit_id, w.name, w.revision, w.updated_at
         FROM workbooks AS w
         INNER JOIN workbook_acl AS a ON a.unit_id = w.unit_id
         WHERE a.subject = ?
         ORDER BY w.updated_at DESC`,
      )
      .all(subject) as Array<{ unit_id: string; name: string; revision: number; updated_at: string }>;
    return rows.map((row) => ({
      unitId: row.unit_id,
      name: row.name,
      revision: row.revision,
      updatedAt: row.updated_at,
    }));
  }

  deleteWorkbook(unitId: string, actorSubject: string): void {
    this.requireRole(unitId, actorSubject, 'owner');
    this.database.exec('BEGIN IMMEDIATE');
    try {
      this.database.prepare('DELETE FROM changesets WHERE unit_id = ?').run(unitId);
      this.database.prepare('DELETE FROM snapshots WHERE unit_id = ?').run(unitId);
      this.database.prepare('DELETE FROM audit_events WHERE unit_id = ?').run(unitId);
      this.database.prepare('DELETE FROM workbook_acl WHERE unit_id = ?').run(unitId);
      const tables = this.database.prepare('SELECT table_id FROM data_tables WHERE unit_id = ?').all(unitId) as Array<{ table_id: string }>;
      for (const table of tables) this.database.prepare('DELETE FROM data_blocks WHERE table_id = ?').run(table.table_id);
      this.database.prepare('DELETE FROM data_tables WHERE unit_id = ?').run(unitId);
      this.database.prepare('DELETE FROM workbooks WHERE unit_id = ?').run(unitId);
      this.database.exec('COMMIT');
    } catch (error) {
      this.database.exec('ROLLBACK');
      throw error;
    }
  }

  createWorkbook(snapshot: WorkbookSnapshotV1, ownerSubject: string): SnapshotResponse {
    const subject = requireSubject(ownerSubject);
    const now = new Date().toISOString();
    this.database.exec('BEGIN IMMEDIATE');
    try {
      const result = this.database
        .prepare(
          `
        INSERT INTO workbooks (unit_id, name, snapshot_json, revision, updated_at)
        VALUES (?, ?, ?, 0, ?)
        ON CONFLICT(unit_id) DO NOTHING
      `,
        )
        .run(snapshot.unitId, snapshot.name, JSON.stringify(snapshot), now);
      if (Number(result.changes) === 0) {
        this.database.exec('ROLLBACK');
        this.requireRole(snapshot.unitId, subject, 'viewer');
        return this.getSnapshot(snapshot.unitId, subject);
      }
      this.database
        .prepare(
          `INSERT INTO workbook_acl (unit_id, subject, role, created_at, updated_at)
           VALUES (?, ?, 'owner', ?, ?)`,
        )
        .run(snapshot.unitId, subject, now, now);
      this.persistRevisionSnapshot(snapshot.unitId, 0, snapshot, now);
      this.database.exec('COMMIT');
      return { snapshot: structuredClone(snapshot), revision: 0 };
    } catch (error) {
      try {
        this.database.exec('ROLLBACK');
      } catch {
        // Preserve the original failure if the transaction was already closed.
      }
      throw error;
    }
  }

  getSnapshot(unitId: string, actorSubject: string): SnapshotResponse {
    this.requireRole(unitId, actorSubject, 'viewer');
    return this.getSnapshotUnchecked(unitId);
  }

  private getSnapshotUnchecked(unitId: string): SnapshotResponse {
    const row = this.database
      .prepare('SELECT snapshot_json, snapshot_revision, revision FROM workbooks WHERE unit_id = ?')
      .get(unitId) as { snapshot_json?: string; snapshot_revision?: number; revision?: number } | undefined;
    if (!row?.snapshot_json || row.revision == null) {
      throw new Error(`Workbook not found: ${unitId}`);
    }
    const baseRevision = row.snapshot_revision ?? 0;
    // The compressed revision snapshot is the integrity-checked source for a
    // baseline. The denormalized workbooks row is only a fast pointer and is
    // never rewritten as a side effect of a read.
    const stored = this.readStoredSnapshot(unitId, baseRevision);
    const snapshot = stored ?? migrateSnapshot(JSON.parse(row.snapshot_json) as WorkbookSnapshotV1);
    if (baseRevision === row.revision) {
      return { snapshot, revision: row.revision };
    }
    const workbook = WorkbookModel.fromSnapshot(snapshot);
    const runtime = createMutationRuntime(workbook);
    const changesets = this.database
      .prepare('SELECT payload_json FROM changesets WHERE unit_id = ? AND revision > ? AND revision <= ? ORDER BY revision ASC')
      .all(unitId, baseRevision, row.revision) as Array<{ payload_json: string }>;
    for (const entry of changesets) {
      const operation = JSON.parse(entry.payload_json) as CommittedOperationEnvelopeV2;
      runtime.applyRemoteMutations(operation.mutations.map((mutation) => ({
        id: mutation.id,
        unitId: operation.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: 'affectedRanges' in mutation ? mutation.affectedRanges : [],
      })));
    }
    return {
      snapshot: workbook.snapshot(),
      revision: row.revision,
    };
  }

  private listAclUnchecked(unitId: string): WorkbookAclRecord[] {
    const rows = this.database
      .prepare('SELECT unit_id, subject, role, created_at, updated_at FROM workbook_acl WHERE unit_id = ? ORDER BY subject')
      .all(unitId) as Array<{ unit_id: string; subject: string; role: string; created_at: string; updated_at: string }>;
    return rows.map((row) => {
      assertRole(row.role);
      return {
        unitId: row.unit_id,
        subject: row.subject,
        role: row.role,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
      };
    });
  }

  private requireRole(unitId: string, actorSubject: string, required: WorkbookAclRole): WorkbookAclRole {
    const subject = requireSubject(actorSubject);
    const row = this.database
      .prepare('SELECT role FROM workbook_acl WHERE unit_id = ? AND subject = ?')
      .get(unitId, subject) as { role?: string } | undefined;
    if (!row?.role) throw new StorageAccessError('Workbook access denied');
    assertRole(row.role);
    if (ROLE_RANK[row.role] < ROLE_RANK[required]) {
      throw new StorageAccessError(`Workbook role ${required} is required`);
    }
    return row.role;
  }

  saveSnapshot(unitId: string, snapshot: WorkbookSnapshotV1, baseRevision: number, actorSubject: string): SnapshotResponse {
    this.requireRole(unitId, actorSubject, 'editor');
    if (snapshot.unitId !== unitId) throw new Error('Snapshot unitId does not match route');
    const current = this.getSnapshotUnchecked(unitId);
    if (current.revision !== baseRevision) {
      throw new Error('Revision conflict');
    }
    const migrated = migrateSnapshot(snapshot);
    const nextRevision = current.revision + 1;
    const now = new Date().toISOString();
    this.database.exec('BEGIN IMMEDIATE');
    try {
      this.database
        .prepare('UPDATE workbooks SET snapshot_json = ?, snapshot_revision = ?, revision = ?, updated_at = ? WHERE unit_id = ?')
        .run(JSON.stringify(migrated), nextRevision, nextRevision, now, unitId);
      this.persistRevisionSnapshot(unitId, nextRevision, migrated, now);
      this.database.exec('COMMIT');
      return { snapshot: migrated, revision: nextRevision };
    } catch (error) {
      this.database.exec('ROLLBACK');
      throw error;
    }
  }

  /**
   * Authenticated V2 operation commit.  The caller supplies only mutation
   * intent; actor identity comes from the verified token and ranges are
   * conservatively derived from the server workbook.  A stale base revision
   * is rejected until the OT/rebase layer supplies a new envelope.
   */
  appendOperation(operationInput: OperationEnvelopeV2, actorSubject: string): { revision: number; operation: CommittedOperationEnvelopeV2 } {
    const operation = validateOperationEnvelopeV2(operationInput);
    if (operation.mutations.some((mutation) => mutation.id === 'workbook.restore')) {
      throw new StorageValidationError('workbook.restore is server-generated; submit targetRevision to the restore endpoint');
    }
    const actor = requireSubject(actorSubject);
    this.requireRole(operation.unitId, actor, 'editor');

    this.database.exec('BEGIN IMMEDIATE');
    try {
      const current = this.getSnapshotUnchecked(operation.unitId);
      const existing = this.database
        .prepare('SELECT revision, payload_json FROM changesets WHERE operation_id = ?')
        .get(operation.operationId) as { revision?: number; payload_json?: string } | undefined;
      if (existing?.revision != null && existing.payload_json) {
        const committed = JSON.parse(existing.payload_json) as CommittedOperationEnvelopeV2;
        if (committed.actorId !== actor) throw new StorageAccessError('Operation belongs to another subject');
        this.database.exec('COMMIT');
        return { revision: existing.revision, operation: committed };
      }
      if (current.revision !== operation.baseRevision) {
        throw new StorageConflictError(`Revision conflict: expected ${current.revision}, received ${operation.baseRevision}`);
      }

      const workbook = WorkbookModel.fromSnapshot(current.snapshot);
      const runtime = createMutationRuntime(workbook);
      const committedMutations = operation.mutations.map((mutation) => {
        const sheet = workbook.getSheet(mutation.sheetId);
        if (!sheet) throw new Error(`Sheet not found: ${mutation.sheetId}`);
        const affectedRanges = [{
          sheetId: mutation.sheetId,
          startRow: 0,
          endRow: Math.max(0, sheet.rowCount - 1),
          startColumn: 0,
          endColumn: Math.max(0, sheet.columnCount - 1),
        }];
        return {
          ...mutation,
          affectedRanges,
        };
      });
      runtime.applyRemoteMutations(committedMutations.map((mutation) => ({
        id: mutation.id,
        unitId: operation.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: mutation.affectedRanges,
      })));

      const nextRevision = current.revision + 1;
      const committedAt = new Date().toISOString();
      const committed: CommittedOperationEnvelopeV2 = {
        ...operation,
        actorId: actor,
        revision: nextRevision,
        committedAt,
        mutations: committedMutations,
      };
      const nextSnapshot = workbook.snapshot();
      this.database
        .prepare(
          `INSERT INTO changesets (operation_id, unit_id, revision, payload_json, created_at)
           VALUES (?, ?, ?, ?, ?)`,
        )
        .run(operation.operationId, operation.unitId, nextRevision, JSON.stringify(committed), committedAt);
      this.database
        .prepare('UPDATE workbooks SET revision = ?, updated_at = ? WHERE unit_id = ?')
        .run(nextRevision, committedAt, operation.unitId);
      if (this.shouldPersistSnapshot(operation.unitId, nextRevision)) {
        this.database
          .prepare('UPDATE workbooks SET snapshot_json = ?, snapshot_revision = ? WHERE unit_id = ?')
          .run(JSON.stringify(nextSnapshot), nextRevision, operation.unitId);
        this.persistRevisionSnapshot(operation.unitId, nextRevision, nextSnapshot, committedAt);
      }
      this.database.exec('COMMIT');
      return { revision: nextRevision, operation: committed };
    } catch (error) {
      try {
        this.database.exec('ROLLBACK');
      } catch {
        // Preserve the original error if SQLite already closed the transaction.
      }
      throw error;
    }
  }

  createDataTable(unitId: string, table: WorkbookTableModel, actorSubject: string): WorkbookTableModel {
    this.requireRole(unitId, actorSubject, 'editor');
    this.database.prepare(`
      INSERT INTO data_tables (table_id, unit_id, name, fields_json, row_count, block_size, revision)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(table.id, unitId, table.name, JSON.stringify(table.fields), table.rowCount, table.blockSize, table.revision);
    return structuredClone(table);
  }

  removeDataTable(unitId: string, tableId: string, actorSubject: string): void {
    this.requireRole(unitId, actorSubject, 'editor');
    const table = this.database.prepare('SELECT table_id FROM data_tables WHERE table_id = ? AND unit_id = ?').get(tableId, unitId) as { table_id?: string } | undefined;
    if (!table?.table_id) throw new Error('Data table not found');
    this.database.prepare('DELETE FROM data_blocks WHERE table_id = ?').run(tableId);
    this.database.prepare('DELETE FROM data_tables WHERE table_id = ? AND unit_id = ?').run(tableId, unitId);
  }

  appendDataBlock(unitId: string, tableId: string, startRow: number, rows: TableScalar[][], actorSubject: string): WorkbookTableBlock {
    this.requireRole(unitId, actorSubject, 'editor');
    const table = this.database.prepare('SELECT * FROM data_tables WHERE table_id = ? AND unit_id = ?').get(tableId, unitId) as { fields_json?: string; block_size?: number; revision?: number } | undefined;
    if (!table?.fields_json || table.block_size == null || table.revision == null) throw new Error('Data table not found');
    if (!Number.isInteger(startRow) || startRow < 0 || rows.length > table.block_size) throw new Error('Invalid data block');
    const fields = JSON.parse(table.fields_json) as Array<{ id: string }>;
    const columns: Record<string, TableScalar[]> = {};
    for (let column = 0; column < fields.length; column++) {
      columns[fields[column]!.id] = rows.map((row) => row[column] ?? null);
    }
    const encoded: EncodedTableBlock = { fields: fields.map((field) => field.id), columns };
    const blockId = `${tableId}:${startRow}`;
    this.database.prepare(`
      INSERT INTO data_blocks (block_id, table_id, start_row, row_count, payload)
      VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(table_id, start_row) DO UPDATE SET row_count = excluded.row_count, payload = excluded.payload
    `).run(blockId, tableId, startRow, rows.length, deflateRawSync(Buffer.from(JSON.stringify(encoded))));
    const nextRowCount = Math.max(Number((this.database.prepare('SELECT row_count FROM data_tables WHERE table_id = ?').get(tableId) as { row_count: number }).row_count), startRow + rows.length);
    const revision = table.revision + 1;
    this.database.prepare('UPDATE data_tables SET row_count = ?, revision = ? WHERE table_id = ? AND unit_id = ?').run(nextRowCount, revision, tableId, unitId);
    return { id: blockId, tableId, startRow, rowCount: rows.length, storageKey: blockId, encoding: 'typed-column-v1' };
  }

  readDataRows(unitId: string, tableId: string, offset: number, limit: number, actorSubject: string): { table: WorkbookTableModel; rows: TableScalar[][]; nextOffset?: number } {
    this.requireRole(unitId, actorSubject, 'viewer');
    const tableRow = this.database.prepare('SELECT * FROM data_tables WHERE table_id = ? AND unit_id = ?').get(tableId, unitId) as { table_id?: string; name?: string; fields_json?: string; row_count?: number; block_size?: number; revision?: number } | undefined;
    if (!tableRow?.table_id || !tableRow.fields_json || tableRow.row_count == null || tableRow.block_size == null || tableRow.revision == null) throw new Error('Data table not found');
    const fields = JSON.parse(tableRow.fields_json) as WorkbookTableModel['fields'];
    const blockRows = this.database.prepare('SELECT start_row, row_count, payload FROM data_blocks WHERE table_id = ? AND start_row <= ? AND start_row + row_count > ? ORDER BY start_row').all(tableId, offset + limit, offset) as Array<{ start_row: number; row_count: number; payload: Buffer }>;
    const rows: TableScalar[][] = [];
    for (const block of blockRows) {
      const encoded = JSON.parse(inflateRawSync(block.payload).toString('utf8')) as EncodedTableBlock;
      for (let rowIndex = 0; rowIndex < block.row_count; rowIndex++) {
        const row = encoded.fields.map((field) => encoded.columns[field]?.[rowIndex] ?? null);
        rows.push(row);
      }
    }
    const sliced = rows.slice(Math.max(0, offset - (blockRows[0]?.start_row ?? offset)), Math.max(0, offset - (blockRows[0]?.start_row ?? offset)) + limit);
    return {
      table: { id: tableRow.table_id, name: tableRow.name ?? tableId, fields, rowCount: tableRow.row_count, blockSize: tableRow.block_size, blocks: [], revision: tableRow.revision },
      rows: sliced,
      nextOffset: offset + sliced.length < tableRow.row_count ? offset + sliced.length : undefined,
    };
  }

  listRevisions(
    unitId: string,
    actorSubject: string,
  ): Array<{
    operationId: string;
    revision: number;
    createdAt: string;
    payload: CommittedOperationEnvelopeV2;
  }> {
    this.requireRole(unitId, actorSubject, 'viewer');
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
    return rows.map((row) => {
      const payload = JSON.parse(row.payload_json) as CommittedOperationEnvelopeV2;
      if (payload.schema !== 'OperationEnvelopeV2' || !payload.actorId) {
        throw new Error(`Unsupported historical operation schema at revision ${row.revision}`);
      }
      return {
        operationId: row.operation_id,
        revision: row.revision,
        createdAt: row.created_at,
        payload,
      };
    });
  }

  getSnapshotAtRevision(unitId: string, targetRevision: number, actorSubject: string): SnapshotResponse {
    this.requireRole(unitId, actorSubject, 'viewer');
    const current = this.getSnapshotUnchecked(unitId);
    if (targetRevision < 0 || targetRevision > current.revision) {
      throw new Error(`Revision not found: ${targetRevision}`);
    }
    if (targetRevision === current.revision) return current;

    const exact = this.readStoredSnapshot(unitId, targetRevision);
    if (exact) {
      return { snapshot: exact, revision: targetRevision };
    }

    const baseRow = this.database
      .prepare(
        'SELECT revision, checksum, payload FROM snapshots WHERE unit_id = ? AND revision <= ? ORDER BY revision DESC LIMIT 1',
      )
      .get(unitId, targetRevision) as { revision?: number; checksum?: string; payload?: Buffer } | undefined;

    let baseSnapshot: WorkbookSnapshotV1;
    let fromRevision: number;
    if (baseRow?.revision != null && baseRow.checksum && baseRow.payload) {
      baseSnapshot = this.decodeStoredSnapshot(baseRow.revision, baseRow.checksum, baseRow.payload);
      fromRevision = baseRow.revision;
    } else if (targetRevision === 0) {
      throw new Error(`Snapshot baseline missing for revision 0: ${unitId}`);
    } else {
      return this.getSnapshotAtRevision(unitId, 0, actorSubject);
    }

    const workbook = WorkbookModel.fromSnapshot(baseSnapshot);
    const runtime = createMutationRuntime(workbook);
    const rows = this.database
      .prepare(
        'SELECT payload_json FROM changesets WHERE unit_id = ? AND revision > ? AND revision <= ? ORDER BY revision ASC',
      )
      .all(unitId, fromRevision, targetRevision) as Array<{ payload_json: string }>;

    for (const row of rows) {
      const changeSet = JSON.parse(row.payload_json) as CommittedOperationEnvelopeV2;
      if (changeSet.schema !== 'OperationEnvelopeV2' || !changeSet.actorId) {
        throw new Error(`Unsupported historical operation schema at revision ${targetRevision}`);
      }
      runtime.applyRemoteMutations(changeSet.mutations.map((mutation) => ({
        id: mutation.id,
        unitId: changeSet.unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: mutation.affectedRanges,
      })));
    }

    return {
      snapshot: workbook.snapshot(),
      revision: targetRevision,
    };
  }

  getRevision(
    unitId: string,
    revision: number,
    actorSubject: string,
  ):
    | {
        operationId: string;
        revision: number;
        createdAt: string;
        payload: CommittedOperationEnvelopeV2;
      }
    | undefined {
    this.requireRole(unitId, actorSubject, 'viewer');
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
    const payload = JSON.parse(row.payload_json) as CommittedOperationEnvelopeV2;
    if (payload.schema !== 'OperationEnvelopeV2' || !payload.actorId) {
      throw new Error(`Unsupported historical operation schema at revision ${revision}`);
    }
    return {
      operationId: row.operation_id,
      revision: row.revision,
      createdAt: row.created_at,
      payload,
    };
  }

  /**
   * Resolve a historical snapshot and commit a server-authored restore
   * operation. The request carries only targetRevision/reason; the snapshot
   * is loaded from the authoritative revision store while the write lock is
   * held, then recorded as one atomic changeset and audit event.
   */
  restoreWorkbook(
    unitId: string,
    targetRevision: number,
    reason: string | undefined,
    actorSubject: string,
  ): HistoryRestoreResponse {
    const actor = requireSubject(actorSubject);
    // Restore rewrites the complete workbook aggregate. Keep the server
    // policy aligned with the canonical history.restore capability: only the
    // persisted workbook owner may create this operation.
    this.requireRole(unitId, actor, 'owner');
    if (!Number.isSafeInteger(targetRevision) || targetRevision < 0) {
      throw new StorageValidationError('targetRevision must be a non-negative safe integer');
    }
    if (reason !== undefined && (typeof reason !== 'string' || reason.length > 1000)) {
      throw new StorageValidationError('reason must be a string with at most 1000 characters');
    }

    this.database.exec('BEGIN IMMEDIATE');
    try {
      const current = this.getSnapshotUnchecked(unitId);
      if (targetRevision > current.revision) {
        throw new Error(`Revision not found: ${targetRevision}`);
      }

      // The historical read is performed by the storage layer, never from a
      // client-provided snapshot. It is safe inside the transaction because
      // getSnapshotAtRevision only reads the same SQLite connection.
      const historical = this.getSnapshotAtRevision(unitId, targetRevision, actor);
      const historicalSnapshot = structuredClone(historical.snapshot);
      const workbook = WorkbookModel.fromSnapshot(current.snapshot);
      const runtime = createMutationRuntime(workbook);
      const operationId = randomUUID();
      const committedAt = new Date().toISOString();
      const operation: OperationEnvelopeV2 = {
        schema: 'OperationEnvelopeV2',
        operationId,
        unitId,
        // Server-authored operations still use the monotonic envelope field;
        // client sequences remain independent and are never trusted here.
        clientSequence: current.revision + 1,
        baseRevision: current.revision,
        createdAt: committedAt,
        mutations: [{
          id: 'workbook.restore',
          sheetId: historicalSnapshot.activeSheetId,
          params: {
            serverGenerated: true,
            targetRevision,
            ...(reason === undefined ? {} : { reason }),
            snapshot: historicalSnapshot,
          },
        }],
      };
      // The canonical workbook.restore registration intentionally declares an
      // empty exact range: the mutation replaces the workbook aggregate, not
      // one cell rectangle. Keep the envelope aligned with that registry so
      // remote/history replay cannot accept a forged range list.
      const affectedRanges: CommittedOperationEnvelopeV2['mutations'][number]['affectedRanges'] = [];
      const committedMutations = operation.mutations.map((mutation) => ({
        ...mutation,
        affectedRanges,
      }));
      runtime.applyRemoteMutations(committedMutations.map((mutation) => ({
        id: mutation.id,
        unitId,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: mutation.affectedRanges,
      })));

      const nextRevision = current.revision + 1;
      const committed: CommittedOperationEnvelopeV2 = {
        ...operation,
        actorId: actor,
        revision: nextRevision,
        committedAt,
        mutations: committedMutations,
      };
      const nextSnapshot = workbook.snapshot();
      this.database
        .prepare(
          `INSERT INTO changesets (operation_id, unit_id, revision, payload_json, created_at)
           VALUES (?, ?, ?, ?, ?)`,
        )
        .run(operationId, unitId, nextRevision, JSON.stringify(committed), committedAt);
      this.database
        .prepare('UPDATE workbooks SET snapshot_json = ?, snapshot_revision = ?, revision = ?, updated_at = ? WHERE unit_id = ?')
        .run(JSON.stringify(nextSnapshot), nextRevision, nextRevision, committedAt, unitId);
      // A restore is a full materialized state transition. Persisting its
      // target snapshot immediately keeps future history preview/replay
      // bounded even when the normal 50 changeset threshold is not reached.
      this.persistRevisionSnapshot(unitId, nextRevision, nextSnapshot, committedAt);
      this.database
        .prepare(
          `INSERT INTO audit_events
             (audit_id, unit_id, operation_id, actor_subject, action, target_revision, revision, reason, created_at)
           VALUES (?, ?, ?, ?, 'workbook.restore', ?, ?, ?, ?)`,
        )
        .run(randomUUID(), unitId, operationId, actor, targetRevision, nextRevision, reason ?? null, committedAt);
      this.database.exec('COMMIT');
      return {
        snapshot: nextSnapshot,
        revision: nextRevision,
        targetRevision,
        operation: committed,
      };
    } catch (error) {
      try {
        this.database.exec('ROLLBACK');
      } catch {
        // Preserve the original failure if SQLite already closed the transaction.
      }
      throw error;
    }
  }

  listHistoryAudit(unitId: string, actorSubject: string): HistoryAuditRecord[] {
    this.requireRole(unitId, actorSubject, 'viewer');
    const rows = this.database
      .prepare(
        `SELECT audit_id, unit_id, operation_id, actor_subject, action,
                target_revision, revision, reason, created_at
           FROM audit_events
          WHERE unit_id = ?
          ORDER BY revision DESC, created_at DESC`,
      )
      .all(unitId) as Array<{
        audit_id: string;
        unit_id: string;
        operation_id: string;
        actor_subject: string;
        action: string;
        target_revision: number;
        revision: number;
        reason: string | null;
        created_at: string;
      }>;
    return rows.map((row) => {
      if (row.action !== 'workbook.restore') throw new Error(`Unsupported audit action: ${row.action}`);
      return {
        auditId: row.audit_id,
        unitId: row.unit_id,
        operationId: row.operation_id,
        actorId: row.actor_subject,
        action: 'workbook.restore',
        targetRevision: row.target_revision,
        revision: row.revision,
        ...(row.reason === null ? {} : { reason: row.reason }),
        createdAt: row.created_at,
      } satisfies HistoryAuditRecord;
    });
  }

  private persistRevisionSnapshot(
    unitId: string,
    revision: number,
    snapshot: WorkbookSnapshotV1,
    createdAt: string,
  ): void {
    const snapshotJson = JSON.stringify(snapshot);
    const checksum = computeSnapshotChecksum(snapshotJson);
    const payload = deflateRawSync(Buffer.from(snapshotJson, 'utf8'));
    this.database.prepare(`
      INSERT INTO snapshots (unit_id, revision, checksum, payload, created_at)
      VALUES (?, ?, ?, ?, ?)
      ON CONFLICT(unit_id, revision) DO UPDATE SET checksum = excluded.checksum, payload = excluded.payload, created_at = excluded.created_at
    `).run(unitId, revision, checksum, payload, createdAt);
  }

  private shouldPersistSnapshot(unitId: string, revision: number): boolean {
    const baseline = this.database
      .prepare('SELECT snapshot_revision FROM workbooks WHERE unit_id = ?')
      .get(unitId) as { snapshot_revision?: number } | undefined;
    const snapshotRevision = baseline?.snapshot_revision ?? 0;
    const pending = this.database
      .prepare('SELECT COUNT(*) AS count, COALESCE(SUM(LENGTH(payload_json)), 0) AS bytes FROM changesets WHERE unit_id = ? AND revision > ?')
      .get(unitId, snapshotRevision) as { count?: number; bytes?: number };
    if ((pending.count ?? 0) >= 50 || (pending.bytes ?? 0) >= 512_000) return true;
    const lastSnapshot = this.database
      .prepare('SELECT created_at FROM snapshots WHERE unit_id = ? AND revision <= ? ORDER BY revision DESC LIMIT 1')
      .get(unitId, revision) as { created_at?: string } | undefined;
    return !lastSnapshot?.created_at || Date.now() - Date.parse(lastSnapshot.created_at) >= 60_000;
  }

  private readStoredSnapshot(unitId: string, revision: number): WorkbookSnapshotV1 | undefined {
    const row = this.database
      .prepare('SELECT revision, checksum, payload FROM snapshots WHERE unit_id = ? AND revision = ?')
      .get(unitId, revision) as { revision?: number; checksum?: string; payload?: Buffer } | undefined;
    if (row?.revision == null || !row.checksum || !row.payload) return undefined;
    return this.decodeStoredSnapshot(row.revision, row.checksum, row.payload);
  }

  private decodeStoredSnapshot(revision: number, checksum: string, payload: Buffer): WorkbookSnapshotV1 {
    const json = inflateRawSync(payload).toString('utf8');
    const actualChecksum = computeSnapshotChecksum(json);
    if (actualChecksum !== checksum) {
      throw new Error(`Snapshot checksum mismatch at revision ${revision}`);
    }
    return migrateSnapshot(JSON.parse(json) as WorkbookSnapshotV1);
  }
}

export { PersistenceSession, computeSnapshotChecksum, verifySnapshotChecksum } from './persistence-session';
export type { PersistenceSnapshot, PersistenceSessionOptions } from './persistence-session';
