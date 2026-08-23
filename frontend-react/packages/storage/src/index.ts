import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import { deflateRawSync, inflateRawSync } from 'node:zlib';
import { CommandRuntime } from '@react-sheets/command-runtime';
import { WorkbookModel, type PivotLayout, type TableScalar, type WorkbookTableBlock, type WorkbookTableModel, type WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { CollaborationChangeSet, CollaborationMutation, SnapshotResponse, WorkbookSummary } from '@react-sheets/protocol';
import { registerSheetCommands } from '@react-sheets/sheet-features';
import { registerProSheetCommands } from '@react-sheets/pro-features';
import { registerSpreadsheetFeatures } from '@react-sheets/spreadsheet-app';
import { PersistenceSession, computeSnapshotChecksum } from './persistence-session';

const databasePath = resolve(process.cwd(), 'data/react-sheets.sqlite');
mkdirSync(dirname(databasePath), { recursive: true });

function createMutationRuntime(workbook: WorkbookModel): CommandRuntime {
  const runtime = new CommandRuntime(workbook);
  registerSheetCommands(runtime);
  registerProSheetCommands(runtime);
  registerSpreadsheetFeatures(runtime);
  return runtime;
}

function rangesOverlap(left: CollaborationMutation['affectedRanges'][number], right: CollaborationMutation['affectedRanges'][number]): boolean {
  return left.sheetId === right.sheetId
    && left.startRow <= right.endRow
    && right.startRow <= left.endRow
    && left.startColumn <= right.endColumn
    && right.startColumn <= left.endColumn;
}

function changesConflict(incoming: CollaborationChangeSet, committed: CollaborationChangeSet): boolean {
  return incoming.mutations.some((left) => committed.mutations.some((right) => {
    if (left.affectedRanges.length === 0 || right.affectedRanges.length === 0) return true;
    return left.affectedRanges.some((leftRange) => right.affectedRanges.some((rightRange) => rangesOverlap(leftRange, rightRange)));
  }));
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
  private readonly database = new DatabaseSync(databasePath);
  readonly persistence: PersistenceSession;

  constructor() {
    this.persistence = new PersistenceSession({
      changesetThreshold: 50,
      byteThreshold: 512_000,
      timeThresholdMs: 60_000,
      appendChangeSet: async (changeSet) => {
        // changeset 已在 appendChangeSet 方法写入 DB
        void changeSet;
      },
      persistSnapshot: async (record) => {
        const json = JSON.stringify(record.snapshot);
        const payload = deflateRawSync(Buffer.from(json, 'utf8'));
        this.database.prepare(`
          INSERT INTO snapshots (unit_id, revision, checksum, payload, created_at)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(unit_id, revision) DO UPDATE SET checksum = excluded.checksum, payload = excluded.payload, created_at = excluded.created_at
        `).run(record.snapshot.unitId, record.revision, record.checksum, payload, record.createdAt);
      },
    });
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

  deleteWorkbook(unitId: string): void {
    this.database.prepare('DELETE FROM changesets WHERE unit_id = ?').run(unitId);
    const tables = this.database.prepare('SELECT table_id FROM data_tables WHERE unit_id = ?').all(unitId) as Array<{ table_id: string }>;
    for (const table of tables) this.database.prepare('DELETE FROM data_blocks WHERE table_id = ?').run(table.table_id);
    this.database.prepare('DELETE FROM data_tables WHERE unit_id = ?').run(unitId);
    this.database.prepare('DELETE FROM workbooks WHERE unit_id = ?').run(unitId);
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
    const snapshot = migrateSnapshot(JSON.parse(row.snapshot_json) as WorkbookSnapshotV1);
    const migratedJson = JSON.stringify(snapshot);
    if (migratedJson !== row.snapshot_json) {
      this.database.prepare('UPDATE workbooks SET snapshot_json = ? WHERE unit_id = ?').run(migratedJson, unitId);
    }
    return {
      snapshot,
      revision: row.revision,
    };
  }

  appendChangeSet(changeSet: CollaborationChangeSet): number {
    if (!changeSet.operationId || !changeSet.actorId || !Number.isSafeInteger(changeSet.clientSequence) || changeSet.clientSequence < 1) {
      throw new Error('Invalid collaboration sequence');
    }
    const current = this.getSnapshot(changeSet.unitId);
    const existing = this.database
      .prepare('SELECT revision FROM changesets WHERE operation_id = ?')
      .get(changeSet.operationId) as { revision?: number } | undefined;
    if (existing?.revision != null) return existing.revision;
    if (current.revision !== changeSet.baseRevision) {
      const committedRows = this.database
        .prepare('SELECT payload_json FROM changesets WHERE unit_id = ? AND revision > ? ORDER BY revision ASC')
        .all(changeSet.unitId, changeSet.baseRevision) as Array<{ payload_json: string }>;
      const hasConflict = committedRows.some((row) => changesConflict(changeSet, JSON.parse(row.payload_json) as CollaborationChangeSet));
      if (hasConflict) throw new Error('Revision conflict');
    }

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
    this.database.exec('BEGIN IMMEDIATE');
    try {
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
      this.database.exec('COMMIT');
      void this.persistence.recordChangeSet(changeSet);
      const snapshotJson = JSON.stringify(nextSnapshot);
      this.persistence.writeSnapshot(nextSnapshot, nextRevision).catch(() => {
        if (this.persistence.shouldSnapshot()) {
          const checksum = computeSnapshotChecksum(snapshotJson);
          const payload = deflateRawSync(Buffer.from(snapshotJson, 'utf8'));
          this.database.prepare(`
            INSERT INTO snapshots (unit_id, revision, checksum, payload, created_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(unit_id, revision) DO UPDATE SET checksum = excluded.checksum, payload = excluded.payload
          `).run(changeSet.unitId, nextRevision, checksum, payload, new Date().toISOString());
        }
      });
      return nextRevision;
    } catch (error) {
      this.database.exec('ROLLBACK');
      throw error;
    }
  }

  createDataTable(unitId: string, table: WorkbookTableModel): WorkbookTableModel {
    this.database.prepare(`
      INSERT INTO data_tables (table_id, unit_id, name, fields_json, row_count, block_size, revision)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(table.id, unitId, table.name, JSON.stringify(table.fields), table.rowCount, table.blockSize, table.revision);
    return structuredClone(table);
  }

  removeDataTable(unitId: string, tableId: string): void {
    const table = this.database.prepare('SELECT table_id FROM data_tables WHERE table_id = ? AND unit_id = ?').get(tableId, unitId) as { table_id?: string } | undefined;
    if (!table?.table_id) throw new Error('Data table not found');
    this.database.prepare('DELETE FROM data_blocks WHERE table_id = ?').run(tableId);
    this.database.prepare('DELETE FROM data_tables WHERE table_id = ? AND unit_id = ?').run(tableId, unitId);
  }

  appendDataBlock(unitId: string, tableId: string, startRow: number, rows: TableScalar[][]): WorkbookTableBlock {
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

  readDataRows(unitId: string, tableId: string, offset: number, limit: number): { table: WorkbookTableModel; rows: TableScalar[][]; nextOffset?: number } {
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

export { PersistenceSession, computeSnapshotChecksum, verifySnapshotChecksum } from './persistence-session';
export type { PersistenceSnapshot, PersistenceSessionOptions } from './persistence-session';
