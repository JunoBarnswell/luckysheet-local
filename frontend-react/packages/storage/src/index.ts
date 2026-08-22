import { mkdirSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import type { CollaborationChangeSet, SnapshotResponse } from '@react-sheets/protocol';
import { WorkbookModel, type CellData, type WorkbookSnapshotV1 } from '@react-sheets/core-model';

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
    if (!row?.snapshot_json || row.revision == null)
      throw new Error(`Workbook not found: ${unitId}`);
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

    for (const mutation of changeSet.mutations) {
      if (mutation.id === 'cell.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as { row: number; column: number; value: CellData };
        sheet.cells.set(params.row, params.column, params.value);
      } else if (mutation.id === 'cell.restore') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as { row: number; column: number; previous?: CellData };
        if (params.previous) sheet.cells.set(params.row, params.column, params.previous);
        else sheet.cells.delete(params.row, params.column);
      } else if (mutation.id === 'range.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as {
          startRow: number;
          startColumn: number;
          values: CellData[][];
        };
        params.values.forEach((row, rowOffset) =>
          row.forEach((value, columnOffset) => {
            sheet.cells.set(
              params.startRow + rowOffset,
              params.startColumn + columnOffset,
              value,
            );
          }),
        );
      } else if (mutation.id === 'style.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as {
          range: { startRow: number; endRow: number; startColumn: number; endColumn: number };
          style: Record<string, unknown> & { numberFormat?: string };
        };
        for (let r = params.range.startRow; r <= params.range.endRow; r++) {
          for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
            let cell = sheet.cells.get(r, c);
            if (!cell) {
              cell = { value: null };
              sheet.cells.set(r, c, cell);
            }
            cell.style = { ...(cell.style ?? {}), ...(params.style as Partial<NonNullable<CellData['style']>>) };
            if (params.style.numberFormat !== undefined) {
              cell.numberFormat = params.style.numberFormat;
            }
          }
        }
      } else if (mutation.id === 'style.clear') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as {
          range: { startRow: number; endRow: number; startColumn: number; endColumn: number };
        };
        for (let r = params.range.startRow; r <= params.range.endRow; r++) {
          for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
            const cell = sheet.cells.get(r, c);
            if (!cell) continue;
            cell.style = undefined;
            cell.styleId = undefined;
            cell.numberFormat = undefined;
          }
        }
      } else if (mutation.id === 'range.clear') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const params = mutation.params as {
          range: { startRow: number; endRow: number; startColumn: number; endColumn: number };
          mode?: 'all' | 'contents' | 'formats';
        };
        for (let r = params.range.startRow; r <= params.range.endRow; r++) {
          for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
            if (params.mode === 'formats') {
              const cell = sheet.cells.get(r, c);
              if (cell) {
                cell.style = undefined;
                cell.styleId = undefined;
                cell.numberFormat = undefined;
              }
            } else {
              sheet.cells.delete(r, c);
            }
          }
        }
      } else if (mutation.id === 'sheet.add') {
        const params = mutation.params as {
          id: string;
          name: string;
          rowCount?: number;
          columnCount?: number;
        };
        workbook.addSheet(params.id, params.name, params.rowCount, params.columnCount);
      } else if (mutation.id === 'sheet.remove') {
        workbook.removeSheet(mutation.sheetId);
      } else if (mutation.id === 'sheet.restore') {
    const restored = mutation.params as { sheet: import('@react-sheets/core-model').WorksheetModel };
    const workbookModel = workbook as unknown as { sheets: Map<string, import('@react-sheets/core-model').WorksheetModel> };
    if (!workbookModel.sheets.has(restored.sheet.id)) {
      workbookModel.sheets.set(restored.sheet.id, restored.sheet);
    }
  } else if (mutation.id === 'sheet.rename') {
        const params = mutation.params as { sheetId: string; name: string };
        workbook.getSheet(params.sheetId).name = params.name;
      } else if (mutation.id === 'merge.set') {
        const params = mutation.params as {
          sheetId: string;
          range: { startRow: number; endRow: number; startColumn: number; endColumn: number };
        };
        const sheet = workbook.getSheet(params.sheetId);
        sheet.merges.push({
          range: { sheetId: params.sheetId, ...params.range },
          anchor: { row: params.range.startRow, column: params.range.startColumn },
        });
      } else if (mutation.id === 'merge.remove') {
        const params = mutation.params as {
          sheetId: string;
          range: { startRow: number; startColumn: number };
        };
        const sheet = workbook.getSheet(params.sheetId);
        const idx = sheet.merges.findIndex(
          (m) =>
            m.range.startRow === params.range.startRow &&
            m.range.startColumn === params.range.startColumn,
        );
        if (idx >= 0) sheet.merges.splice(idx, 1);
      } else if (mutation.id === 'freeze.set') {
        const params = mutation.params as { sheetId: string; freeze: any };
        workbook.getSheet(params.sheetId).freeze = { ...params.freeze };
      } else if (mutation.id === 'row.resize') {
        const params = mutation.params as { sheetId: string; row: number; height: number };
        workbook.getSheet(params.sheetId).rowHeights[params.row] = params.height;
      } else if (mutation.id === 'column.resize') {
        const params = mutation.params as { sheetId: string; column: number; width: number };
        workbook.getSheet(params.sheetId).columnWidths[params.column] = params.width;
      } else if (mutation.id === 'chart.add') {
        workbook.getSheet(mutation.sheetId).charts.push(mutation.params as never);
      } else if (mutation.id === 'chart.remove') {
        const items = workbook.getSheet(mutation.sheetId).charts;
        const index = items.findIndex((item) => item.id === mutation.params);
        if (index >= 0) items.splice(index, 1);
      } else if (mutation.id === 'pivot.add') {
        workbook.getSheet(mutation.sheetId).pivots.push(mutation.params as never);
      } else if (mutation.id === 'pivot.remove') {
        const items = workbook.getSheet(mutation.sheetId).pivots;
        const index = items.findIndex((item) => item.id === mutation.params);
        if (index >= 0) items.splice(index, 1);
      } else if (mutation.id === 'shape.add') {
        workbook.getSheet(mutation.sheetId).shapes.push(mutation.params as never);
      } else if (mutation.id === 'shape.remove') {
        const items = workbook.getSheet(mutation.sheetId).shapes;
        const index = items.findIndex((item) => item.id === mutation.params);
        if (index >= 0) items.splice(index, 1);
      } else if (mutation.id === 'sparkline.add') {
        workbook.getSheet(mutation.sheetId).sparklines.push(mutation.params as never);
      } else if (mutation.id === 'sparkline.remove') {
        const items = workbook.getSheet(mutation.sheetId).sparklines;
        const index = items.findIndex((item) => item.id === mutation.params);
        if (index >= 0) items.splice(index, 1);
      } else if (mutation.id === 'cf.add') {
        const sheet = workbook.getSheet(mutation.sheetId);
        sheet.conditionalFormats.push(mutation.params as never);
      } else if (mutation.id === 'cf.remove') {
        const sheet = workbook.getSheet(mutation.sheetId);
        const index = sheet.conditionalFormats.findIndex((item) => item.id === mutation.params);
        if (index >= 0) sheet.conditionalFormats.splice(index, 1);
      } else if (mutation.id === 'cf.clear') {
        const sheet = workbook.getSheet(mutation.sheetId);
        sheet.conditionalFormats.length = 0;
      } else if (mutation.id === 'filter.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        sheet.filter = mutation.params as never;
      } else if (mutation.id === 'filter.remove') {
        const sheet = workbook.getSheet(mutation.sheetId);
        sheet.filter = undefined;
      } else if (mutation.id === 'rows.inserted') {
        const params = mutation.params as { at: number; count: number };
        workbook.getSheet(mutation.sheetId).insertRows(params.at, params.count);
      } else if (mutation.id === 'rows.deleted') {
        const params = mutation.params as { at: number; count: number };
        workbook.getSheet(mutation.sheetId).deleteRows(params.at, params.count);
      } else if (mutation.id === 'columns.inserted') {
        const params = mutation.params as { at: number; count: number };
        workbook.getSheet(mutation.sheetId).insertColumns(params.at, params.count);
      } else if (mutation.id === 'columns.deleted') {
        const params = mutation.params as { at: number; count: number };
        workbook.getSheet(mutation.sheetId).deleteColumns(params.at, params.count);
      } else if (mutation.id === 'banded.set') {
        const sheet = workbook.getSheet(mutation.sheetId);
        sheet.bandedRule = (mutation.params ?? undefined) as never;
      } else if (mutation.id === 'image.add') {
        workbook.getSheet(mutation.sheetId).images.push(mutation.params as never);
      } else if (mutation.id === 'image.remove') {
        const items = workbook.getSheet(mutation.sheetId).images;
        const index = items.findIndex((item) => item.id === mutation.params);
        if (index >= 0) items.splice(index, 1);
      } else if (mutation.id === 'image.update') {
        const params = mutation.params as { id: string; bounds: unknown };
        const image = workbook.getSheet(mutation.sheetId).images.find((item) => item.id === params.id);
        if (image) image.bounds = params.bounds as typeof image.bounds;
      } else if (mutation.id === 'dv.add') {
    const sheet = workbook.getSheet(mutation.sheetId);
    const rule = (mutation.params as { rule: unknown }).rule as never;
    const index = sheet.dataValidations.findIndex((item) => item.id === (rule as { id: string }).id);
    if (index >= 0) sheet.dataValidations[index] = structuredClone(rule);
    else sheet.dataValidations.push(structuredClone(rule));
  } else if (mutation.id === 'dv.remove') {
    const sheet = workbook.getSheet(mutation.sheetId);
    const index = sheet.dataValidations.findIndex((item) => item.id === mutation.params);
    if (index >= 0) sheet.dataValidations.splice(index, 1);
  } else if (mutation.id === 'chart.update' || mutation.id === 'pro.chart.move') {
    const sheet = workbook.getSheet(mutation.sheetId);
    const params = mutation.params as { id: string; bounds?: { x: number; y: number; width: number; height: number }; sourceRanges?: unknown };
    const chart = sheet.charts.find((item) => item.id === params.id);
    if (chart) {
      if (params.bounds) chart.bounds = { ...params.bounds };
      if (params.sourceRanges) chart.sourceRanges = structuredClone(params.sourceRanges as never);
    }
  } else if (mutation.id === 'shape.update' || mutation.id === 'pro.shape.move') {
    const sheet = workbook.getSheet(mutation.sheetId);
    const params = mutation.params as { id: string; bounds: { x: number; y: number; width: number; height: number } };
    const shape = sheet.shapes.find((item) => item.id === params.id);
    if (shape && params.bounds) shape.bounds = { ...params.bounds };
  } else if (mutation.id === 'pro.pivot.write') {
    const sheet = workbook.getSheet(mutation.sheetId);
    const params = mutation.params as {
      pivotId: string;
      targetStartRow: number;
      targetStartColumn: number;
      values: Array<Array<{ value: import('@react-sheets/core-model').CellValue; style?: import('@react-sheets/core-model').CellStyle }>>;
    };
    for (let r = 0; r < params.values.length; r++) {
      const rowValues = params.values[r]!;
      for (let c = 0; c < rowValues.length; c++) {
        sheet.cells.set(params.targetStartRow + r, params.targetStartColumn + c, structuredClone(rowValues[c]!));
      }
    }
    const pivot = sheet.pivots.find((item) => item.id === params.pivotId);
    if (pivot) pivot.lastWrittenAt = Date.now();
  } else if (mutation.id === 'rows.hidden') {
    workbook.getSheet(mutation.sheetId).hiddenRows.add((mutation.params as { index: number }).index);
  } else if (mutation.id === 'rows.unhidden.all') {
    workbook.getSheet(mutation.sheetId).hiddenRows.clear();
  } else if (mutation.id === 'columns.hidden') {
    workbook.getSheet(mutation.sheetId).hiddenColumns.add((mutation.params as { index: number }).index);
  } else if (mutation.id === 'columns.unhidden.all') {
    workbook.getSheet(mutation.sheetId).hiddenColumns.clear();
  } else if (mutation.id === 'name.set') {
        const params = mutation.params as { name: string; reference: string };
        workbook.definedNames[params.name] = params.reference;
      } else if (mutation.id === 'name.remove') {
        delete workbook.definedNames[mutation.params as string];
      }
    }

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
    if (!row?.operation_id || row.revision == null || !row.created_at || !row.payload_json)
      return undefined;
    return {
      operationId: row.operation_id,
      revision: row.revision,
      createdAt: row.created_at,
      payload: JSON.parse(row.payload_json) as CollaborationChangeSet,
    };
  }
}
