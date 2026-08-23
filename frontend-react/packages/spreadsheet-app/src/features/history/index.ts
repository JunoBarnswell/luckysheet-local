import { WorkbookModel, type PivotResultTree, type WorkbookSnapshot } from '@react-sheets/core-model';
import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import { FormulaEngine, type SheetTableRef } from '@react-sheets/formula-engine';
import { computePivotResult } from '../pivot/engine';
import { buildAllSheetSnapshots, type CanvasSheetSnapshot } from '../../ui-snapshot';

export interface HistoryEntryMeta {
  revision: number;
  operationId: string;
  actorId?: string;
  category?: string;
  description?: string;
  createdAt: string;
}

export interface HistoryPreviewProjection {
  readonly revision: number;
  readonly activeSheetId: string;
  readonly sheets: readonly CanvasSheetSnapshot[];
}

/**
 * A history preview owns its workbook, formula engine and derived projections.
 * It never reuses the active runtime objects, and it exposes no command
 * runtime, so rendering a preview cannot create mutations or history entries.
 */
export class HistoryPreviewSession {
  readonly workbook: WorkbookModel;
  readonly formula: FormulaEngine;
  readonly revision: number;
  readonly meta: HistoryEntryMeta;
  readonly derivedCache: ReadonlyMap<string, PivotResultTree>;
  private readonly projection: readonly CanvasSheetSnapshot[];
  private disposed = false;

  private constructor(
    workbook: WorkbookModel,
    formula: FormulaEngine,
    meta: HistoryEntryMeta,
    derivedCache: ReadonlyMap<string, PivotResultTree>,
    projection: readonly CanvasSheetSnapshot[],
  ) {
    this.workbook = workbook;
    this.formula = formula;
    this.revision = meta.revision;
    this.meta = meta;
    this.derivedCache = derivedCache;
    this.projection = projection;
  }

  static fromSnapshot(meta: HistoryEntryMeta, snapshot: WorkbookSnapshot): HistoryPreviewSession {
    const workbook = WorkbookModel.fromSnapshot(snapshot);
    const formula = hydratePreviewFormula(workbook);
    const derivedCache = new Map<string, PivotResultTree>();
    const pivotResults: Record<string, PivotResultTree> = {};
    for (const sheet of workbook.getSheets()) {
      for (const pivot of sheet.pivots) {
        try {
          const result = computePivotResult(workbook, pivot);
          const cacheKey = pivotCacheKey(meta.revision, pivot.id);
          derivedCache.set(cacheKey, structuredClone(result));
          pivotResults[pivot.id] = result;
        } catch {
          // A corrupt historical pivot must not prevent the rest of the
          // workbook from being previewed.
        }
      }
    }
    const projection = buildAllSheetSnapshots(workbook, formula, pivotResults);
    return new HistoryPreviewSession(workbook, formula, meta, derivedCache, projection);
  }

  get ui(): HistoryPreviewProjection {
    if (this.disposed) throw new Error('History preview session has been disposed');
    return {
      revision: this.revision,
      activeSheetId: this.workbook.activeSheetId,
      sheets: this.projection,
    };
  }

  get sheets(): readonly CanvasSheetSnapshot[] {
    return this.ui.sheets;
  }

  getSheet(sheetId: string): CanvasSheetSnapshot | undefined {
    return this.ui.sheets.find((sheet) => sheet.id === sheetId);
  }

  dispose(): void {
    this.disposed = true;
  }
}

export interface RestoreCommandParams {
  targetRevision: number;
  reason?: string;
}

/** Server-produced mutation payload. The client command never accepts this shape. */
export interface ServerRestoreMutationParams extends RestoreCommandParams {
  serverGenerated: true;
  snapshot: WorkbookSnapshot;
}

function isServerRestoreMutationParams(value: unknown): value is ServerRestoreMutationParams {
  if (!value || typeof value !== 'object') return false;
  const input = value as Record<string, unknown>;
  return input.serverGenerated === true
    && Number.isSafeInteger(input.targetRevision)
    && Number(input.targetRevision) >= 0
    && Boolean(input.snapshot)
    && (input.snapshot as { schema?: string }).schema === 'WorkbookSnapshot';
}

function applyRestoredWorkbook(target: WorkbookModel, snapshot: WorkbookSnapshot): void {
  const restored = WorkbookModel.fromSnapshot(snapshot);
  target.sheets.clear();
  target.tables.clear();
  target.definedNameModels.splice(0, target.definedNameModels.length, ...structuredClone(restored.definedNameModels));
  target.name = restored.name;
  target.activeSheetId = restored.activeSheetId;
  target.sheetOrder = [...restored.sheetOrder];
  target.definedNames = { ...restored.definedNames };
  for (const [id, sheet] of restored.sheets) target.sheets.set(id, sheet);
  for (const [id, table] of restored.tables) target.tables.set(id, table);
}

/**
 * Register the server-authoritative restore mutation and the client request
 * command. A client request intentionally does not mutate the workbook: the
 * server resolves targetRevision, authorizes it, and broadcasts the signed
 * `workbook.restore` mutation carrying the materialized historical snapshot.
 */
export function registerHistoryCommands(registry: CommandRegistry): void {
  registry.registerMutation<ServerRestoreMutationParams>(
    'workbook.restore',
    (item, context) => {
      if (!isServerRestoreMutationParams(item.params)) {
        throw new Error('workbook.restore must be a server-generated targetRevision mutation');
      }
      applyRestoredWorkbook(context.workbook, item.params.snapshot);
    },
    {
      schema: { name: 'ServerRestoreMutationParams', validate: isServerRestoreMutationParams },
      permission: { capability: 'history.restore' },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['workbook.restore'],
    },
  );

  registry.registerCommand<RestoreCommandParams>({
    id: 'history.restore',
    execute(params: RestoreCommandParams, _context): CommandResult {
      if (!Number.isSafeInteger(params?.targetRevision) || params.targetRevision < 0) {
        throw new Error('history.restore requires a non-negative targetRevision');
      }
      throw new Error('history.restore is server-authorized; submit the targetRevision request to the server');
    },
  });
}

function pivotCacheKey(revision: number, pivotId: string): string {
  return `pivot:${pivotId}:source:${revision}:layout:${revision}:filter:${revision}`;
}

function hydratePreviewFormula(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.activeSheetId });
  engine.setRecalculationMode('manual');
  engine.setDefinedNames(workbook.definedNames);
  const tableRefs: SheetTableRef[] = workbook.getSheets().flatMap((sheet) => sheet.sheetTables.map((table) => ({
    id: table.id,
    sheetId: table.sheetId,
    name: table.name,
    range: table.range,
    hasHeaderRow: table.hasHeaderRow,
    hasTotalRow: table.hasTotalRow,
    columns: table.columns.map((column) => ({ id: column.id, name: column.name })),
  })));
  engine.setSheetTables(tableRefs);
  for (const sheet of workbook.getSheets()) {
    engine.setSpillEnvironment(sheet.id, {
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      isOccupied: (row, column) => sheet.cells.get(row, column) !== undefined,
    });
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula !== undefined) engine.setFormula(address, cell.formula);
      else if (cell.value !== null) engine.setValue(address, cell.value as never);
    });
  }
  engine.setRecalculationMode('automatic');
  engine.recalculate();
  return engine;
}

export class HistoryPanelStore {
  private entries: HistoryEntryMeta[] = [];

  setEntries(entries: HistoryEntryMeta[]): void {
    this.entries = [...entries].sort((a, b) => b.revision - a.revision);
  }

  search(query: string): HistoryEntryMeta[] {
    const q = query.trim().toLowerCase();
    if (!q) return [...this.entries];
    return this.entries.filter((e) =>
      e.description?.toLowerCase().includes(q)
      || e.actorId?.toLowerCase().includes(q)
      || e.category?.toLowerCase().includes(q),
    );
  }

  getByRevision(revision: number): HistoryEntryMeta | undefined {
    return this.entries.find((e) => e.revision === revision);
  }
}

export * from './replay';
