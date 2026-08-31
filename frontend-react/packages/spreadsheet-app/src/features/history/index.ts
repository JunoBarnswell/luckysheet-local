import { pivotSourceIdentity, WorkbookModel, type PivotResultTree, type WorkbookSnapshot } from '@react-sheets/core-model';
import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import { FormulaEngine, type CalculationSessionPort, type SheetTableRef } from '@react-sheets/formula-engine';
import { preparePivotTaskDescriptor, preparePivotTaskInputAsync } from '../pivot/engine';
import { InlinePivotTaskPort, type PivotTaskPort } from '../pivot/task-port';
import { createPivotCalculateRequest, createPivotSourceRegisterRequest, createPivotSourceReleaseRequest, type PivotTaskError } from '../pivot/task-protocol';
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
  readonly pivotErrors: ReadonlyMap<string, PivotTaskError>;
  private readonly projection: readonly CanvasSheetSnapshot[];
  private disposed = false;

  private constructor(
    workbook: WorkbookModel,
    formula: FormulaEngine,
    meta: HistoryEntryMeta,
    derivedCache: ReadonlyMap<string, PivotResultTree>,
    pivotErrors: ReadonlyMap<string, PivotTaskError>,
    projection: readonly CanvasSheetSnapshot[],
  ) {
    this.workbook = workbook;
    this.formula = formula;
    this.revision = meta.revision;
    this.meta = meta;
    this.derivedCache = derivedCache;
    this.pivotErrors = pivotErrors;
    this.projection = projection;
  }

  static async fromSnapshot(meta: HistoryEntryMeta, snapshot: WorkbookSnapshot, taskPort?: PivotTaskPort, calculationSessionPort?: CalculationSessionPort): Promise<HistoryPreviewSession> {
    const workbook = WorkbookModel.fromSnapshot(snapshot);
    const formula = await hydratePreviewFormula(workbook, calculationSessionPort);
    const derivedCache = new Map<string, PivotResultTree>();
    const pivotResults: Record<string, PivotResultTree> = {};
    const pivotErrors: Record<string, PivotTaskError> = {};
    const activePort = taskPort ?? new InlinePivotTaskPort();
    const registered = new Map<string, string>();
    let generation = 0;
    try {
      for (const sheet of workbook.getSheets()) for (const pivot of sheet.pivots) {
        const sourceIdentity = `history:${meta.revision}:${workbook.unitId}:${pivotSourceIdentity(pivot.source)}`;
        generation += 1;
        try {
          let descriptor = preparePivotTaskDescriptor(workbook, pivot, formula);
          if (registered.get(sourceIdentity) !== descriptor.revisions.sourceRevision) {
            const prepared = await preparePivotTaskInputAsync(workbook, pivot, formula);
            descriptor = { definition: prepared.definition, controls: prepared.controls, revisions: prepared.revisions, targetBounds: prepared.targetBounds };
            const registration = await activePort.submit(createPivotSourceRegisterRequest(`history-source:${generation}`, generation, sourceIdentity, prepared.revisions.sourceRevision, prepared.source));
            if (registration.status !== 'accepted') {
              if (registration.status === 'failed') pivotErrors[pivot.id] = registration.error;
              continue;
            }
            registered.set(sourceIdentity, prepared.revisions.sourceRevision);
          }
          const task = await activePort.submit(createPivotCalculateRequest(`history-calculate:${generation}`, generation, sourceIdentity, descriptor.definition, descriptor.controls, descriptor.revisions, descriptor.targetBounds));
          if (task.status !== 'completed') {
            if (task.status === 'failed') pivotErrors[pivot.id] = task.error;
            continue;
          }
          const cacheKey = pivotCacheKey(meta.revision, pivot.id);
          derivedCache.set(cacheKey, structuredClone(task.result));
          pivotResults[pivot.id] = task.result;
        } catch (error) {
          pivotErrors[pivot.id] = {
            code: 'PIVOT_TASK_FAILED',
            message: error instanceof Error ? error.message : `Historical Pivot failed: ${pivot.id}`,
            pivotId: pivot.id,
            sourceIdentity,
            sourceRevision: 'unknown',
            recovery: 'retry',
          };
        }
      }
    } finally {
      for (const [sourceIdentity, sourceRevision] of registered) {
        generation += 1;
        await activePort.submit(createPivotSourceReleaseRequest(`history-release:${generation}`, generation, sourceIdentity, sourceRevision));
      }
      if (!taskPort) activePort.dispose();
    }
    const projection = buildAllSheetSnapshots(workbook, formula, pivotResults, new Map(), pivotErrors);
    return new HistoryPreviewSession(workbook, formula, meta, derivedCache, new Map(Object.entries(pivotErrors)), projection);
  }

  get ui(): HistoryPreviewProjection {
    if (this.disposed) throw new Error('History preview session has been disposed');
    return {
      revision: this.revision,
      activeSheetId: this.workbook.primarySheetId,
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
    this.formula.disposeCalculationTasks();
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
  target.dataModel.tables.clear();
  target.dataModel.sources.clear();
  target.dataModel.relationships.clear();
  target.dataModel.views.clear();
  target.definedNameModels.splice(0, target.definedNameModels.length, ...structuredClone(restored.definedNameModels));
  target.name = restored.name;
  target.sheetOrder = [...restored.sheetOrder];
  // `definedNameModels` is the canonical store; the workbook-scoped formula
  // map is a derived read-only projection and must never be assigned.
  for (const [id, sheet] of restored.sheets) target.sheets.set(id, sheet);
  for (const [id, table] of restored.dataModel.tables) target.dataModel.tables.set(id, table);
  for (const [id, source] of restored.dataModel.sources) target.dataModel.sources.set(id, source);
  for (const [id, relationship] of restored.dataModel.relationships) target.dataModel.relationships.set(id, relationship);
  for (const [id, view] of restored.dataModel.views) target.dataModel.views.set(id, view);
}

/**
 * Register the server-authoritative restore mutation and the client request
 * command. A client request intentionally does not mutate the workbook: the
 * server resolves targetRevision, authorizes it, and broadcasts the signed
 * `workbook.restore` mutation carrying the materialized historical snapshot.
 */
export function registerHistoryCommands(registry: CommandRegistry): void {
  registry.registerMutation<ServerRestoreMutationParams>({
    id: 'workbook.restore',
    handler: (item, context) => {
      if (!isServerRestoreMutationParams(item.params)) {
        throw new Error('workbook.restore must be a server-generated targetRevision mutation');
      }
      applyRestoredWorkbook(context.workbook, item.params.snapshot);
    },
    metadata: {
      schema: { name: 'ServerRestoreMutationParams', validate: isServerRestoreMutationParams },
      permission: { capability: 'history.restore' },
      affectedRanges: { resolve: () => [], mode: 'exact' },
      inverseIds: ['workbook.restore'],
    },
  });

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

async function hydratePreviewFormula(workbook: WorkbookModel, calculationSessionPort?: CalculationSessionPort): Promise<FormulaEngine> {
  if (!calculationSessionPort && typeof window === 'undefined') {
    throw new Error('CALCULATION_SESSION_PORT_REQUIRED: history preview on Node requires an explicit calculation session port');
  }
  const engine = new FormulaEngine({ defaultSheetId: workbook.primarySheetId, calculationSessionPort });
  engine.setRecalculationMode('manual');
  engine.setDefinedNameModels(workbook.definedNameModels);
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
  await engine.recalculateAsync();
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
