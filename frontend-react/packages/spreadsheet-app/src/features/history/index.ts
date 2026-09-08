import { pivotSourceIdentity, WorkbookModel, type PivotResultTree } from '@react-sheets/core-model';
import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import { FormulaEngine } from '@react-sheets/formula-engine';
import type { KernelPagePayload, WorkbookManifest } from '@react-sheets/protocol';
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

  static async fromManifest(meta: HistoryEntryMeta, manifest: WorkbookManifest, pages: readonly KernelPagePayload[], taskPort?: PivotTaskPort): Promise<HistoryPreviewSession> {
    if (manifest.revision !== meta.revision) throw new Error(`HISTORY_REVISION_MISMATCH: manifest ${manifest.revision} does not match requested ${meta.revision}`);
    const workbook = WorkbookModel.fromManifest(structuredClone(manifest), structuredClone(pages));
    const formula = new FormulaEngine({ unitId: workbook.unitId, revision: () => workbook.revision, defaultSheetId: workbook.primarySheetId });
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
  }
}

export interface RestoreCommandParams {
  targetRevision: number;
  reason?: string;
}

/**
 * Register the client history request command. A client request intentionally
 * does not mutate the workbook: the server resolves targetRevision and
 * authorizes the history intent.
 */
export function registerHistoryCommands(registry: CommandRegistry): void {
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
