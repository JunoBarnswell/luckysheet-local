import { WorkbookModel, type PivotResultTree } from '@react-sheets/core-model';
import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import { FormulaEngine } from '@react-sheets/formula-engine';
import type { KernelPagePayload, WorkbookManifest } from '@react-sheets/protocol';
import type { PivotTaskError } from '../pivot/server-task-port';
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

  static async fromManifest(meta: HistoryEntryMeta, manifest: WorkbookManifest, pages: readonly KernelPagePayload[]): Promise<HistoryPreviewSession> {
    if (manifest.revision !== meta.revision) throw new Error(`HISTORY_REVISION_MISMATCH: manifest ${manifest.revision} does not match requested ${meta.revision}`);
    const workbook = WorkbookModel.fromManifest(structuredClone(manifest), structuredClone(pages));
    const formula = new FormulaEngine({ unitId: workbook.unitId, revision: () => workbook.revision, defaultSheetId: workbook.primarySheetId });
    const derivedCache = new Map<string, PivotResultTree>();
    const pivotResults: Record<string, PivotResultTree> = {};
    const pivotErrors: Record<string, PivotTaskError> = {};
    for (const sheet of workbook.getSheets()) for (const pivot of sheet.pivots) {
      pivotErrors[pivot.id] = {
        code: 'PIVOT_SOURCE_UNAVAILABLE',
        message: 'Historical Pivot results are not embedded in this revision preview',
        pivotId: pivot.id,
        sourceIdentity: `${workbook.unitId}:${pivot.id}`,
        sourceRevision: String(meta.revision),
        recovery: 'retry',
      };
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
