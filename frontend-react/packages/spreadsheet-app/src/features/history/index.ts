import { WorkbookModel, type WorkbookSnapshotV1 } from '@react-sheets/core-model';
import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';

export interface HistoryEntryMeta {
  revision: number;
  operationId: string;
  actorId?: string;
  category?: string;
  description?: string;
  createdAt: string;
  snapshot?: WorkbookSnapshotV1;
}

/** 只读独立 WorkbookModel — 不污染当前 session */
export class HistoryPreviewSession {
  readonly workbook: WorkbookModel;
  readonly revision: number;
  readonly meta: HistoryEntryMeta;

  private constructor(workbook: WorkbookModel, meta: HistoryEntryMeta) {
    this.workbook = workbook;
    this.revision = meta.revision;
    this.meta = meta;
  }

  static fromSnapshot(meta: HistoryEntryMeta, snapshot: WorkbookSnapshotV1): HistoryPreviewSession {
    const workbook = WorkbookModel.fromSnapshot(snapshot);
    return new HistoryPreviewSession(workbook, meta);
  }

  dispose(): void {
    // headless session — GC 即可
  }
}

export interface RestoreCommandParams {
  targetRevision: number;
  snapshot: WorkbookSnapshotV1;
  reason?: string;
}

/** Restore = 基于历史 snapshot 发新 Restore Command，生成新 revision，审计链不断 */
export function registerHistoryCommands(registry: CommandRegistry): void {
  registry.registerMutation('workbook.restore', (item, context) => {
    const params = item.params as RestoreCommandParams;
    const restored = WorkbookModel.fromSnapshot(params.snapshot);
    context.workbook.sheets.clear();
    context.workbook.tables.clear();
    context.workbook.name = restored.name;
    context.workbook.activeSheetId = restored.activeSheetId;
    context.workbook.definedNames = { ...restored.definedNames };
    for (const [id, sheet] of restored.sheets) {
      context.workbook.sheets.set(id, sheet);
    }
    for (const [id, table] of restored.tables) {
      context.workbook.tables.set(id, table);
    }
  });

  registry.registerCommand({
    id: 'history.restore',
    execute(params: RestoreCommandParams, context): CommandResult {
      const previous = context.workbook.snapshot();
      context.applyMutation({
        id: 'workbook.restore',
        unitId: context.workbook.unitId,
        sheetId: context.workbook.activeSheetId,
        params,
        affectedRanges: [],
        inverse: [{
          id: 'workbook.restore',
          unitId: context.workbook.unitId,
          sheetId: context.workbook.activeSheetId,
          params: { targetRevision: -1, snapshot: previous },
          affectedRanges: [],
        }],
        apply() {
          const restored = WorkbookModel.fromSnapshot(params.snapshot);
          context.workbook.sheets.clear();
          context.workbook.tables.clear();
          context.workbook.name = restored.name;
          context.workbook.activeSheetId = restored.activeSheetId;
          context.workbook.definedNames = { ...restored.definedNames };
          for (const [id, sheet] of restored.sheets) {
            context.workbook.sheets.set(id, sheet);
          }
          for (const [id, table] of restored.tables) {
            context.workbook.tables.set(id, table);
          }
        },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
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
