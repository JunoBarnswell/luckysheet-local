import { WorkbookModel, type CellData, type RangeRef } from '@react-sheets/core-model';
import { CommandRuntime, type HistoryEntry } from '@react-sheets/command-runtime';
import { FormulaEngine } from '@react-sheets/formula-engine';
import {
  WorkbookApiClient,
  type CollaborationMessage,
  type CollaborationMutation,
  type SnapshotResponse,
} from '@react-sheets/protocol';
import { CollabSocketClient } from '@react-sheets/protocol';
import { registerSpreadsheetFeatures } from './feature-registry';

export interface RuntimeHandlers {
  onSaveState?: (state: import('./types').SaveState) => void;
  onNotice?: (message: string) => void;
  onMutationsApplied?: () => void;
  onPhaseChange?: (phase: import('./types').AppPhase) => void;
  onActiveSheetChange?: (sheetId: string) => void;
  onRemoteRevisions?: (revisions: import('@react-sheets/protocol').RevisionRecord[]) => void;
  onCollabStatus?: (status: 'connecting' | 'open' | 'closed') => void;
  onPeersChange?: (peers: import('./types').PeerCursor[]) => void;
}

export interface SpreadsheetRuntime {
  api: WorkbookApiClient;
  formula: FormulaEngine;
  model: WorkbookModel;
  commands: CommandRuntime;
  remoteConnected: boolean;
  remoteRevision: number;
  pendingMutations: CollaborationMutation[];
  detachers: Array<() => void>;
  handlers: RuntimeHandlers;
  ownOperationIds: Set<string>;
  nextClientSequence: number;
  pivotResults: Record<string, import('@react-sheets/core-model').PivotResultTree>;
  collab: CollabSocketClient | null;
  collabDispose: (() => void) | null;
  bootstrapDispose: (() => void) | null;
}

const UNIT_ID_STORAGE_KEY = 'react-sheets:unitId';
const ACTOR_ID_STORAGE_KEY = 'react-sheets:actorId';

export const PEER_COLORS = ['#2563eb', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4'];

export function hashCode(input: string): number {
  let hash = 0;
  for (let i = 0; i < input.length; i++) {
    hash = ((hash << 5) - hash + input.charCodeAt(i)) | 0;
  }
  return hash;
}

export function resolveUnitId(): string {
  if (typeof window === 'undefined') return 'wb-server-default';
  const routeMatch = /^\/workbooks\/([^/]+)\/?$/.exec(window.location.pathname);
  if (routeMatch?.[1]) return decodeURIComponent(routeMatch[1]);
  const existing = window.localStorage.getItem(UNIT_ID_STORAGE_KEY);
  if (existing) return existing;
  const generated = typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID()
    : 'wb-' + Date.now().toString(36);
  window.localStorage.setItem(UNIT_ID_STORAGE_KEY, generated);
  return generated;
}

export function resolveActorId(): string {
  if (typeof window === 'undefined') return 'actor-server';
  const existing = window.localStorage.getItem(ACTOR_ID_STORAGE_KEY);
  if (existing) return existing;
  const generated = typeof crypto !== 'undefined' && crypto.randomUUID
    ? crypto.randomUUID().slice(0, 8)
    : 'actor-' + Date.now().toString(36);
  window.localStorage.setItem(ACTOR_ID_STORAGE_KEY, generated);
  return generated;
}

export function createSpreadsheetRuntime(): SpreadsheetRuntime {
  const model = new WorkbookModel(resolveUnitId(), 'Untitled workbook');
  const commands = new CommandRuntime(model);
  registerSpreadsheetFeatures(commands);
  const runtime: SpreadsheetRuntime = {
    api: new WorkbookApiClient(),
    formula: new FormulaEngine({ defaultSheetId: 'sheet-1' }),
    model,
    commands,
    remoteConnected: false,
    remoteRevision: 0,
    pendingMutations: [],
    detachers: [],
    handlers: {},
    ownOperationIds: new Set(),
    nextClientSequence: 0,
    pivotResults: {},
    collab: null,
    collabDispose: null,
    bootstrapDispose: null,
  };
  attachCoreListeners(runtime);
  return runtime;
}

function syncEngineCell(
  engine: FormulaEngine,
  sheetId: string,
  row: number,
  column: number,
  data: CellData | undefined,
): void {
  const address = { sheetId, row, column };
  const hasContent = data !== undefined && (data.formula !== undefined || data.value != null);
  if (!hasContent) {
    engine.clearCell(address);
    return;
  }
  if (data!.formula) {
    engine.setFormula(address, data!.formula);
  } else {
    engine.setValue(address, data!.value as never);
  }
}

export function attachCoreListeners(runtime: SpreadsheetRuntime): void {
  detachCoreListeners(runtime);

  runtime.detachers.push(
    runtime.commands.onMutation((mutation) => {
      const changedSheet = runtime.model.getSheets().find((sheet) => sheet.id === mutation.sheetId);
      for (const pivot of changedSheet?.pivots ?? []) {
        delete runtime.pivotResults[pivot.id];
        if (
          mutation.id === 'cell.set' ||
          mutation.id === 'cell.restore' ||
          mutation.id === 'range.set' ||
          mutation.id === 'range.clear' ||
          mutation.id === 'rows.inserted' ||
          mutation.id === 'rows.deleted' ||
          mutation.id === 'columns.inserted' ||
          mutation.id === 'columns.deleted'
        ) {
          pivot.fieldCatalog = undefined;
        }
      }
      switch (mutation.id) {
        case 'cell.set': {
          const params = mutation.params as { row: number; column: number; value: CellData };
          syncEngineCell(runtime.formula, mutation.sheetId, params.row, params.column, params.value);
          break;
        }
        case 'cell.restore': {
          const params = mutation.params as { row: number; column: number; previous?: CellData };
          syncEngineCell(runtime.formula, mutation.sheetId, params.row, params.column, params.previous);
          break;
        }
        case 'range.set': {
          const params = mutation.params as { startRow: number; startColumn: number; values: CellData[][] };
          params.values.forEach((rowValues, rowOffset) =>
            rowValues.forEach((value, columnOffset) => {
              syncEngineCell(
                runtime.formula,
                mutation.sheetId,
                params.startRow + rowOffset,
                params.startColumn + columnOffset,
                value,
              );
            }),
          );
          break;
        }
        case 'range.clear': {
          const params = mutation.params as { range: RangeRef; mode?: 'all' | 'contents' | 'formats' };
          if (params.mode === 'formats') break;
          for (let r = params.range.startRow; r <= params.range.endRow; r++) {
            for (let c = params.range.startColumn; c <= params.range.endColumn; c++) {
              runtime.formula.clearCell({ sheetId: mutation.sheetId, row: r, column: c });
            }
          }
          break;
        }
        case 'rows.inserted':
        case 'rows.deleted': {
          const params = mutation.params as { at: number; count: number };
          runtime.formula.remapStructure(mutation.sheetId, {
            axis: 'row',
            at: params.at,
            count: params.count,
            op: mutation.id === 'rows.inserted' ? 'insert' : 'delete',
          });
          break;
        }
        case 'columns.inserted':
        case 'columns.deleted': {
          const params = mutation.params as { at: number; count: number };
          runtime.formula.remapStructure(mutation.sheetId, {
            axis: 'column',
            at: params.at,
            count: params.count,
            op: mutation.id === 'columns.inserted' ? 'insert' : 'delete',
          });
          break;
        }
        case 'name.set':
        case 'name.remove':
          runtime.formula.setDefinedNames(runtime.model.definedNames);
          break;
        default:
          break;
      }
    }),
  );

  runtime.detachers.push(
    runtime.commands.onMutation((mutation, source) => {
      if (source !== 'command') return;
      runtime.pendingMutations.push({
        id: mutation.id,
        sheetId: mutation.sheetId,
        params: mutation.params,
        affectedRanges: mutation.affectedRanges,
      });
    }),
  );

  runtime.detachers.push(
    runtime.commands.onCommand((_commandId, _params, result) => {
      if (runtime.commands.activeDepth > 0) return;
      const batch = runtime.pendingMutations;
      runtime.pendingMutations = [];
      runtime.handlers.onMutationsApplied?.();
      if (batch.length === 0) return;
      submitChangeset(runtime, result.operationId, batch);
    }),
  );
}

function detachCoreListeners(runtime: SpreadsheetRuntime): void {
  for (const detach of runtime.detachers) detach();
  runtime.detachers = [];
  runtime.pendingMutations = [];
}

function submitChangeset(
  runtime: SpreadsheetRuntime,
  operationId: string,
  mutations: CollaborationMutation[],
): void {
  if (!runtime.collab) {
    runtime.handlers.onSaveState?.('offline');
    return;
  }
  runtime.ownOperationIds.add(operationId);
  const changeSet = {
    schema: 'CollaborationChangeSetV1' as const,
    operationId,
    unitId: runtime.model.unitId,
    actorId: resolveActorId(),
    clientSequence: ++runtime.nextClientSequence,
    baseRevision: runtime.remoteRevision,
    mutations,
    createdAt: new Date().toISOString(),
  };
  runtime.handlers.onSaveState?.('saving');
  if (!runtime.collab.send({ type: 'changeset.submit', payload: changeSet })) {
    runtime.handlers.onSaveState?.('syncing');
  }
}

function rebuildFormulaEngine(workbook: WorkbookModel): FormulaEngine {
  const engine = new FormulaEngine({ defaultSheetId: workbook.activeSheetId });
  engine.setDefinedNames(workbook.definedNames);
  for (const sheet of workbook.getSheets()) {
    sheet.cells.forEach((cell, row, column) => {
      const address = { sheetId: sheet.id, row, column };
      if (cell.formula) engine.setFormula(address, cell.formula);
      else if (cell.value != null) engine.setValue(address, cell.value as never);
    });
  }
  return engine;
}

export function hydrateRuntime(runtime: SpreadsheetRuntime, response: SnapshotResponse): void {
  const workbook = WorkbookModel.fromSnapshot(response.snapshot);
  detachCoreListeners(runtime);
  runtime.model = workbook;
  runtime.commands = new CommandRuntime(workbook);
  registerSpreadsheetFeatures(runtime.commands);
  runtime.formula = rebuildFormulaEngine(workbook);
  attachCoreListeners(runtime);
  runtime.remoteRevision = response.revision;
  runtime.pivotResults = {};
}

export function startCollaborationSession(
  runtime: SpreadsheetRuntime,
  actorId: string,
  getSelectionKey: () => string,
): () => void {
  if (typeof window === 'undefined') return () => undefined;

  const protocol = window.location.protocol === 'https:' ? 'wss' : 'ws';
  const client = new CollabSocketClient(protocol + '://' + window.location.host + '/api/v1/collab');
  runtime.collab = client;

  const applyRemote = (message: CollaborationMessage) => {
    if (message.type === 'snapshot.response') {
      const response = message.payload ?? (message.snapshot && message.revision != null ? { snapshot: message.snapshot, revision: message.revision } : undefined);
      if (!response || (message.unitId && message.unitId !== runtime.model.unitId)) return;
      hydrateRuntime(runtime, response);
      runtime.remoteRevision = response.revision;
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
      runtime.handlers.onMutationsApplied?.();
    } else if (message.type === 'revision.created') {
      if (message.payload.unitId !== runtime.model.unitId) return;
      if (runtime.ownOperationIds.has(message.payload.operationId)) return;
      runtime.commands.applyRemoteMutations(
        message.payload.mutations.map((mutation) => ({ ...mutation, unitId: runtime.model.unitId })),
      );
      runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
      runtime.handlers.onMutationsApplied?.();
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
    } else if (message.type === 'changeset.ack') {
      runtime.ownOperationIds.add(message.operationId);
      runtime.remoteRevision = Math.max(runtime.remoteRevision, message.revision);
      runtime.handlers.onSaveState?.('saved');
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => undefined);
    } else if (message.type === 'changeset.reject') {
      runtime.ownOperationIds.delete(message.operationId);
      runtime.handlers.onSaveState?.('conflict');
      runtime.handlers.onNotice?.(`Change rejected: ${message.error.message}`);
      void runtime.api.getSnapshot(runtime.model.unitId).then((snapshot) => {
        hydrateRuntime(runtime, snapshot);
        runtime.handlers.onMutationsApplied?.();
      }).catch(() => undefined);
    } else if (message.type === 'cursor.updated' || message.type === 'presence.updated') {
      if (!message.unitId || message.unitId !== runtime.model.unitId) return;
      if (message.type === 'presence.updated' && (message.state as { status?: string } | null)?.status === 'offline') {
        runtime.handlers.onPeersChange?.([]);
        return;
      }
      const cursorState = message.state as { row?: number; column?: number; name?: string; sheetId?: string } | null;
      const color = PEER_COLORS[Math.abs(hashCode(message.actorId)) % PEER_COLORS.length]!;
      runtime.handlers.onPeersChange?.([
        {
          actorId: message.actorId,
          name: cursorState?.name ?? message.actorId.slice(0, 6),
          color,
          sheetId: cursorState?.sheetId ?? runtime.model.activeSheetId,
          row: cursorState?.row ?? 0,
          column: cursorState?.column ?? 0,
        },
      ]);
    }
  };

  const detachMessage = client.onMessage(applyRemote);
  const detachStatus = client.onStatus((status: 'connecting' | 'open' | 'closed') => {
    runtime.handlers.onCollabStatus?.(status);
    runtime.remoteConnected = status !== 'closed';
    if (status === 'closed') runtime.handlers.onSaveState?.('offline');
    else if (status === 'connecting') runtime.handlers.onSaveState?.('syncing');
    if (status === 'open') client.send({ type: 'snapshot.request', unitId: runtime.model.unitId });
  });
  client.open();

  let lastBroadcast = '';
  const broadcastTimer = window.setInterval(() => {
    const key = getSelectionKey();
    if (key === lastBroadcast) return;
    lastBroadcast = key;
    const parts = key.split(':');
    client.send({
      type: 'cursor.updated',
      unitId: runtime.model.unitId,
      actorId,
      state: { row: Number(parts[1]), column: Number(parts[2]), sheetId: parts[0] },
    });
  }, 400);

  return () => {
    window.clearInterval(broadcastTimer);
    detachMessage();
    detachStatus();
    client.close();
    runtime.collab = null;
  };
}

export function startPersistenceSession(runtime: SpreadsheetRuntime): () => void {
  let active = true;
  void (async () => {
    try {
      const snapshotResponse = await runtime.api.getSnapshot(runtime.model.unitId);
      hydrateRuntime(runtime, snapshotResponse);
      void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => runtime.handlers.onRemoteRevisions?.([]));
      runtime.remoteConnected = true;
      if (!active) return;
      runtime.handlers.onSaveState?.('saved');
      runtime.handlers.onNotice?.('Workbook restored from server');
      runtime.handlers.onPhaseChange?.('ready');
      runtime.handlers.onActiveSheetChange?.(runtime.model.activeSheetId);
      runtime.handlers.onMutationsApplied?.();
    } catch {
      try {
        const created = await runtime.api.createWorkbook(runtime.model.snapshot());
        void runtime.api.listRevisions(runtime.model.unitId).then((revs) => runtime.handlers.onRemoteRevisions?.(revs)).catch(() => runtime.handlers.onRemoteRevisions?.([]));
        runtime.remoteConnected = true;
        runtime.remoteRevision = Math.max(runtime.remoteRevision, created.revision);
        if (!active) return;
        runtime.handlers.onSaveState?.('saved');
        runtime.handlers.onNotice?.('SQLite sync connected');
        runtime.handlers.onPhaseChange?.('ready');
        runtime.handlers.onMutationsApplied?.();
      } catch {
        if (!active) return;
        runtime.handlers.onSaveState?.('offline');
        runtime.handlers.onNotice?.('Running local in-memory engine');
        runtime.handlers.onPhaseChange?.('ready');
        runtime.handlers.onMutationsApplied?.();
      }
    }
  })();
  return () => {
    active = false;
  };
}

export type { HistoryEntry };
