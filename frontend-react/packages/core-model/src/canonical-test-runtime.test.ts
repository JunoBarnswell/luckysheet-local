import { initializeNodeKernel } from '@react-sheets/kernel-client/node';
import { kernelInvoke } from '@react-sheets/kernel-client';
import { WorkbookModel, type KernelReplicaManifest } from './index';
import { CommandRuntime, type KernelCommitRequest, type KernelCommittedOperation } from '@react-sheets/command-runtime';

let kernelReady: Promise<void> | undefined;

/** Open the same manifest/page backed model used by the browser host. */
export async function openCanonicalTestRuntime(unitId: string, name = 'Canonical test'): Promise<{
  workbook: WorkbookModel;
  runtime: CommandRuntime;
  close: () => void;
}> {
  kernelReady ??= initializeNodeKernel();
  await kernelReady;
  const manifest = kernelInvoke<KernelReplicaManifest>('create', {
    unitId,
    name,
    sheets: [{ sheetId: 'sheet-1', name: 'Sheet1', rowCount: 256, columnCount: 32, metadata: {} }],
  });
  const workbook = WorkbookModel.fromManifest(manifest);
  const runtime = new CommandRuntime(workbook);
  const history = new Map<string, { baseRevision: number; record: unknown }>();
  runtime.setCommitPort(async (request: KernelCommitRequest) => {
    const target = request.intent ? history.get(request.intent.targetOperationId) : undefined;
    if (request.intent && (!target || target.baseRevision !== request.intent.targetBaseRevision)) {
      throw new Error('UNDO_TARGET_NOT_FOUND');
    }
    const committed = kernelInvoke<KernelCommittedOperation & { history: unknown }>('command', {
      unitId,
      baseRevision: request.baseRevision,
      operationId: request.operationId,
      commandId: request.intent ? 'history.undo' : 'operation.apply',
      accessRole: 'owner',
      params: request.intent ? { history: target!.record } : {
        mutations: request.mutations.map(({ id, sheetId, params }) => ({ id, sheetId, params })),
      },
    });
    history.set(request.operationId, { baseRevision: request.baseRevision, record: committed.history });
    return committed;
  });
  let closed = false;
  return {
    workbook,
    runtime,
    close: () => {
      if (closed) return;
      closed = true;
      kernelInvoke('close', { unitId });
    },
  };
}

export async function seedCanonicalCells(
  runtime: CommandRuntime,
  sheetId: string,
  cells: ReadonlyArray<{ row: number; column: number; value: unknown; formula?: string }>,
): Promise<void> {
  for (const cell of cells) {
    await runtime.execute('sheet.cell.set', {
      sheetId,
      row: cell.row,
      column: cell.column,
      value: { value: cell.value, ...(cell.formula ? { formula: cell.formula } : {}) },
    });
  }
}
