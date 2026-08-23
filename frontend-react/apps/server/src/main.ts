import { createServer } from 'node:http';
import { randomUUID } from 'node:crypto';
import { cpus } from 'node:os';
import { WebSocketServer } from 'ws';
import { WorkbookModel } from '@react-sheets/core-model';
import {
  decodeClientOperationMessageV2,
  encodeOperationMessageV2,
  validateHistoryRestoreRequest,
  type ApiError,
  type OperationMessageV2,
  type WorkbookAclRole,
} from '@react-sheets/protocol';
import { StorageValidationError, WorkbookStorage } from '@react-sheets/storage';
import { computePivotResult } from '@react-sheets/pro-features';
import { importXlsx, exportXlsx, parseXlsxXmlToSnapshot } from '@react-sheets/exchange-xlsx';
import { AuthenticationError, JwtAuthenticator } from './auth';

const storage = new WorkbookStorage();
const authenticator = new JwtAuthenticator();
const port = Number(process.env.REACT_SHEETS_SERVER_PORT ?? 4181);
const clientsByUnit = new Map<string, Set<import('ws').WebSocket>>();

class CalculationQueue {
  private active = 0;
  private readonly pending: Array<{ task: () => Promise<unknown>; resolve: (value: unknown) => void; reject: (reason: unknown) => void }> = [];

  constructor(private readonly concurrency: number) {}

  run<T>(task: () => Promise<T>): Promise<T> {
    return new Promise<T>((resolve, reject) => {
      this.pending.push({ task, resolve: resolve as (value: unknown) => void, reject });
      this.drain();
    });
  }

  private drain(): void {
    while (this.active < this.concurrency && this.pending.length > 0) {
      const item = this.pending.shift()!;
      this.active += 1;
      void item.task().then(item.resolve, item.reject).finally(() => {
        this.active -= 1;
        this.drain();
      });
    }
  }
}

const calculationQueue = new CalculationQueue(Math.max(1, cpus().length - 1));

function joinUnit(unitId: string, socket: import('ws').WebSocket): Set<import('ws').WebSocket> {
  const clients = clientsByUnit.get(unitId) ?? new Set<import('ws').WebSocket>();
  clients.add(socket);
  clientsByUnit.set(unitId, clients);
  return clients;
}

function leaveUnit(unitId: string | null, socket: import('ws').WebSocket): void {
  if (!unitId) return;
  const clients = clientsByUnit.get(unitId);
  if (!clients) return;
  clients.delete(socket);
  if (clients.size === 0) clientsByUnit.delete(unitId);
}

function broadcastRevision(unitId: string, operation: import('@react-sheets/protocol').CommittedOperationEnvelopeV2): void {
  const clients = clientsByUnit.get(unitId);
  if (!clients) return;
  const message = encodeOperationMessageV2({
    type: 'revision.created',
    payload: operation,
    revision: operation.revision,
  });
  for (const client of clients) {
    if (client.readyState === client.OPEN) client.send(message);
  }
}

function createDefaultSnapshot(unitId: string) {
  return new WorkbookModel(unitId, 'React Sheets Workbook').snapshot();
}

async function readBody(request: import('node:http').IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of request) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString('utf8');
}

function sendJson(response: import('node:http').ServerResponse, status: number, body: unknown): void {
  response.writeHead(status, { 'content-type': 'application/json; charset=utf-8' });
  response.end(JSON.stringify(body));
}

function errorStatus(error: unknown): number {
  if (typeof error === 'object' && error !== null && 'status' in error && typeof (error as { status?: unknown }).status === 'number') {
    return (error as { status: number }).status;
  }
  if (error instanceof Error && /not found/i.test(error.message)) return 404;
  if (error instanceof Error && /conflict/i.test(error.message)) return 409;
  if (error instanceof Error && /required|invalid|unsupported|must be/i.test(error.message)) return 400;
  return 500;
}

function errorCode(error: unknown): string {
  if (typeof error === 'object' && error !== null && 'code' in error && typeof (error as { code?: unknown }).code === 'string') {
    return (error as { code: string }).code;
  }
  return errorStatus(error) >= 500 ? 'INTERNAL_ERROR' : 'REQUEST_FAILED';
}

function protocolErrorCode(error: unknown): ApiError['code'] {
  const code = errorCode(error);
  if (code === 'VALIDATION_ERROR' || code === 'NOT_FOUND' || code === 'CONFLICT'
    || code === 'UNAUTHENTICATED' || code === 'FORBIDDEN'
    || code === 'AUTH_CONFIGURATION_ERROR' || code === 'INTERNAL_ERROR') return code;
  return errorStatus(error) >= 500 ? 'INTERNAL_ERROR' : 'VALIDATION_ERROR';
}

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? '/', `http://${request.headers.host ?? '127.0.0.1'}`);

    if (request.method === 'GET' && url.pathname === '/health') {
      sendJson(response, 200, { status: 'ok', service: 'react-sheets-server' });
      return;
    }

    let principal;
    try {
      principal = await authenticator.authenticateRequest(request);
    } catch (error) {
      const status = error instanceof AuthenticationError ? 401 : errorStatus(error);
      sendJson(response, status, {
        code: errorCode(error),
        message: error instanceof Error ? error.message : 'Authentication failed',
      });
      return;
    }

    if (request.method === 'GET' && url.pathname === '/api/v1/workbooks') {
      sendJson(response, 200, storage.listWorkbooks(principal.subject));
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/acl$/.test(url.pathname)) {
      const unitId = decodeURIComponent(url.pathname.split('/')[4] ?? '');
      if (!unitId) throw new Error('Workbook id is required');
      sendJson(response, 200, { entries: storage.listAcl(unitId, principal.subject) });
      return;
    }

    if (request.method === 'PUT' && /^\/api\/v1\/workbooks\/[^/]+\/acl\/[^/]+$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = decodeURIComponent(parts[4] ?? '');
      const subject = decodeURIComponent(parts[6] ?? '');
      const body = JSON.parse(await readBody(request)) as { role?: string };
      if (!unitId || !subject || !body.role) {
        sendJson(response, 400, { code: 'VALIDATION_ERROR', message: 'ACL role is required' });
        return;
      }
      if (!['owner', 'editor', 'commenter', 'viewer'].includes(body.role)) {
        sendJson(response, 400, { code: 'VALIDATION_ERROR', message: 'Unsupported ACL role' });
        return;
      }
      sendJson(response, 200, storage.grantAccess(unitId, principal.subject, subject, body.role as WorkbookAclRole));
      return;
    }

    if (request.method === 'DELETE' && /^\/api\/v1\/workbooks\/[^/]+\/acl\/[^/]+$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = decodeURIComponent(parts[4] ?? '');
      const subject = decodeURIComponent(parts[6] ?? '');
      if (!unitId || !subject) throw new Error('Workbook and subject are required');
      storage.revokeAccess(unitId, principal.subject, subject);
      sendJson(response, 200, {});
      return;
    }

    if (request.method === 'DELETE' && /^\/api\/v1\/workbooks\/[^/]+$/.test(url.pathname)) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      storage.deleteWorkbook(unitId, principal.subject);
      sendJson(response, 200, {});
      return;
    }

    if (request.method === 'GET' && url.pathname.startsWith('/api/v1/workbooks/') && url.pathname.endsWith('/snapshot')) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      sendJson(response, 200, storage.getSnapshot(unitId, principal.subject));
      return;
    }

    if (request.method === 'PUT' && url.pathname.startsWith('/api/v1/workbooks/') && url.pathname.endsWith('/snapshot')) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      const body = JSON.parse(await readBody(request)) as {
        snapshot?: ReturnType<typeof createDefaultSnapshot>;
        baseRevision?: number;
      };
      if (!body.snapshot || body.baseRevision == null) {
        sendJson(response, 400, { code: 'VALIDATION_ERROR', message: 'snapshot and baseRevision are required' });
        return;
      }
      try {
        sendJson(response, 200, storage.saveSnapshot(unitId, body.snapshot, body.baseRevision, principal.subject));
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Save failed';
        const status = errorStatus(error);
        sendJson(response, status, { code: protocolErrorCode(error), message });
      }
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/revisions$/.test(url.pathname)) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      sendJson(response, 200, { revisions: storage.listRevisions(unitId, principal.subject) });
      return;
    }

    if (request.method === 'POST' && /^\/api\/v1\/workbooks\/[^/]+\/restore$/.test(url.pathname)) {
      const unitId = decodeURIComponent(url.pathname.split('/')[4] ?? '');
      if (!unitId) throw new StorageValidationError('Workbook id is required');
      let rawBody: unknown;
      try {
        rawBody = JSON.parse(await readBody(request)) as unknown;
      } catch {
        throw new StorageValidationError('History restore request must be valid JSON');
      }
      const restoreRequest = validateHistoryRestoreRequest(rawBody);
      const result = storage.restoreWorkbook(
        unitId,
        restoreRequest.targetRevision,
        restoreRequest.reason,
        principal.subject,
      );
      // REST restore and WebSocket changesets share the same committed
      // operation broadcast. Every connected peer receives the authoritative
      // server-generated mutation, including the requester if it also has a
      // collaboration socket.
      broadcastRevision(unitId, result.operation);
      sendJson(response, 200, result);
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/audit$/.test(url.pathname)) {
      const unitId = decodeURIComponent(url.pathname.split('/')[4] ?? '');
      if (!unitId) throw new StorageValidationError('Workbook id is required');
      sendJson(response, 200, { events: storage.listHistoryAudit(unitId, principal.subject) });
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/revisions\/\d+\/snapshot$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const revision = Number(parts[6]);
      if (!unitId) throw new Error('Workbook id is required');
      try {
        sendJson(response, 200, storage.getSnapshotAtRevision(unitId, revision, principal.subject));
      } catch (error) {
        sendJson(response, errorStatus(error), { code: protocolErrorCode(error), message: error instanceof Error ? error.message : 'Revision snapshot not found' });
      }
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/revisions\/\d+$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const revision = Number(parts[6]);
      if (!unitId) throw new Error('Workbook id is required');
      const result = storage.getRevision(unitId, revision, principal.subject);
      if (!result) {
        sendJson(response, 404, { code: 'NOT_FOUND', message: 'Revision not found' });
        return;
      }
      sendJson(response, 200, result);
      return;
    }

    if (request.method === 'POST' && /^\/api\/v1\/workbooks\/[^/]+\/calculations\/pivot$/.test(url.pathname)) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      const body = JSON.parse(await readBody(request)) as { pivotId?: string };
      if (!body.pivotId) throw new Error('Pivot id is required');
      const snapshot = storage.getSnapshot(unitId, principal.subject);
      const workbook = WorkbookModel.fromSnapshot(snapshot.snapshot);
      const pivot = workbook.getSheets().flatMap((sheet) => sheet.pivots).find((entry) => entry.id === body.pivotId);
      if (!pivot) {
        sendJson(response, 404, { code: 'NOT_FOUND', message: 'Pivot not found' });
        return;
      }
      const result = await calculationQueue.run(async () => computePivotResult(workbook, pivot));
      sendJson(response, 200, {
        unitId,
        pivotId: body.pivotId,
        revision: snapshot.revision,
        result,
      });
      return;
    }

    if (request.method === 'POST' && /^\/api\/v1\/workbooks\/[^/]+\/tables$/.test(url.pathname)) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      const input = JSON.parse(await readBody(request)) as Partial<import('@react-sheets/core-model').WorkbookTableModel>;
      if (!input.name || !input.fields || input.fields.length === 0) throw new Error('Table name and fields are required');
      const table: import('@react-sheets/core-model').WorkbookTableModel = {
        id: input.id ?? randomUUID(),
        name: input.name,
        sourceSheetId: input.sourceSheetId,
        rowCount: 0,
        fields: structuredClone(input.fields),
        blockSize: Math.max(1, Math.min(4096, input.blockSize ?? 4096)),
        blocks: [],
        revision: 0,
      };
      sendJson(response, 201, storage.createDataTable(unitId, table, principal.subject));
      return;
    }

    if (request.method === 'POST' && /^\/api\/v1\/workbooks\/[^/]+\/tables\/[^/]+\/blocks$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const tableId = parts[6];
      if (!unitId || !tableId) throw new Error('Workbook and table ids are required');
      const body = JSON.parse(await readBody(request)) as { startRow?: number; rows?: Array<Array<string | number | boolean | null>> };
      if (body.startRow == null || !body.rows) throw new Error('Block startRow and rows are required');
      sendJson(response, 201, storage.appendDataBlock(unitId, tableId, body.startRow, body.rows, principal.subject));
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/tables\/[^/]+\/rows$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const tableId = parts[6];
      if (!unitId || !tableId) throw new Error('Workbook and table ids are required');
      const offset = Math.max(0, Number(url.searchParams.get('offset') ?? 0) || 0);
      const limit = Math.max(1, Math.min(4096, Number(url.searchParams.get('limit') ?? 500) || 500));
      sendJson(response, 200, storage.readDataRows(unitId, tableId, offset, limit, principal.subject));
      return;
    }

    if (request.method === 'DELETE' && /^\/api\/v1\/workbooks\/[^/]+\/tables\/[^/]+$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const tableId = parts[6];
      if (!unitId || !tableId) throw new Error('Workbook and table ids are required');
      storage.removeDataTable(unitId, tableId, principal.subject);
      sendJson(response, 200, {});
      return;
    }

    if (request.method === 'POST' && url.pathname === '/api/v1/workbooks') {
      const input = JSON.parse(await readBody(request)) as {
        name?: string;
        snapshot?: ReturnType<typeof createDefaultSnapshot>;
      };
      const snapshot = input.snapshot ?? createDefaultSnapshot(randomUUID());
      snapshot.name = input.name?.trim() || snapshot.name;
      sendJson(response, 201, storage.createWorkbook(snapshot, principal.subject));
      return;
    }

    if (request.method === 'POST' && url.pathname === '/api/v1/files/import-xlsx') {
      const body = JSON.parse(await readBody(request)) as { base64?: string; fileName?: string };
      if (!body.base64) throw new Error('base64 payload is required');
      const imported = await importXlsx({
        fileName: body.fileName ?? 'import.xlsx',
        base64: body.base64,
        options: { compatibilityTarget: 'B' },
      });
      const result = storage.createWorkbook(imported.snapshot, principal.subject);
      sendJson(response, 200, { ...result, report: imported.report });
      return;
    }

    if (request.method === 'POST' && url.pathname === '/api/v1/files/import') {
      const body = JSON.parse(await readBody(request)) as { files?: Record<string, string>; snapshot?: unknown };
      let snapshot = body.snapshot ? (body.snapshot as ReturnType<typeof createDefaultSnapshot>) : undefined;
      if (!snapshot && body.files) {
        snapshot = parseXlsxXmlToSnapshot(body.files);
      }
      if (!snapshot) throw new Error('No valid workbook data provided to import');
      const result = storage.createWorkbook(snapshot, principal.subject);
      sendJson(response, 200, result);
      return;
    }

    if (request.method === 'GET' && url.pathname.startsWith('/api/v1/files/') && url.pathname.endsWith('/export')) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('File unitId is required');
      const snapshotRes = storage.getSnapshot(unitId, principal.subject);
      const fileName = url.searchParams.get('fileName') ?? `${snapshotRes.snapshot.name || 'workbook'}.xlsx`;
      const exported = await exportXlsx({
        snapshot: snapshotRes.snapshot,
        fileName,
        options: { compatibilityTarget: 'B', includeCachedValues: true },
      });
      sendJson(response, 200, {
        unitId,
        base64: exported.base64,
        fileName: exported.fileName,
        report: exported.report,
        snapshot: snapshotRes.snapshot,
      });
      return;
    }

    sendJson(response, 404, { code: 'NOT_FOUND', message: 'Route not found' });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Internal error';
    sendJson(response, errorStatus(error), { code: protocolErrorCode(error), message });
  }
});

const webSocketServer = new WebSocketServer({
  server,
  verifyClient: (info, done) => {
    void authenticator.authenticateRequest(info.req).then(
      () => done(true),
      (error) => done(false, errorStatus(error), error instanceof Error ? error.message : 'Bearer authentication failed'),
    );
  },
});

webSocketServer.on('connection', async (socket, request) => {
  let principal;
  try {
    // verifyClient protects the handshake; authenticate again here so the
    // connection handler never relies on a client-provided actor field.
    principal = await authenticator.authenticateRequest(request);
  } catch {
    socket.close(1008, 'Bearer authentication is required');
    return;
  }
  let clientUnitId: string | null = null;
  const clientActorId = principal.subject;

  socket.once('close', () => {
    const unitId = clientUnitId;
    leaveUnit(unitId, socket);
    if (!unitId) return;
    const clients = clientsByUnit.get(unitId);
    if (!clients) return;
    const message = encodeOperationMessageV2({ type: 'presence.broadcast', unitId, actorId: clientActorId, state: { status: 'offline' } });
    for (const client of clients) if (client.readyState === client.OPEN) client.send(message);
  });

  socket.on('message', (raw) => {
    void (async () => {
      let message: OperationMessageV2;
      try {
        message = decodeClientOperationMessageV2(raw.toString());
      } catch (error) {
        socket.send(
          encodeOperationMessageV2({
            type: 'changeset.reject',
            operationId: 'unknown',
            error: {
              code: 'VALIDATION_ERROR',
              message: error instanceof Error ? error.message : 'Invalid message',
            },
          }),
        );
        return;
      }

      try {
        if (message.type === 'changeset.submit') {
          if (message.payload.unitId !== clientUnitId) {
            leaveUnit(clientUnitId, socket);
            clientUnitId = message.payload.unitId;
          }
          const result = storage.appendOperation(message.payload, principal.subject);
          const unitClients = joinUnit(message.payload.unitId, socket);
          socket.send(
            encodeOperationMessageV2({
              type: 'changeset.ack',
              operationId: message.payload.operationId,
              revision: result.revision,
            }),
          );
          const revisionMessage = encodeOperationMessageV2({
            type: 'revision.created',
            payload: result.operation,
            revision: result.revision,
          });
          for (const client of unitClients) {
            if (client !== socket && client.readyState === client.OPEN) client.send(revisionMessage);
          }
          return;
        }

        if (message.type === 'presence.updated' || message.type === 'cursor.updated') {
          const role = storage.getRole(message.unitId, principal.subject);
          if (!role) throw new Error('Workbook access denied');
          leaveUnit(clientUnitId, socket);
          clientUnitId = message.unitId;
          const unitClients = joinUnit(clientUnitId, socket);
          const broadcast = encodeOperationMessageV2({
            type: message.type === 'presence.updated' ? 'presence.broadcast' : 'cursor.broadcast',
            unitId: message.unitId,
            actorId: principal.subject,
            state: message.state,
          });
          for (const client of unitClients) {
            if (client !== socket && client.readyState === client.OPEN) client.send(broadcast);
          }
          return;
        }

        if (message.type === 'snapshot.request') {
          const snapshotRes = storage.getSnapshot(message.unitId, principal.subject);
          // A snapshot request establishes the collaboration membership for
          // this workbook. Without joining here, a connected read-only peer
          // would miss a REST-triggered restore broadcast until it emitted a
          // later cursor/presence or changeset message.
          leaveUnit(clientUnitId, socket);
          clientUnitId = message.unitId;
          joinUnit(clientUnitId, socket);
          socket.send(
            encodeOperationMessageV2({
              type: 'snapshot.response',
              payload: snapshotRes,
            }),
          );
          return;
        }

        throw new Error('Unsupported client collaboration message');
      } catch (error) {
        socket.send(
          encodeOperationMessageV2({
            type: 'changeset.reject',
            operationId: message.type === 'changeset.submit' ? message.payload.operationId : 'unknown',
            error: {
              code: protocolErrorCode(error),
              message: error instanceof Error ? error.message : 'Message rejected',
            },
          }),
        );
      }
    })();
  });
});

server.listen(port, '127.0.0.1', () => {
  console.log(`React Sheets server listening on http://127.0.0.1:${port}`);
});
