import { createServer } from 'node:http';
import { randomUUID } from 'node:crypto';
import { WebSocketServer } from 'ws';
import { WorkbookModel } from '@react-sheets/core-model';
import { decodeMessage, encodeMessage } from '@react-sheets/protocol';
import type { CollaborationChangeSet } from '@react-sheets/protocol';
import { WorkbookStorage } from '@react-sheets/storage';
import { exportSnapshotToXlsxXml, parseXlsxXmlToSnapshot } from '@react-sheets/pro-features';

const storage = new WorkbookStorage();
const port = Number(process.env.REACT_SHEETS_SERVER_PORT ?? 4181);
const clients = new Set<import('ws').WebSocket>();

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

const server = createServer(async (request, response) => {
  try {
    const url = new URL(request.url ?? '/', `http://${request.headers.host ?? '127.0.0.1'}`);

    if (request.method === 'GET' && url.pathname === '/health') {
      sendJson(response, 200, { status: 'ok', service: 'react-sheets-server' });
      return;
    }

    if (request.method === 'GET' && url.pathname.startsWith('/api/v1/workbooks/') && url.pathname.endsWith('/snapshot')) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      sendJson(response, 200, storage.getSnapshot(unitId));
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/revisions$/.test(url.pathname)) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      sendJson(response, 200, { revisions: storage.listRevisions(unitId) });
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/revisions\/\d+$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const revision = Number(parts[6]);
      if (!unitId) throw new Error('Workbook id is required');
      const result = storage.getRevision(unitId, revision);
      if (!result) {
        sendJson(response, 404, { code: 'NOT_FOUND', message: 'Revision not found' });
        return;
      }
      sendJson(response, 200, result);
      return;
    }

    if (request.method === 'POST' && url.pathname === '/api/v1/workbooks') {
      const input = JSON.parse(await readBody(request)) as {
        name?: string;
        snapshot?: ReturnType<typeof createDefaultSnapshot>;
      };
      const snapshot = input.snapshot ?? createDefaultSnapshot(randomUUID());
      snapshot.name = input.name?.trim() || snapshot.name;
      sendJson(response, 201, storage.createWorkbook(snapshot));
      return;
    }

    if (request.method === 'POST' && url.pathname.startsWith('/api/v1/workbooks/') && url.pathname.endsWith('/changesets')) {
      const body = JSON.parse(await readBody(request)) as CollaborationChangeSet;
      const revision = storage.appendChangeSet(body);
      sendJson(response, 200, { operationId: body.operationId, revision });
      return;
    }

    if (request.method === 'POST' && url.pathname === '/api/v1/files/import') {
      const body = JSON.parse(await readBody(request)) as { files?: Record<string, string>; snapshot?: unknown };
      let snapshot = body.snapshot ? (body.snapshot as ReturnType<typeof createDefaultSnapshot>) : undefined;
      if (!snapshot && body.files) {
        snapshot = parseXlsxXmlToSnapshot(body.files);
      }
      if (!snapshot) throw new Error('No valid workbook data provided to import');
      const result = storage.createWorkbook(snapshot);
      sendJson(response, 200, result);
      return;
    }

    if (request.method === 'GET' && url.pathname.startsWith('/api/v1/files/') && url.pathname.endsWith('/export')) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('File unitId is required');
      const snapshotRes = storage.getSnapshot(unitId);
      const xmlFiles = exportSnapshotToXlsxXml(snapshotRes.snapshot);
      sendJson(response, 200, { unitId, files: xmlFiles, snapshot: snapshotRes.snapshot });
      return;
    }

    sendJson(response, 404, { code: 'NOT_FOUND', message: 'Route not found' });
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Internal error';
    sendJson(response, message.includes('not found') ? 404 : 409, { code: 'REQUEST_FAILED', message });
  }
});

const webSocketServer = new WebSocketServer({ server });
webSocketServer.on('connection', (socket) => {
  clients.add(socket);
  socket.once('close', () => clients.delete(socket));
  socket.send(
    encodeMessage({
      type: 'presence.updated',
      unitId: '',
      actorId: 'server',
      state: { status: 'connected' },
    }),
  );

  socket.on('message', (raw) => {
    try {
      const message = decodeMessage(raw.toString());

      if (message.type === 'changeset.submit') {
        const revision = storage.appendChangeSet(message.payload);
        socket.send(
          encodeMessage({
            type: 'changeset.ack',
            operationId: message.payload.operationId,
            revision,
          }),
        );
        const revisionMessage = encodeMessage({
          type: 'revision.created',
          payload: message.payload,
          revision,
        });
        for (const client of clients) {
          if (client !== socket && client.readyState === client.OPEN) {
            client.send(revisionMessage);
          }
        }
      } else if (message.type === 'presence.updated' || message.type === 'cursor.updated') {
        // Broadcast presence/cursor to other clients
        const broadcast = encodeMessage(message as never);
        for (const client of clients) {
          if (client !== socket && client.readyState === client.OPEN) {
            client.send(broadcast);
          }
        }
      } else if (message.type === 'snapshot.request') {
        const snapshotRes = storage.getSnapshot(message.unitId);
        socket.send(
          encodeMessage({
            type: 'snapshot.response',
            unitId: message.unitId,
            snapshot: snapshotRes.snapshot,
            revision: snapshotRes.revision,
          }),
        );
      }
    } catch (error) {
      socket.send(
        encodeMessage({
          type: 'changeset.reject',
          operationId: 'unknown',
          error: {
            code: 'VALIDATION_ERROR',
            message: error instanceof Error ? error.message : 'Invalid message',
          },
        }),
      );
    }
  });
});

server.listen(port, '127.0.0.1', () => {
  console.log(`React Sheets server listening on http://127.0.0.1:${port}`);
});
