import { createServer } from 'node:http';
import { randomUUID } from 'node:crypto';
import { inflateRawSync } from 'node:zlib';
import { WebSocketServer } from 'ws';
import { WorkbookModel } from '@react-sheets/core-model';
import { decodeMessage, encodeMessage } from '@react-sheets/protocol';
import { WorkbookStorage } from '@react-sheets/storage';
import { computePivotResult, exportSnapshotToXlsxXml, parseXlsxXmlToSnapshot } from '@react-sheets/pro-features';

/** 解析上传的 XLSX(zip)Base64:解 STORE/DEFLATE 条目并返回文件名到文本内容映射 */
function unzipXlsxBase64(base64: string): Record<string, string> {
  const buffer = Buffer.from(base64.replace(/^data:[^;]+;base64,/, ''), 'base64');
  const files: Record<string, string> = {};
  let offset = 0;
  while (offset + 4 <= buffer.length) {
    const signature = buffer.readUInt32LE(offset);
    if (signature !== 0x04034b50) break;
    const method = buffer.readUInt16LE(offset + 8);
    const compressedSize = buffer.readUInt32LE(offset + 18);
    const nameLength = buffer.readUInt16LE(offset + 26);
    const extraLength = buffer.readUInt16LE(offset + 28);
    const nameStart = offset + 30;
    const name = buffer.subarray(nameStart, nameStart + nameLength).toString('utf8');
    const dataStart = nameStart + nameLength + extraLength;
    const rawData = buffer.subarray(dataStart, dataStart + compressedSize);
    try {
      files[name] = (method === 0 ? rawData : inflateRawSync(rawData)).toString('utf8');
    } catch {
      // 跳过无法解码的条目(如内嵌图片)
    }
    offset = dataStart + compressedSize;
  }
  return files;
}

const storage = new WorkbookStorage();
const port = Number(process.env.REACT_SHEETS_SERVER_PORT ?? 4181);
const clientsByUnit = new Map<string, Set<import('ws').WebSocket>>();

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

    if (request.method === 'GET' && url.pathname === '/api/v1/workbooks') {
      sendJson(response, 200, storage.listWorkbooks());
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

    if (request.method === 'POST' && /^\/api\/v1\/workbooks\/[^/]+\/calculations\/pivot$/.test(url.pathname)) {
      const unitId = url.pathname.split('/')[4];
      if (!unitId) throw new Error('Workbook id is required');
      const body = JSON.parse(await readBody(request)) as { pivotId?: string };
      if (!body.pivotId) throw new Error('Pivot id is required');
      const snapshot = storage.getSnapshot(unitId);
      const workbook = WorkbookModel.fromSnapshot(snapshot.snapshot);
      const pivot = workbook.getSheets().flatMap((sheet) => sheet.pivots).find((entry) => entry.id === body.pivotId);
      if (!pivot) {
        sendJson(response, 404, { code: 'NOT_FOUND', message: 'Pivot not found' });
        return;
      }
      sendJson(response, 200, {
        unitId,
        pivotId: body.pivotId,
        revision: snapshot.revision,
        result: computePivotResult(workbook, pivot),
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
      sendJson(response, 201, storage.createDataTable(unitId, table));
      return;
    }

    if (request.method === 'POST' && /^\/api\/v1\/workbooks\/[^/]+\/tables\/[^/]+\/blocks$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const tableId = parts[6];
      if (!unitId || !tableId) throw new Error('Workbook and table ids are required');
      const body = JSON.parse(await readBody(request)) as { startRow?: number; rows?: Array<Array<string | number | boolean | null>> };
      if (body.startRow == null || !body.rows) throw new Error('Block startRow and rows are required');
      sendJson(response, 201, storage.appendDataBlock(unitId, tableId, body.startRow, body.rows));
      return;
    }

    if (request.method === 'GET' && /^\/api\/v1\/workbooks\/[^/]+\/tables\/[^/]+\/rows$/.test(url.pathname)) {
      const parts = url.pathname.split('/');
      const unitId = parts[4];
      const tableId = parts[6];
      if (!unitId || !tableId) throw new Error('Workbook and table ids are required');
      const offset = Math.max(0, Number(url.searchParams.get('offset') ?? 0) || 0);
      const limit = Math.max(1, Math.min(4096, Number(url.searchParams.get('limit') ?? 500) || 500));
      sendJson(response, 200, storage.readDataRows(unitId, tableId, offset, limit));
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

    if (request.method === 'POST' && url.pathname === '/api/v1/files/import-xlsx') {
      const body = JSON.parse(await readBody(request)) as { base64?: string };
      if (!body.base64) throw new Error('base64 payload is required');
      const files = unzipXlsxBase64(body.base64);
      if (!files['xl/workbook.xml']) throw new Error('Not a valid XLSX package');
      const snapshot = parseXlsxXmlToSnapshot(files);
      const result = storage.createWorkbook(snapshot);
      sendJson(response, 200, result);
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
  let clientUnitId: string | null = null;
  socket.once('close', () => leaveUnit(clientUnitId, socket));
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
        try {
          leaveUnit(clientUnitId, socket);
          clientUnitId = message.payload.unitId;
          const unitClients = joinUnit(clientUnitId, socket);
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
          for (const client of unitClients) {
            if (client !== socket && client.readyState === client.OPEN) {
              client.send(revisionMessage);
            }
          }
        } catch (error) {
          socket.send(
            encodeMessage({
              type: 'changeset.reject',
              operationId: message.payload.operationId,
              error: {
                code: 'CONFLICT',
                message: error instanceof Error ? error.message : 'Changeset rejected',
              },
            }),
          );
        }
      } else if (message.type === 'presence.updated' || message.type === 'cursor.updated') {
        leaveUnit(clientUnitId, socket);
        clientUnitId = message.unitId;
        const unitClients = joinUnit(clientUnitId, socket);
        // Broadcast presence/cursor to other clients
        const broadcast = encodeMessage(message as never);
        for (const client of unitClients) {
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
