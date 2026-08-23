import type { QueryDefinitionSnapshot, TableScalar } from '@react-sheets/core-model';
import { validateQuerySteps, type QueryDefinition, type QueryRefreshPolicy, type QueryStep } from './query-steps';

export type ConnectorKind = 'csv' | 'tsv' | 'json' | 'rest' | 'xlsx' | 'sqlite' | 'jdbc';
export type ConnectorExecution = 'local' | 'server';

export interface ConnectorContext {
  signal?: AbortSignal;
  baseUrl?: string;
}

export interface QueryResult {
  columns: string[];
  rows: TableScalar[][];
  rowCount: number;
}

/** Canonical persistence-safe workbook state. */
export type QueryDefinitionPersistence = QueryDefinitionSnapshot;

/** DataConnector SPI. Only connectors registered by the host are executable. */
export interface DataConnector {
  readonly kind: ConnectorKind;
  readonly id: string;
  readonly execution: ConnectorExecution;
  connect(config: Record<string, unknown>): Promise<void>;
  disconnect(): Promise<void>;
  executeQuery(query: string, context?: ConnectorContext): Promise<QueryResult>;
  testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }>;
}

export class ConnectorRegistry {
  private readonly connectors = new Map<string, DataConnector>();

  register(connector: DataConnector): void {
    if (this.connectors.has(connector.id)) throw new Error(`Connector already registered: ${connector.id}`);
    if (!connector.id.trim() || !connector.kind) throw new Error('A connector must declare id and kind');
    this.connectors.set(connector.id, connector);
  }

  get(id: string): DataConnector {
    const connector = this.connectors.get(id);
    if (!connector) throw new Error(`Unknown connector: ${id}`);
    return connector;
  }

  list(): DataConnector[] {
    return [...this.connectors.values()];
  }
}

/**
 * Return a persistence-safe query definition.  Secrets and bearer material are
 * never written to workbook snapshots or changesets.  The caller that owns a
 * credential vault can merge secrets back at execution time.
 */
export function serializeQueryDefinition(query: QueryDefinition): QueryDefinitionPersistence {
  if (!query.id.trim() || !query.name.trim()) throw new Error('Query id and name are required');
  if (!query.connectorId.trim()) throw new Error('Query connectorId is required');
  validateQuerySteps(query.steps);
  return {
    schema: 'QueryDefinition',
    id: query.id,
    name: query.name,
    connectorId: query.connectorId,
    connectorConfig: redactConnectorConfig(query.connectorConfig),
    steps: structuredClone(query.steps),
    ...(query.refreshOnOpen === undefined ? {} : { refreshOnOpen: query.refreshOnOpen }),
    ...(query.refreshPolicy === undefined ? {} : { refreshPolicy: structuredClone(query.refreshPolicy) }),
    ...(query.lastTarget === undefined ? {} : { lastTarget: structuredClone(query.lastTarget) }),
    sourceRevision: query.sourceRevision ?? 0,
  };
}

export function deserializeQueryDefinition(
  persisted: QueryDefinitionPersistence,
  secretConfig: Record<string, unknown> = {},
): QueryDefinition {
  if (persisted.schema !== 'QueryDefinition') throw new Error('Unsupported query definition schema');
  validateQuerySteps(persisted.steps);
  return {
    id: persisted.id,
    name: persisted.name,
    connectorId: persisted.connectorId,
    connectorConfig: { ...structuredClone(persisted.connectorConfig), ...structuredClone(secretConfig) },
    steps: structuredClone(persisted.steps),
    ...(persisted.refreshOnOpen === undefined ? {} : { refreshOnOpen: persisted.refreshOnOpen }),
    ...(persisted.refreshPolicy === undefined ? {} : { refreshPolicy: structuredClone(persisted.refreshPolicy) }),
    ...(persisted.lastTarget === undefined ? {} : { lastTarget: structuredClone(persisted.lastTarget) }),
    sourceRevision: persisted.sourceRevision,
  };
}

const SECRET_KEY = /(?:pass(word)?|secret|token|api[-_]?key|credential|authorization|private[-_]?key|client[-_]?secret)/i;

function redactConnectorConfig(config: Record<string, unknown>): Record<string, unknown> {
  const redact = (value: unknown, key?: string): unknown => {
    if (key && SECRET_KEY.test(key)) return '[redacted]';
    if (Array.isArray(value)) return value.map((item) => redact(item));
    if (value && typeof value === 'object') {
      return Object.fromEntries(Object.entries(value as Record<string, unknown>).map(([entryKey, entryValue]) => [entryKey, redact(entryValue, entryKey)]));
    }
    if (typeof value === 'function' || typeof value === 'bigint' || typeof value === 'symbol' || value === undefined) {
      throw new Error('Query connector configuration is not serializable');
    }
    return value;
  };
  return redact(config) as Record<string, unknown>;
}

function asScalar(value: unknown): TableScalar {
  if (value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  throw new Error('Query connector returned a non-scalar value');
}

function recordsToResult(records: Record<string, unknown>[]): QueryResult {
  const columns = [...new Set(records.flatMap((record) => Object.keys(record)))];
  const rows = records.map((record) => columns.map((column) => record[column] == null ? null : asScalar(record[column])));
  return { columns, rows, rowCount: rows.length };
}

function parseJsonRecords(value: unknown): QueryResult {
  const parsed = typeof value === 'string' ? JSON.parse(value) as unknown : value;
  if (!Array.isArray(parsed) || parsed.some((row) => !row || typeof row !== 'object' || Array.isArray(row))) {
    throw new Error('JSON connector requires an array of objects');
  }
  return recordsToResult(parsed as Record<string, unknown>[]);
}

function parseDelimited(text: string, delimiter: string): QueryResult {
  const source = text.replace(/^\uFEFF/, '');
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]!;
    if (character === '"') {
      if (quoted && source[index + 1] === '"') { cell += '"'; index += 1; }
      else quoted = !quoted;
    } else if (character === delimiter && !quoted) {
      row.push(cell); cell = '';
    } else if ((character === '\n' || character === '\r') && !quoted) {
      if (character === '\r' && source[index + 1] === '\n') index += 1;
      row.push(cell); rows.push(row); row = []; cell = '';
    } else cell += character;
  }
  if (quoted) throw new Error('Delimited query input contains an unterminated quote');
  if (cell.length > 0 || row.length > 0) { row.push(cell); rows.push(row); }
  if (rows.length === 0) return { columns: [], rows: [], rowCount: 0 };
  const columns = rows.shift()!.map((column, index) => column.trim() || `Column${index + 1}`);
  const seen = new Set<string>();
  for (const column of columns) {
    if (seen.has(column)) throw new Error(`Delimited query input contains duplicate column "${column}"`);
    seen.add(column);
  }
  const values = rows.map((values) => {
    if (values.length !== columns.length) throw new Error('Delimited query row width does not match the header');
    return values.map(parseDelimitedScalar);
  });
  return { columns, rows: values, rowCount: values.length };
}

function parseDelimitedScalar(value: string): TableScalar {
  if (value === '') return null;
  if (value === 'true') return true;
  if (value === 'false') return false;
  const number = Number(value);
  return Number.isFinite(number) && value.trim() !== '' ? number : value;
}

async function readText(config: Record<string, unknown>): Promise<string> {
  if (typeof config.text === 'string') return config.text;
  if (typeof config.data === 'string') return config.data;
  const file = config.file;
  if (file && typeof file === 'object' && 'text' in file && typeof (file as { text?: unknown }).text === 'function') return (file as Blob).text();
  throw new Error('Local text connector requires text, data, or file input');
}

async function readBytes(config: Record<string, unknown>): Promise<Uint8Array> {
  const data = config.bytes ?? config.buffer ?? config.file;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (data && typeof data === 'object' && 'arrayBuffer' in data && typeof (data as { arrayBuffer?: unknown }).arrayBuffer === 'function') return new Uint8Array(await (data as Blob).arrayBuffer());
  const base64 = config.base64;
  if (typeof base64 === 'string') {
    const buffer = (globalThis as { Buffer?: { from(value: string, encoding: string): Uint8Array } }).Buffer;
    const binary = typeof atob === 'function' ? atob(base64) : String.fromCharCode(...(buffer?.from(base64, 'base64') ?? new Uint8Array()));
    return Uint8Array.from(binary, (character) => character.charCodeAt(0));
  }
  throw new Error('Local binary connector requires bytes, buffer, file, or base64 input');
}

function snapshotSheetResult(snapshot: import('@react-sheets/core-model').WorkbookSnapshot['sheets'][number]): QueryResult {
  const matrix = new Map<number, Map<number, TableScalar>>();
  let maxRow = -1;
  let maxColumn = -1;
  for (const [rowKey, columns] of Object.entries(snapshot.cells)) {
    const row = Number(rowKey);
    if (!Number.isInteger(row)) continue;
    const rowValues = matrix.get(row) ?? new Map<number, TableScalar>();
    for (const [columnKey, cell] of Object.entries(columns)) {
      const column = Number(columnKey);
      if (!Number.isInteger(column)) continue;
      rowValues.set(column, cell.value);
      maxColumn = Math.max(maxColumn, column);
    }
    matrix.set(row, rowValues); maxRow = Math.max(maxRow, row);
  }
  if (maxRow < 0 || maxColumn < 0) return { columns: [], rows: [], rowCount: 0 };
  const grid = Array.from({ length: maxRow + 1 }, (_, row) => Array.from({ length: maxColumn + 1 }, (_, column) => matrix.get(row)?.get(column) ?? null));
  const header = grid.shift()!.map((value, index) => value == null || value === '' ? `Column${index + 1}` : String(value));
  const columns = header.map((value, index) => {
    const candidate = value.trim() || `Column${index + 1}`;
    return header.slice(0, index).includes(candidate) ? `${candidate}_${index + 1}` : candidate;
  });
  const rows = grid.map((row) => row.slice(0, columns.length));
  return { columns, rows, rowCount: rows.length };
}

/** Built-in JSON connector for in-memory local data. */
export class JsonDataConnector implements DataConnector {
  readonly kind = 'json' as const;
  readonly id = 'json';
  readonly execution = 'local' as const;
  private result: QueryResult = { columns: [], rows: [], rowCount: 0 };

  async connect(config: Record<string, unknown>): Promise<void> { this.result = parseJsonRecords(config.data ?? config.text); }
  async disconnect(): Promise<void> { this.result = { columns: [], rows: [], rowCount: 0 }; }
  async testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }> { try { await this.connect(config); return { ok: true, message: `${this.result.rowCount} record(s) ready` }; } catch (error) { return { ok: false, message: error instanceof Error ? error.message : 'Invalid JSON data' }; } }
  async executeQuery(_query: string): Promise<QueryResult> { return structuredClone(this.result); }
}

export class CsvDataConnector implements DataConnector {
  readonly kind: ConnectorKind = 'csv';
  readonly id: string = 'csv';
  readonly execution = 'local' as const;
  protected result: QueryResult = { columns: [], rows: [], rowCount: 0 };
  protected delimiter = ',';
  async connect(config: Record<string, unknown>): Promise<void> { this.result = parseDelimited(await readText(config), this.delimiter); }
  async disconnect(): Promise<void> { this.result = { columns: [], rows: [], rowCount: 0 }; }
  async testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }> { try { await this.connect(config); return { ok: true, message: `${this.result.rowCount} record(s) ready` }; } catch (error) { return { ok: false, message: error instanceof Error ? error.message : 'Invalid CSV data' }; } }
  async executeQuery(_query: string): Promise<QueryResult> { return structuredClone(this.result); }
}

export class TsvDataConnector extends CsvDataConnector {
  readonly kind: ConnectorKind = 'tsv';
  readonly id = 'tsv';
  protected delimiter = '\t';
}

export class XlsxDataConnector implements DataConnector {
  readonly kind = 'xlsx' as const;
  readonly id = 'xlsx';
  readonly execution = 'local' as const;
  private result: QueryResult = { columns: [], rows: [], rowCount: 0 };
  async connect(config: Record<string, unknown>): Promise<void> {
    const bytes = await readBytes(config);
    const { importXlsx } = await import('@react-sheets/exchange-xlsx');
    const imported = await importXlsx({ fileName: typeof config.fileName === 'string' ? config.fileName : 'query.xlsx', buffer: bytes.slice().buffer as ArrayBuffer, options: { compatibilityTarget: 'A' } });
    const first = imported.snapshot.sheets[0];
    if (!first) throw new Error('XLSX workbook contains no worksheets');
    this.result = snapshotSheetResult(first);
  }
  async disconnect(): Promise<void> { this.result = { columns: [], rows: [], rowCount: 0 }; }
  async testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }> { try { await this.connect(config); return { ok: true, message: `${this.result.rowCount} record(s) ready` }; } catch (error) { return { ok: false, message: error instanceof Error ? error.message : 'Invalid XLSX data' }; } }
  async executeQuery(_query: string): Promise<QueryResult> { return structuredClone(this.result); }
}

/** Server-only connector descriptor. It is never registered by the local default registry. */
export class RestDataConnector implements DataConnector {
  readonly kind = 'rest' as const;
  readonly id = 'rest';
  readonly execution = 'server' as const;
  async connect(): Promise<void> { throw new Error('REST connector is server-only'); }
  async disconnect(): Promise<void> {}
  async testConnection(): Promise<{ ok: boolean; message?: string }> { return { ok: false, message: 'REST connector is server-only' }; }
  async executeQuery(): Promise<QueryResult> { throw new Error('REST connector is server-only'); }
}

export function createDefaultConnectorRegistry(): ConnectorRegistry {
  const registry = new ConnectorRegistry();
  registry.register(new JsonDataConnector());
  registry.register(new CsvDataConnector());
  registry.register(new TsvDataConnector());
  registry.register(new XlsxDataConnector());
  return registry;
}

export * from './query-steps';
export * from './commands';
export * from './runtime';
