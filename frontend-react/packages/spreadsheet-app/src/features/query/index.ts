import type { TableScalar } from '@react-sheets/core-model';

export type ConnectorKind = 'csv' | 'tsv' | 'json' | 'rest' | 'xlsx' | 'sqlite';

export interface ConnectorContext {
  signal?: AbortSignal;
  baseUrl?: string;
}

export interface QueryResult {
  columns: string[];
  rows: TableScalar[][];
  rowCount: number;
}

/** DataConnector SPI — CSV/TSV/JSON/REST/XLSX/SQLite */
export interface DataConnector {
  readonly kind: ConnectorKind;
  readonly id: string;
  connect(config: Record<string, unknown>): Promise<void>;
  disconnect(): Promise<void>;
  executeQuery(query: string, context?: ConnectorContext): Promise<QueryResult>;
  testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }>;
}

export class ConnectorRegistry {
  private readonly connectors = new Map<string, DataConnector>();

  register(connector: DataConnector): void {
    if (this.connectors.has(connector.id)) throw new Error(`Connector already registered: ${connector.id}`);
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

/** 内置 JSON connector — 支持内联数组或 JSON 字符串 */
export class JsonDataConnector implements DataConnector {
  readonly kind = 'json' as const;
  readonly id = 'json';
  private rows: Record<string, unknown>[] = [];

  async connect(config: Record<string, unknown>): Promise<void> {
    const data = config.data;
    if (typeof data === 'string') {
      const parsed = JSON.parse(data) as unknown;
      if (!Array.isArray(parsed)) throw new Error('JSON data must be an array of objects');
      this.rows = parsed as Record<string, unknown>[];
      return;
    }
    if (!Array.isArray(data)) throw new Error('JSON connector requires data array');
    this.rows = data as Record<string, unknown>[];
  }

  async disconnect(): Promise<void> {
    this.rows = [];
  }

  async testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }> {
    try {
      await this.connect(config);
      return { ok: true, message: `${this.rows.length} record(s) ready` };
    } catch (error) {
      return { ok: false, message: error instanceof Error ? error.message : 'Invalid JSON data' };
    }
  }

  async executeQuery(_query: string): Promise<QueryResult> {
    const columns = this.rows.length > 0 ? Object.keys(this.rows[0]!) : [];
    return {
      columns,
      rows: this.rows.map((row) => columns.map((column) => (row[column] ?? null) as TableScalar)),
      rowCount: this.rows.length,
    };
  }
}

/** 内置 REST connector */
export class RestDataConnector implements DataConnector {
  readonly kind = 'rest' as const;
  readonly id = 'rest';
  private config: Record<string, unknown> = {};

  async connect(config: Record<string, unknown>): Promise<void> {
    this.config = config;
  }

  async disconnect(): Promise<void> {
    this.config = {};
  }

  async testConnection(config: Record<string, unknown>): Promise<{ ok: boolean; message?: string }> {
    const url = config.url as string | undefined;
    if (!url) return { ok: false, message: 'url is required' };
    return { ok: true };
  }

  async executeQuery(_query: string, context?: ConnectorContext): Promise<QueryResult> {
    const url = (this.config.url as string) ?? context?.baseUrl;
    if (!url) throw new Error('REST connector not configured');
    const response = await fetch(url, { signal: context?.signal });
    if (!response.ok) throw new Error(`REST query failed: ${response.status}`);
    const body = await response.json() as unknown;
    if (!Array.isArray(body)) throw new Error('REST response must be an array of objects');
    const rows = body as Record<string, unknown>[];
    const columns = rows.length > 0 ? Object.keys(rows[0]!) : [];
    return {
      columns,
      rows: rows.map((row) => columns.map((col) => (row[col] ?? null) as TableScalar)),
      rowCount: rows.length,
    };
  }
}

export function createDefaultConnectorRegistry(): ConnectorRegistry {
  const registry = new ConnectorRegistry();
  registry.register(new JsonDataConnector());
  registry.register(new RestDataConnector());
  return registry;
}

export * from './query-steps';
export * from './commands';
