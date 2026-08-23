import type { TableScalar } from '@react-sheets/core-model';
import { validateQuerySteps, type QueryDefinition, type QueryRefreshPolicy, type QueryStep } from './query-steps';

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

export interface QueryDefinitionPersistence {
  schema: 'QueryDefinitionV1';
  id: string;
  name: string;
  connectorId: string;
  /** Connector configuration is deliberately redacted before persistence. */
  connectorConfig: Record<string, unknown>;
  steps: QueryStep[];
  refreshOnOpen?: boolean;
  refreshPolicy?: QueryRefreshPolicy;
  sourceRevision: number;
}

/** DataConnector SPI. Only connectors registered by the host are executable. */
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
    schema: 'QueryDefinitionV1',
    id: query.id,
    name: query.name,
    connectorId: query.connectorId,
    connectorConfig: redactConnectorConfig(query.connectorConfig),
    steps: structuredClone(query.steps),
    ...(query.refreshOnOpen === undefined ? {} : { refreshOnOpen: query.refreshOnOpen }),
    ...(query.refreshPolicy === undefined ? {} : { refreshPolicy: structuredClone(query.refreshPolicy) }),
    sourceRevision: query.sourceRevision ?? 0,
  };
}

export function deserializeQueryDefinition(
  persisted: QueryDefinitionPersistence,
  secretConfig: Record<string, unknown> = {},
): QueryDefinition {
  if (persisted.schema !== 'QueryDefinitionV1') throw new Error('Unsupported query definition schema');
  validateQuerySteps(persisted.steps);
  return {
    id: persisted.id,
    name: persisted.name,
    connectorId: persisted.connectorId,
    connectorConfig: { ...structuredClone(persisted.connectorConfig), ...structuredClone(secretConfig) },
    steps: structuredClone(persisted.steps),
    ...(persisted.refreshOnOpen === undefined ? {} : { refreshOnOpen: persisted.refreshOnOpen }),
    ...(persisted.refreshPolicy === undefined ? {} : { refreshPolicy: structuredClone(persisted.refreshPolicy) }),
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

/** Built-in JSON connector — supports inline arrays or JSON strings. */
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
export * from './runtime';
