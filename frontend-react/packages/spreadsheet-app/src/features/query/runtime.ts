import type { CellData, RangeRef } from '@react-sheets/core-model';
import type { ConnectorRegistry, QueryResult } from './index';
import {
  QueryStepPipeline,
  type LoadTarget,
  type QueryDefinition,
} from './query-steps';

export interface QueryResultSnapshot {
  queryId: string;
  queryName: string;
  columns: string[];
  rowCount: number;
  loadedAt: string;
  target: LoadTarget;
}

export interface QuerySessionEntry {
  definition: QueryDefinition;
  lastResult?: QueryResultSnapshot;
}

export interface QueryLoadCommandPayload {
  query: QueryDefinition;
  target: LoadTarget;
  result: QueryResult;
}

export function queryResultToRangeValues(result: QueryResult): CellData[][] {
  return [
    result.columns.map((name) => ({ value: name })),
    ...result.rows.map((row) => row.map((value) => ({ value: value ?? '' }))),
  ];
}

export async function executeQueryDefinition(
  connectors: ConnectorRegistry,
  query: QueryDefinition,
): Promise<QueryResult> {
  const connector = connectors.get(query.connectorId);
  await connector.connect(query.connectorConfig);
  try {
    const raw = await connector.executeQuery(query.id);
    const pipeline = new QueryStepPipeline(query.steps);
    const transformed = pipeline.applySteps({
      columns: raw.columns,
      rows: raw.rows,
    });
    return {
      columns: transformed.columns,
      rows: transformed.rows as QueryResult['rows'],
      rowCount: transformed.rows.length,
    };
  } finally {
    await connector.disconnect();
  }
}

export function resolveLoadTarget(activeSheetId: string, selectionRange: RangeRef): LoadTarget {
  return {
    kind: 'range',
    sheetId: activeSheetId,
    range: {
      startRow: selectionRange.startRow,
      startColumn: selectionRange.startColumn,
    },
  };
}

export function buildQueryResultSnapshot(
  query: QueryDefinition,
  result: QueryResult,
  target: LoadTarget,
): QueryResultSnapshot {
  return {
    queryId: query.id,
    queryName: query.name,
    columns: [...result.columns],
    rowCount: result.rowCount,
    loadedAt: new Date().toISOString(),
    target,
  };
}

export function summarizeQueryResult(snapshot: QueryResultSnapshot): string {
  return `Loaded ${snapshot.rowCount} rows × ${snapshot.columns.length} columns into the sheet`;
}

export function createInlineJsonQuery(
  id: string,
  name: string,
  data: Record<string, unknown>[],
  steps: QueryDefinition['steps'] = [],
): QueryDefinition {
  return {
    id,
    name,
    connectorId: 'json',
    connectorConfig: { data },
    steps,
  };
}
