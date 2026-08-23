import React, { useMemo, useState } from 'react';
import type { QueryDefinition } from '@react-sheets/spreadsheet-app';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, Textarea } from '@react-sheets/ui-system';

export interface QueryPanelSnapshot {
  queryId: string;
  queryName: string;
  columns: readonly string[];
  rowCount: number;
  loadedAt: string;
}

export interface QueryPanelProps {
  connectors: readonly string[];
  loadedQueries: readonly QueryPanelSnapshot[];
  lastResult: QueryPanelSnapshot | null;
  canQuery: boolean;
  onLoadQuery: (query: QueryDefinition) => Promise<void>;
  onRefreshQuery: (queryId: string) => Promise<void>;
  onTestConnection: (connectorId: string, config: Record<string, unknown>) => Promise<{ ok: boolean; message?: string }>;
  onClose?: () => void;
}

const SAMPLE_JSON = `[
  { "Region": "East", "Product": "Alpha", "Units": 120, "Revenue": 4800 },
  { "Region": "West", "Product": "Beta", "Units": 95, "Revenue": 3325 },
  { "Region": "East", "Product": "Gamma", "Units": 64, "Revenue": 2560 }
]`;

export function QueryPanel({
  canQuery,
  connectors,
  lastResult,
  loadedQueries,
  onClose,
  onLoadQuery,
  onRefreshQuery,
  onTestConnection,
}: QueryPanelProps) {
  const [connectorId, setConnectorId] = useState(connectors[0] ?? 'json');
  const [jsonData, setJsonData] = useState(SAMPLE_JSON);
  const [filterRegion, setFilterRegion] = useState('');
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const queryDefinition = useMemo<QueryDefinition>(() => {
    const steps = filterRegion.trim()
      ? [{
          id: 'filter-region',
          kind: 'filter' as const,
          name: 'Filter Region',
          config: { column: 'Region', value: filterRegion.trim() },
          enabled: true,
        }]
      : [];
    return {
      id: 'inline-json-query',
      name: filterRegion.trim() ? `Inline JSON (${filterRegion.trim()})` : 'Inline JSON Query',
      connectorId,
      connectorConfig: connectorId === 'json' ? { data: jsonData } : { url: jsonData },
      steps,
    };
  }, [connectorId, filterRegion, jsonData]);

  const runTest = async () => {
    setBusy(true);
    setStatus(null);
    try {
      const result = await onTestConnection(connectorId, queryDefinition.connectorConfig);
      setStatus(result.ok ? (result.message ?? 'Connection OK') : (result.message ?? 'Connection failed'));
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Connection failed');
    } finally {
      setBusy(false);
    }
  };

  const runLoad = async () => {
    setBusy(true);
    setStatus(null);
    try {
      await onLoadQuery(queryDefinition);
      setStatus('Query loaded into the active sheet');
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Query load failed');
    } finally {
      setBusy(false);
    }
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Data Query</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Connector</Text>
            <Select
              value={connectorId}
              onChange={(event) => setConnectorId(event.target.value)}
              sizeVariant="sm"
              disabled={!canQuery}
            >
              {connectors.map((connector) => (
                <option key={connector} value={connector}>{connector.toUpperCase()}</option>
              ))}
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              {connectorId === 'json' ? 'JSON Records' : 'REST URL'}
            </Text>
            <Textarea
              value={jsonData}
              onChange={(event) => setJsonData(event.target.value)}
              rows={8}
              disabled={!canQuery}
              className="font-mono text-xs"
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Optional filter (Region)</Text>
            <Select
              value={filterRegion}
              onChange={(event) => setFilterRegion(event.target.value)}
              sizeVariant="sm"
              disabled={!canQuery}
            >
              <option value="">No filter</option>
              <option value="East">East</option>
              <option value="West">West</option>
            </Select>
          </Box>

          {status ? (
            <Box className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-700">
              {status}
            </Box>
          ) : null}

          {lastResult ? (
            <Box className="rounded-lg border border-emerald-100 bg-emerald-50/70 px-3 py-2 text-xs text-emerald-800">
              Last load: {lastResult.rowCount} rows × {lastResult.columns.length} columns
            </Box>
          ) : null}

          <Stack gap="sm">
            <Button variant="outline" size="sm" disabled={!canQuery || busy} onClick={() => { void runTest(); }}>
              Test connection
            </Button>
            <Button variant="primary" size="sm" icon="table" disabled={!canQuery || busy} onClick={() => { void runLoad(); }}>
              Load into sheet
            </Button>
          </Stack>

          {loadedQueries.length > 0 ? (
            <Stack gap="sm">
              <Text size="xs" weight="semibold" className="text-slate-700">Loaded queries</Text>
              {loadedQueries.map((query) => (
                <Panel key={query.queryId} className="shadow-none">
                  <PanelBody className="p-3">
                    <Stack gap="xs">
                      <Text size="sm" weight="semibold">{query.queryName}</Text>
                      <Text size="xs" tone="muted">{query.rowCount} rows · {query.columns.join(', ')}</Text>
                      <Button
                        size="xs"
                        variant="ghost"
                        disabled={!canQuery || busy}
                        onClick={() => { void onRefreshQuery(query.queryId); }}
                      >
                        Refresh
                      </Button>
                    </Stack>
                  </PanelBody>
                </Panel>
              ))}
            </Stack>
          ) : null}
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
