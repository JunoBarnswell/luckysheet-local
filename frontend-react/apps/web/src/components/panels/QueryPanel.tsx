import React, { useEffect, useMemo, useState } from 'react';
import type { ConnectorManifest, QueryDefinition } from '@react-sheets/spreadsheet-app';
import { Box, Button, FileButton, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, Textarea, TextInput } from '@react-sheets/ui-system';

export interface QueryPanelSnapshot {
  queryId: string;
  queryName: string;
  columns: readonly string[];
  rowCount: number;
  loadedAt: string;
}

export interface QueryPanelProps {
  connectors: readonly ConnectorManifest[];
  loadedQueries: readonly QueryPanelSnapshot[];
  lastResult: QueryPanelSnapshot | null;
  canQuery: boolean;
  onLoadQuery: (query: QueryDefinition) => Promise<void>;
  onRefreshQuery: (queryId: string) => Promise<void>;
  onCancelQuery: (queryId: string) => Promise<void>;
  onTestConnection: (connectorId: string, config: Record<string, unknown>) => Promise<{ ok: boolean; message?: string }>;
  onClose?: () => void;
}

function allocateQueryId(): string {
  if (!globalThis.crypto?.randomUUID) throw new Error('QUERY_IDENTITY_UNAVAILABLE: Web Crypto randomUUID is required');
  return `query-${globalThis.crypto.randomUUID()}`;
}

function initialConfig(manifest: ConnectorManifest | undefined): Record<string, unknown> {
  if (!manifest) return {};
  return Object.fromEntries(manifest.fields.map((field) => [field.key, field.kind === 'select' ? (field.options?.[0]?.value ?? '') : '']));
}

export function assertConnectorConfig(manifest: ConnectorManifest, config: Record<string, unknown>): void {
  for (const field of manifest.fields) {
    if (!field.required) continue;
    const value = config[field.key];
    if (field.kind === 'file') {
      if (!(typeof File !== 'undefined' && value instanceof File)) throw new Error(`${field.label} is required`);
    } else if (typeof value !== 'string' || !value.trim()) {
      throw new Error(`${field.label} is required`);
    }
  }
  if ((manifest.kind === 'csv' || manifest.kind === 'tsv')
    && !(typeof config.text === 'string' && config.text.length > 0)
    && !(typeof File !== 'undefined' && config.file instanceof File)) {
    throw new Error(`${manifest.label} requires text or a file`);
  }
}

export function normalizeConnectorConfig(manifest: ConnectorManifest, source: Record<string, unknown>): Record<string, unknown> {
  assertConnectorConfig(manifest, source);
  const config = { ...source };
  if (typeof config.body === 'string' && config.body.trim()) {
    try { config.body = JSON.parse(config.body) as unknown; }
    catch { throw new Error('REST POST body must be valid JSON'); }
  } else if (config.body === '') {
    delete config.body;
  }
  if (typeof File !== 'undefined' && config.file instanceof File) config.fileName = config.file.name;
  return config;
}

export function QueryPanel({
  canQuery,
  connectors,
  lastResult,
  loadedQueries,
  onClose,
  onLoadQuery,
  onRefreshQuery,
  onCancelQuery,
  onTestConnection,
}: QueryPanelProps) {
  const [queryId, setQueryId] = useState(allocateQueryId);
  const [connectorId, setConnectorId] = useState(connectors[0]?.id ?? '');
  const manifest = useMemo(() => connectors.find((connector) => connector.id === connectorId), [connectorId, connectors]);
  const [connectorConfig, setConnectorConfig] = useState<Record<string, unknown>>(() => initialConfig(connectors[0]));
  const [filterColumn, setFilterColumn] = useState('');
  const [filterValue, setFilterValue] = useState('');
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const connectorAvailable = manifest?.available !== false;

  useEffect(() => {
    if (manifest || !connectors[0]) return;
    setConnectorId(connectors[0].id);
    setConnectorConfig(initialConfig(connectors[0]));
  }, [connectors, manifest]);

  const buildQuery = (): QueryDefinition => {
    if (!manifest) throw new Error('Select an available connector');
    const config = normalizeConnectorConfig(manifest, connectorConfig);
    const column = filterColumn.trim();
    const value = filterValue.trim();
    if ((column && !value) || (!column && value)) throw new Error('Filter column and value must be provided together');
    return {
      id: queryId,
      name: `${manifest.label} query`,
      connectorId: manifest.id,
      connectorConfig: config,
      steps: column ? [{ id: `filter-${column}`, kind: 'filter', name: `Filter ${column}`, config: { column, value }, enabled: true }] : [],
    };
  };

  const runTest = async () => {
    setBusy(true);
    setStatus(null);
    try {
      const query = buildQuery();
      const result = await onTestConnection(query.connectorId, query.connectorConfig);
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
      await onLoadQuery(buildQuery());
      setStatus('Query loaded into the active sheet');
      setQueryId(allocateQueryId());
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
              onChange={(event) => {
                const id = event.target.value;
                const next = connectors.find((connector) => connector.id === id);
                setConnectorId(id);
                setConnectorConfig(initialConfig(next));
                setStatus(null);
              }}
              sizeVariant="sm"
              disabled={!canQuery}
              options={connectors.map((connector) => ({ value: connector.id, label: `${connector.label}${connector.execution === 'server' ? ' (server)' : ''}` }))}
            />
          </Box>

          {manifest?.fields.map((field) => (
            <Box key={field.key}>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">{field.label}</Text>
              {field.kind === 'multiline-text' ? (
                <Textarea
                  value={typeof connectorConfig[field.key] === 'string' ? String(connectorConfig[field.key]) : ''}
                  onChange={(event) => setConnectorConfig((current) => ({ ...current, [field.key]: event.target.value }))}
                  rows={field.key === 'statement' ? 5 : 8}
                  placeholder={field.placeholder}
                  disabled={!canQuery || !connectorAvailable}
                  className="font-mono text-xs"
                />
              ) : field.kind === 'file' ? (
                <Stack gap="xs">
                  <FileButton
                    accept={field.accept}
                    onFile={(file) => setConnectorConfig((current) => ({ ...current, [field.key]: file }))}
                    variant="outline"
                    size="sm"
                    disabled={!canQuery || !connectorAvailable}
                  >
                    Select file
                  </FileButton>
                  {typeof File !== 'undefined' && connectorConfig[field.key] instanceof File
                    ? <Text size="xs" tone="muted">{(connectorConfig[field.key] as File).name}</Text>
                    : null}
                </Stack>
              ) : field.kind === 'select' ? (
                <Select
                  value={typeof connectorConfig[field.key] === 'string' ? String(connectorConfig[field.key]) : ''}
                  onChange={(event) => setConnectorConfig((current) => ({ ...current, [field.key]: event.target.value }))}
                  sizeVariant="sm"
                  disabled={!canQuery || !connectorAvailable}
                  options={field.options}
                />
              ) : (
                <TextInput
                  value={typeof connectorConfig[field.key] === 'string' ? String(connectorConfig[field.key]) : ''}
                  onChange={(event) => setConnectorConfig((current) => ({ ...current, [field.key]: event.target.value }))}
                  placeholder={field.placeholder}
                  disabled={!canQuery || !connectorAvailable}
                />
              )}
            </Box>
          ))}

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Optional filter column</Text>
            <TextInput value={filterColumn} onChange={(event) => setFilterColumn(event.target.value)} disabled={!canQuery || !connectorAvailable} />
          </Box>
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">Optional filter value</Text>
            <TextInput value={filterValue} onChange={(event) => setFilterValue(event.target.value)} disabled={!canQuery || !connectorAvailable} />
          </Box>

          {manifest?.unavailableReason ? <Text size="xs" tone="danger">{manifest.unavailableReason}</Text> : null}
          {status ? <Box className="rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 text-xs text-slate-700">{status}</Box> : null}
          {lastResult ? <Box className="rounded-lg border border-emerald-100 bg-emerald-50/70 px-3 py-2 text-xs text-emerald-800">Last load: {lastResult.rowCount} rows × {lastResult.columns.length} columns</Box> : null}

          <Stack gap="sm">
            <Button variant="outline" size="sm" disabled={!canQuery || busy || !manifest || !connectorAvailable} onClick={() => { void runTest(); }}>Test connection</Button>
            <Button variant="primary" size="sm" icon="table" disabled={!canQuery || busy || !manifest || !connectorAvailable} onClick={() => { void runLoad(); }}>Load into sheet</Button>
            {busy && manifest?.execution === 'server' ? (
              <Button variant="danger" size="sm" onClick={() => {
                void onCancelQuery(queryId).then(() => setStatus('Query cancellation requested')).catch((cause) => setStatus(cause instanceof Error ? cause.message : 'Query cancellation failed'));
              }}>Cancel query</Button>
            ) : null}
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
                      <Button size="xs" variant="ghost" disabled={!canQuery || busy} onClick={() => { void onRefreshQuery(query.queryId); }}>Refresh</Button>
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
