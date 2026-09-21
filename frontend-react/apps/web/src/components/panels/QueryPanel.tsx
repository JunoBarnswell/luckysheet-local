import React, { useMemo, useState } from 'react';
import type { QueryDefinition, QueryStep } from '@react-sheets/spreadsheet-app';
import { Box, Button, CheckToggle, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput, Textarea } from '@react-sheets/ui-system';

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

type RecipeKind = 'trim-text' | 'split-column' | 'remove-duplicates' | 'sort';

function parseColumnList(value: string): string[] {
  return [...new Set(value.split(',').map((entry) => entry.trim()).filter(Boolean))];
}

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
  const [recipeKind, setRecipeKind] = useState<RecipeKind>('trim-text');
  const [recipeColumn, setRecipeColumn] = useState('');
  const [recipeColumns, setRecipeColumns] = useState('');
  const [recipeValue, setRecipeValue] = useState(',');
  const [recipeOutputs, setRecipeOutputs] = useState('');
  const [recipeAscending, setRecipeAscending] = useState(true);
  const [recipeSteps, setRecipeSteps] = useState<QueryStep[]>([]);
  const [recipeError, setRecipeError] = useState<string | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const queryDefinition = useMemo<QueryDefinition>(() => {
    const quickFilter = filterRegion.trim()
      ? [{
          id: 'filter-region',
          kind: 'filter' as const,
          name: 'Filter Region',
          config: { column: 'Region', value: filterRegion.trim() },
          enabled: true,
        }]
      : [];
    const steps = [...quickFilter, ...recipeSteps];
    return {
      id: 'inline-json-query',
      name: filterRegion.trim() ? `Inline JSON (${filterRegion.trim()})` : recipeSteps.length > 0 ? 'Inline JSON (cleaning recipe)' : 'Inline JSON Query',
      connectorId,
      connectorConfig: connectorId === 'json' ? { data: jsonData } : { url: jsonData },
      steps,
    };
  }, [connectorId, filterRegion, jsonData, recipeSteps]);

  const addRecipeStep = () => {
    const columns = parseColumnList(recipeColumns || recipeColumn);
    const column = recipeColumn.trim();
    let step: QueryStep;
    try {
      if (recipeKind === 'trim-text') {
        if (columns.length === 0) throw new Error('Enter one or more columns, separated by commas');
        step = { id: `trim-${Date.now().toString(36)}`, kind: 'trim-text', name: `Trim ${columns.join(', ')}`, config: { columns }, enabled: true };
      } else if (recipeKind === 'split-column') {
        const outputs = parseColumnList(recipeOutputs);
        if (!column || !recipeValue) throw new Error('Split requires a source column and delimiter');
        if (outputs.length < 2) throw new Error('Split requires at least two output column names');
        step = { id: `split-${Date.now().toString(36)}`, kind: 'split-column', name: `Split ${column}`, config: { column, delimiter: recipeValue, outputColumns: outputs }, enabled: true };
      } else if (recipeKind === 'remove-duplicates') {
        if (columns.length === 0) throw new Error('Enter the duplicate key columns, separated by commas');
        step = { id: `dedupe-${Date.now().toString(36)}`, kind: 'remove-duplicates', name: `Remove duplicates by ${columns.join(', ')}`, config: { columns }, enabled: true };
      } else {
        if (!column) throw new Error('Sort requires a column');
        step = { id: `sort-${Date.now().toString(36)}`, kind: 'sort', name: `Sort ${column}`, config: { column, ascending: recipeAscending }, enabled: true };
      }
      setRecipeSteps((current) => [...current, step]);
      setRecipeError(null);
    } catch (error) {
      setRecipeError(error instanceof Error ? error.message : 'Invalid recipe step');
    }
  };

  const moveRecipeStep = (index: number, direction: -1 | 1) => {
    setRecipeSteps((current) => {
      const target = index + direction;
      if (target < 0 || target >= current.length) return current;
      const next = [...current];
      const [step] = next.splice(index, 1);
      if (!step) return current;
      next.splice(target, 0, step);
      return next;
    });
  };

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

          <Panel className="border border-slate-200 shadow-none">
            <PanelHeader className="border-b border-slate-200 px-3 py-2">
              <Stack gap="none">
                <PanelTitle size="sm">Cleaning recipe</PanelTitle>
                <Text size="xs" tone="muted">Steps are saved with the query and replayed on refresh.</Text>
              </Stack>
            </PanelHeader>
            <PanelBody className="space-y-2 p-3">
              <Select aria-label="Cleaning step type" value={recipeKind} onChange={(event) => setRecipeKind(event.target.value as RecipeKind)} sizeVariant="sm" disabled={!canQuery || busy}>
                <option value="trim-text">Trim text</option>
                <option value="split-column">Split column</option>
                <option value="remove-duplicates">Remove duplicates</option>
                <option value="sort">Sort rows</option>
              </Select>
              {recipeKind === 'split-column' ? (
                <Stack gap="xs">
                  <TextInput aria-label="Split source column" placeholder="Source column, e.g. Address" value={recipeColumn} onChange={(event) => setRecipeColumn(event.target.value)} disabled={!canQuery || busy} />
                  <Inline gap="xs"><TextInput aria-label="Split delimiter" className="min-w-0 flex-1" placeholder="Delimiter" value={recipeValue} onChange={(event) => setRecipeValue(event.target.value)} disabled={!canQuery || busy} /><TextInput aria-label="Split output columns" className="min-w-0 flex-[2]" placeholder="Output columns, e.g. City,State" value={recipeOutputs} onChange={(event) => setRecipeOutputs(event.target.value)} disabled={!canQuery || busy} /></Inline>
                </Stack>
              ) : recipeKind === 'sort' ? (
                <Stack gap="xs">
                  <TextInput aria-label="Sort column" placeholder="Column, e.g. Revenue" value={recipeColumn} onChange={(event) => setRecipeColumn(event.target.value)} disabled={!canQuery || busy} />
                  <CheckToggle label={recipeAscending ? 'Ascending' : 'Descending'} checked={recipeAscending} onChange={(event) => setRecipeAscending(event.currentTarget.checked)} disabled={!canQuery || busy} />
                </Stack>
              ) : (
                <TextInput aria-label="Cleaning columns" placeholder="Columns, comma separated" value={recipeColumns} onChange={(event) => setRecipeColumns(event.target.value)} disabled={!canQuery || busy} />
              )}
              <Button size="xs" variant="outline" disabled={!canQuery || busy} onClick={addRecipeStep}>Add step</Button>
              {recipeError ? <Text size="xs" tone="danger">{recipeError}</Text> : null}
              {recipeSteps.length > 0 ? (
                <Stack gap="xs">
                  {recipeSteps.map((step, index) => (
                    <Inline key={step.id} gap="xs" className="items-center rounded border border-slate-200 bg-slate-50 px-2 py-1">
                      <CheckToggle label="" aria-label={`Enable ${step.name}`} checked={step.enabled} onChange={(event) => setRecipeSteps((current) => current.map((entry) => entry.id === step.id ? { ...entry, enabled: event.currentTarget.checked } : entry))} />
                      <Text size="xs" className="min-w-0 flex-1 truncate">{index + 1}. {step.name}</Text>
                      <Button icon="chevron-up" iconOnly size="xs" variant="ghost" aria-label="Move step up" disabled={index === 0} onClick={() => moveRecipeStep(index, -1)} />
                      <Button icon="chevron-down" iconOnly size="xs" variant="ghost" aria-label="Move step down" disabled={index === recipeSteps.length - 1} onClick={() => moveRecipeStep(index, 1)} />
                      <Button icon="x" iconOnly size="xs" variant="ghost" aria-label="Remove step" onClick={() => setRecipeSteps((current) => current.filter((entry) => entry.id !== step.id))} />
                    </Inline>
                  ))}
                  <Button size="xs" variant="ghost" onClick={() => setRecipeSteps([])}>Clear recipe</Button>
                </Stack>
              ) : <Text size="xs" tone="muted">No cleaning steps yet.</Text>}
            </PanelBody>
          </Panel>

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
