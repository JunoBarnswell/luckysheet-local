import { useState } from 'react';
import { Box, Button, Dialog, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import type { WorkbookTableModel } from '@react-sheets/core-model';
import type { TableRowsResponse } from '@react-sheets/spreadsheet-app';

export interface DataModelPanelProps {
  tables: readonly WorkbookTableModel[];
  onReadRows: (tableId: string, offset?: number, limit?: number) => Promise<TableRowsResponse>;
  onRemove: (tableId: string) => Promise<void>;
}

export function DataModelPanel({ onReadRows, onRemove, tables }: DataModelPanelProps) {
  const [preview, setPreview] = useState<Record<string, TableRowsResponse>>({});
  const [loading, setLoading] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [deleteTableId, setDeleteTableId] = useState<string | null>(null);

  if (tables.length === 0) {
    return <StatePanel kind="empty" title="No data tables" description="Select a range and choose Data → Create Data Table to build a paged data source." />;
  }

  return (
    <Stack gap="md">
      {error ? <StatePanel kind="error" description={error} actionLabel="Dismiss" onAction={() => setError(null)} /> : null}
      {tables.map((table) => {
        const result = preview[table.id];
        return (
          <Panel key={table.id} className="shadow-none">
            <PanelHeader className="px-3 py-2">
              <Inline gap="sm" className="justify-between">
                <Stack gap="none" className="min-w-0">
                  <PanelTitle size="sm" className="truncate">{table.name}</PanelTitle>
                  <Text size="xs" tone="subtle">{table.rowCount.toLocaleString()} rows · revision {table.revision}</Text>
                </Stack>
              <Inline gap="xs">
                <Button size="xs" variant="outline" loading={loading === table.id} onClick={() => {
                  setLoading(table.id);
                  setError(null);
                  void onReadRows(table.id, 0, 20).then((response) => setPreview((current) => ({ ...current, [table.id]: response }))).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : 'Data table query failed')).finally(() => setLoading(null));
                }}>Preview rows</Button>
                {preview[table.id]?.nextOffset != null ? <Button size="xs" variant="ghost" loading={loading === table.id} onClick={() => {
                  const offset = preview[table.id]?.nextOffset;
                  if (offset == null) return;
                  setLoading(table.id);
                  void onReadRows(table.id, offset, 20).then((response) => setPreview((current) => {
                    const previous = current[table.id];
                    return { ...current, [table.id]: { ...response, rows: [...(previous?.rows ?? []), ...response.rows] } };
                  })).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : 'Data table query failed')).finally(() => setLoading(null));
                }}>Load more</Button> : null}
                <Button size="xs" variant="danger" icon="trash" onClick={() => setDeleteTableId(table.id)}>Delete</Button>
                </Inline>
              </Inline>
            </PanelHeader>
            <PanelBody className="p-3">
              <Inline gap="xs" className="flex-wrap">
                {table.fields.map((field) => <Text key={field.id} size="xs" className="rounded bg-slate-100 px-2 py-1">{field.name} · {field.type}</Text>)}
              </Inline>
              {result ? (
                <Box className="mt-3 overflow-auto rounded border border-slate-200">
                  <Stack gap="none" className="min-w-max">
                    <Inline gap="none" className="border-b border-slate-200 bg-slate-50">
                      {result.table.fields.map((field) => <Text key={field.id} size="xs" weight="semibold" className="w-32 border-r border-slate-200 px-2 py-1">{field.name}</Text>)}
                    </Inline>
                    {result.rows.map((row, rowIndex) => <Inline key={rowIndex} gap="none" className="border-b border-slate-100 last:border-0">{row.map((value, columnIndex) => <Text key={columnIndex} size="xs" className="w-32 truncate border-r border-slate-100 px-2 py-1">{value == null ? '' : String(value)}</Text>)}</Inline>)}
                  </Stack>
                </Box>
              ) : null}
            </PanelBody>
          </Panel>
        );
      })}
      <Dialog open={deleteTableId !== null} title="Delete data table" onClose={() => setDeleteTableId(null)} maxWidth="sm">
        <Text size="sm">Delete {tables.find((table) => table.id === deleteTableId)?.name ?? 'this data table'}? The operation can be undone.</Text>
        <Inline gap="sm" className="mt-4 justify-end">
          <Button size="sm" variant="ghost" onClick={() => setDeleteTableId(null)}>Cancel</Button>
          <Button size="sm" variant="danger" onClick={() => {
            if (!deleteTableId) return;
            void onRemove(deleteTableId).catch((reason: unknown) => setError(reason instanceof Error ? reason.message : 'Data table deletion failed')).finally(() => setDeleteTableId(null));
          }}>Delete</Button>
        </Inline>
      </Dialog>
    </Stack>
  );
}
