import { Button, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import type { DataSourceManifest, SheetDataRegion } from '@react-sheets/core-model';

export interface DataSourcePanelProps {
  sources: readonly DataSourceManifest[];
  regions: readonly SheetDataRegion[];
  onRemoveSource: (sourceId: string) => void;
  onRemoveRegion: (regionId: string) => void;
}

/**
 * Data tab's canonical binding surface.  It only consumes the manifest and
 * sheet-region projections; block bytes remain owned by the data-block store.
 */
export function DataSourcePanel({ onRemoveRegion, onRemoveSource, regions, sources }: DataSourcePanelProps) {
  if (sources.length === 0 && regions.length === 0) {
    return <StatePanel kind="empty" title="No data sources" description="Create a binding from a headered range to add a Data Source." />;
  }

  return (
    <Stack gap="md">
      <Panel className="shadow-none">
        <PanelHeader className="px-3 py-2">
          <PanelTitle size="sm">Data Source</PanelTitle>
        </PanelHeader>
        <PanelBody className="p-3">
          <Stack gap="sm">
            <Text size="xs" tone="subtle">Load Schema and Save Schema operate on the canonical DataSourceManifest boundary. Binary blocks stay outside workbook snapshots, operations, and WebSocket payloads.</Text>
            {sources.map((source) => (
              <Panel key={source.id} className="border border-slate-200 shadow-none">
                <PanelBody className="p-3">
                  <Inline gap="sm" className="items-start justify-between">
                    <Stack gap="none" className="min-w-0">
                      <Text size="sm" weight="semibold" className="truncate">{source.name}</Text>
                      <Text size="xs" tone="subtle">{source.kind} · {source.rowCount.toLocaleString()} rows · {source.fields.length} fields · revision {source.revision}</Text>
                    </Stack>
                    <Button size="xs" variant="danger" icon="trash" onClick={() => onRemoveSource(source.id)}>Remove</Button>
                  </Inline>
                  <Inline gap="xs" className="mt-2 flex-wrap">
                    {source.fields.map((field) => <Text key={field.id} size="xs" className="rounded bg-slate-100 px-2 py-1">{field.name} · {field.type}</Text>)}
                  </Inline>
                </PanelBody>
              </Panel>
            ))}
          </Stack>
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader className="px-3 py-2">
          <PanelTitle size="sm">Table Binding Source</PanelTitle>
        </PanelHeader>
        <PanelBody className="p-3">
          {regions.length === 0 ? <Text size="xs" tone="subtle">No sheet binding regions on the active sheet.</Text> : (
            <Stack gap="xs">
              {regions.map((region) => (
                <Inline key={region.id} gap="sm" className="items-center justify-between rounded border border-slate-200 px-2 py-2">
                  <Stack gap="none" className="min-w-0">
                    <Text size="xs" weight="semibold" className="truncate">{region.id}</Text>
                    <Text size="xs" tone="subtle">source {region.sourceId} · header row {region.headerRow + 1} · revision {region.revision}</Text>
                  </Stack>
                  <Button size="xs" variant="ghost" onClick={() => onRemoveRegion(region.id)}>Unbind</Button>
                </Inline>
              ))}
            </Stack>
          )}
        </PanelBody>
      </Panel>
    </Stack>
  );
}
