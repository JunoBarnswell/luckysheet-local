import { useCallback, useEffect, useState } from 'react';
import { Box, Button, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Stack, StatePanel, Text } from '@react-sheets/ui-system';
import { WorkbookModel } from '@react-sheets/core-model';
import { getLocalWorkspaceStore, type LocalWorkspaceSummary } from '@react-sheets/spreadsheet-app';

const workspaceStore = getLocalWorkspaceStore();

function createWorkspaceId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return `wb-${crypto.randomUUID()}`;
  return `wb-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

export function WorkbookCatalog() {
  const [workbooks, setWorkbooks] = useState<LocalWorkspaceSummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string>();
  const [creating, setCreating] = useState(false);
  const [deleting, setDeleting] = useState<string | null>(null);
  const load = useCallback(async () => {
    setLoading(true);
    setError(undefined);
    try {
      setWorkbooks(await workspaceStore.list());
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not load workbooks');
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => { void load(); }, [load]);

  const create = async () => {
    setCreating(true);
    setError(undefined);
    try {
      const unitId = createWorkspaceId();
      const snapshot = new WorkbookModel(unitId, 'Untitled workbook').snapshot();
      const created = await workspaceStore.create({
        unitId,
        snapshot,
        localRevision: 0,
        serverRevision: 0,
        syncMode: 'local-only',
        operations: [],
        nextClientSequence: 0,
      });
      window.location.assign(`/workbooks/${encodeURIComponent(created.unitId)}`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not create workbook');
      setCreating(false);
    }
  };

  const remove = async (unitId: string) => {
    if (!window.confirm('Delete this workbook and its revisions?')) return;
    setDeleting(unitId);
    setError(undefined);
    try {
      await workspaceStore.delete(unitId);
      await load();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : 'Could not delete workbook');
    } finally {
      setDeleting(null);
    }
  };

  return (
    <Box as="main" className="min-h-screen bg-canvas p-8 text-ink">
      <Box className="mx-auto max-w-4xl">
        <Inline gap="lg" className="mb-6 items-end justify-between"><Stack gap="xs"><Text size="xs" tone="subtle" weight="bold" className="uppercase tracking-[0.18em]">SHEETS</Text><PanelTitle as="h1" size="lg">Workbooks</PanelTitle><Text size="sm" tone="muted">Open a real workbook or create a new one.</Text></Stack><Button disabled={creating} loading={creating} variant="primary" onClick={() => void create()}>New workbook</Button></Inline>
        <Panel>
          <PanelHeader><PanelTitle size="sm">Recent workbooks</PanelTitle><Button disabled={loading} size="xs" variant="ghost" onClick={() => void load()}>Refresh</Button></PanelHeader>
          <PanelBody>
            {loading ? <StatePanel kind="loading" description="Loading workbook resources." /> : null}
            {error ? <StatePanel kind="error" title="Workbook resources unavailable" description={error} actionLabel="Retry" onAction={() => void load()} /> : null}
            {!loading && !error && workbooks.length === 0 ? <StatePanel kind="empty" description="No workbooks have been created yet." actionLabel="Create workbook" onAction={() => void create()} /> : null}
            {!loading && !error && workbooks.length > 0 ? <Stack gap="xs">{workbooks.map((workbook) => <Inline key={workbook.unitId} gap="md" className="items-center justify-between rounded-lg border border-line bg-white px-4 py-3"><Stack gap="none" className="min-w-0"><Text size="sm" weight="semibold" className="truncate">{workbook.name}</Text><Text size="xs" tone="subtle">Revision {Math.max(workbook.localRevision, workbook.serverRevision)} · {new Date(workbook.updatedAt).toLocaleString()}</Text></Stack><Inline gap="xs"><Button size="sm" variant="secondary" onClick={() => window.location.assign(`/workbooks/${encodeURIComponent(workbook.unitId)}`)}>Open</Button><Button size="sm" variant="danger" loading={deleting === workbook.unitId} onClick={() => void remove(workbook.unitId)}>Delete</Button></Inline></Inline>)}</Stack> : null}
          </PanelBody>
        </Panel>
      </Box>
    </Box>
  );
}
