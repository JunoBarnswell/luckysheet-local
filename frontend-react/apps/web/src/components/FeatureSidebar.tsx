import { Box, Button, Heading, Icon, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Stack, StatePanel, Tab, TabList, Tabs, Text } from '@react-sheets/ui-system';
import type { SheetView, SidebarPanelId, WorkspacePhase } from '../state/workspace';

export interface FeatureSidebarProps {
  activePanel: SidebarPanelId;
  activeCell: string;
  onPanelChange: (panel: SidebarPanelId) => void;
  onRetry: () => void;
  phase: WorkspacePhase;
  sheet: SheetView;
}

const panels: Array<{ icon: React.ComponentProps<typeof Icon>['name']; id: SidebarPanelId; label: string }> = [
  { id: 'inspector', label: 'Inspect', icon: 'sliders' },
  { id: 'data', label: 'Data', icon: 'table' },
  { id: 'comments', label: 'Notes', icon: 'comment' },
  { id: 'automations', label: 'Flow', icon: 'sparkles' },
];

interface InsightRowProps {
  label: string;
  value: string;
  tone?: 'accent' | 'muted' | 'success';
}

function InsightRow({ label, tone = 'muted', value }: InsightRowProps) {
  return (
    <Inline gap="sm" className="justify-between border-b border-line/70 py-2.5 last:border-0">
      <Text size="xs" tone="muted">{label}</Text>
      <Text size="xs" tone={tone} weight="semibold">{value}</Text>
    </Inline>
  );
}

function InspectorPanel({ activeCell, sheet }: { activeCell: string; sheet: SheetView }) {
  const cells = sheet.rows.flatMap((row) => row.cells).filter((cell) => cell.value !== '');
  const selected = cells.find((cell) => cell.address === activeCell);
  const numericValues = cells.map((cell) => Number(cell.value.replace(/[$,%]/g, ''))).filter((value) => Number.isFinite(value));
  const average = numericValues.length > 0 ? Math.round(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length).toLocaleString('en-US') : '—';
  return (
    <Stack gap="md">
      <Panel tone="accent" className="overflow-hidden shadow-none">
        <PanelBody>
          <Inline gap="sm" className="mb-3">
            <Box className="flex h-9 w-9 items-center justify-center rounded-xl bg-white text-accent shadow-sm">
              <Icon name="sparkles" size="md" />
            </Box>
            <Stack gap="none">
              <PanelTitle as="h3" size="sm">Quick insight</PanelTitle>
              <Text size="xs" tone="muted">Selected cell · {activeCell}</Text>
            </Stack>
          </Inline>
          <Text size="sm" tone="default" weight="medium">{selected?.value || 'Empty cell selected'}</Text>
          <Inline gap="xs" className="mt-3">
            <Text size="xs" tone="accent" weight="bold">{selected?.formula ? 'Formula' : 'Value'}</Text>
            <Text size="xs" tone="muted">from WorkbookModel</Text>
          </Inline>
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader>
          <Inline gap="sm">
            <Icon name="chart" size="sm" className="text-accent" />
            <PanelTitle as="h3" size="sm">Range summary</PanelTitle>
          </Inline>
          <Button aria-label="Range summary options" icon="more-horizontal" iconOnly size="sm" variant="ghost" />
        </PanelHeader>
        <PanelBody className="py-2">
          <InsightRow label="Numeric average" value={average} tone="accent" />
          <InsightRow label="Cells with values" value={String(cells.length)} />
          <InsightRow label="Active sheet" value={sheet.name} />
        </PanelBody>
      </Panel>

      <Panel tone="subtle" className="shadow-none">
        <PanelBody>
          <Inline gap="sm" className="mb-2">
            <Icon name="lock" size="sm" className="text-muted" />
            <Text size="xs" weight="semibold">Protected surface</Text>
          </Inline>
          <Text size="xs" tone="muted">Locking, permissions and audit history will attach to this surface through the runtime.</Text>
        </PanelBody>
      </Panel>
    </Stack>
  );
}

function DataPanel() {
  return (
    <Stack gap="md">
      <Stack gap="xs">
        <Heading as="h3" size="sm">Data tools</Heading>
        <Text size="xs" tone="muted">Keep the working set tidy while the grid is in preview mode.</Text>
      </Stack>
      <StatePanel kind="disabled" title="Feature command not registered" description="Sort, filter and export will be enabled when their typed commands are registered in the Sheet feature package." actionLabel="Command registry" actionDisabled />
    </Stack>
  );
}

function CommentsPanel() {
  return <StatePanel kind="empty" title="No notes on this sheet" description="Comments will appear here when collaborators leave feedback on a range." actionLabel="Add a note" />;
}

function AutomationsPanel() {
  return <StatePanel kind="disabled" title="Flows are not connected" description="Automation actions become available when a command runtime is attached." actionLabel="Learn about flows" actionDisabled />;
}

export function FeatureSidebar({ activePanel, activeCell, onPanelChange, onRetry, phase, sheet }: FeatureSidebarProps) {
  const disabled = phase !== 'ready';
  const activePanelLabel = panels.find((panel) => panel.id === activePanel)?.label ?? 'Inspect';
  return (
    <Box as="aside" aria-label="Feature sidebar" className="hidden w-[304px] shrink-0 flex-col border-l border-line bg-slate-50/65 lg:flex">
      <Tabs className="shrink-0 border-b border-line bg-white px-3 pt-3">
        <TabList label="Feature panels" className="grid grid-cols-4 gap-1">
          {panels.map((panel) => (
            <Tab key={panel.id} active={panel.id === activePanel} disabled={disabled} onClick={() => onPanelChange(panel.id)} className="flex-col gap-1 px-1 py-2">
              <Icon name={panel.icon} size="sm" />
              <Text size="xs" weight="medium">{panel.label}</Text>
            </Tab>
          ))}
        </TabList>
      </Tabs>

      <Box className="min-h-0 flex-1 overflow-auto p-4">
        <Inline gap="sm" className="mb-4">
          <Stack gap="none" className="min-w-0">
            <Text size="xs" tone="subtle" weight="bold" className="uppercase tracking-[0.14em]">Feature panel</Text>
            <Heading as="h2" size="md">{activePanelLabel}</Heading>
          </Stack>
          <Button aria-label="Sidebar settings" disabled={disabled} icon="settings" iconOnly size="sm" variant="ghost" className="ml-auto" />
        </Inline>

        {phase === 'loading' ? <StatePanel kind="loading" description="Preparing panel data." /> : null}
        {phase === 'error' ? <StatePanel actionLabel="Try again" description="Panel data could not be loaded." kind="error" onAction={onRetry} /> : null}
        {phase === 'empty' ? <StatePanel actionLabel="Try again" description="Open a workbook to inspect its ranges." kind="empty" onAction={onRetry} /> : null}
        {phase === 'ready' && activePanel === 'inspector' ? <InspectorPanel activeCell={activeCell} sheet={sheet} /> : null}
        {phase === 'ready' && activePanel === 'data' ? <DataPanel /> : null}
        {phase === 'ready' && activePanel === 'comments' ? <CommentsPanel /> : null}
        {phase === 'ready' && activePanel === 'automations' ? <AutomationsPanel /> : null}
      </Box>

      <Box className="shrink-0 border-t border-line bg-white px-4 py-3">
        <Inline gap="sm">
          <Box className="flex h-7 w-7 items-center justify-center rounded-lg bg-slate-100 text-muted">
            <Icon name="help" size="sm" />
          </Box>
          <Stack gap="none" className="min-w-0">
            <Text size="xs" weight="semibold">Need a hand?</Text>
            <Text size="xs" tone="muted">Open the shortcuts for quick navigation.</Text>
          </Stack>
          <Icon name="chevron-right" size="sm" className="ml-auto text-slate-400" />
        </Inline>
      </Box>
    </Box>
  );
}
