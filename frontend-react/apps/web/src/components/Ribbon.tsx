import { Box, Button, Divider, Icon, Inline, Stack, Tab, TabList, Tabs, Text } from '@react-sheets/ui-system';
import type { RibbonTabId, WorkspacePhase } from '../state/workspace';

export interface RibbonProps {
  activeTab: RibbonTabId;
  onAction: (action: string) => void;
  onTabChange: (tab: RibbonTabId) => void;
  phase: WorkspacePhase;
}

const ribbonTabs: Array<{ id: RibbonTabId; label: string }> = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
];

interface RibbonGroupProps {
  children: React.ReactNode;
  label: string;
}

function RibbonGroup({ children, label }: RibbonGroupProps) {
  return (
    <Stack gap="xs" className="shrink-0">
      <Inline gap="xs" className="min-h-8">{children}</Inline>
      <Text size="xs" tone="subtle" className="text-center uppercase tracking-[0.12em]">{label}</Text>
    </Stack>
  );
}

interface ToolButtonProps {
  action: string;
  disabled: boolean;
  icon: React.ComponentProps<typeof Icon>['name'];
  label: string;
  onAction: (action: string) => void;
}

function ToolButton({ action, disabled, icon, label, onAction }: ToolButtonProps) {
  return <Button aria-label={label} disabled={disabled} icon={icon} iconOnly onClick={() => onAction(action)} size="sm" variant="ghost" />;
}

export function Ribbon({ activeTab, onAction, onTabChange, phase }: RibbonProps) {
  const disabled = phase !== 'ready';
  return (
    <Tabs className="border-b border-line/80">
      <Inline gap="lg" className="h-11 overflow-x-auto px-5">
        <Text size="xs" tone="subtle" weight="bold" className="shrink-0 uppercase tracking-[0.16em]">Tools</Text>
        <TabList label="Workbook ribbon tabs" className="h-full gap-1">
          {ribbonTabs.map((tab) => (
            <Tab key={tab.id} active={activeTab === tab.id} disabled={disabled} onClick={() => onTabChange(tab.id)}>
              {tab.label}
            </Tab>
          ))}
        </TabList>
        <Inline gap="xs" className="ml-auto shrink-0 border-l border-line pl-4">
          <Icon name="cloud-check" size="sm" className={disabled ? 'text-slate-300' : 'text-emerald-500'} />
          <Text size="xs" tone="muted">Local sync ready</Text>
        </Inline>
      </Inline>

      <Box className="overflow-x-auto bg-slate-50/70 px-5 py-3">
        <Inline gap="lg" className="min-w-max items-start">
          <RibbonGroup label="History">
            <ToolButton action="undo" disabled={disabled} icon="undo" label="Undo" onAction={onAction} />
            <ToolButton action="redo" disabled={disabled} icon="redo" label="Redo" onAction={onAction} />
          </RibbonGroup>
          <Divider orientation="vertical" className="mt-1" />
          <RibbonGroup label="Text">
            <ToolButton action="bold" disabled={disabled} icon="bold" label="Bold" onAction={onAction} />
            <ToolButton action="italic" disabled={disabled} icon="italic" label="Italic" onAction={onAction} />
            <ToolButton action="underline" disabled={disabled} icon="underline" label="Underline" onAction={onAction} />
          </RibbonGroup>
          <Divider orientation="vertical" className="mt-1" />
          <RibbonGroup label="Align">
            <ToolButton action="align-left" disabled={disabled} icon="align-left" label="Align left" onAction={onAction} />
            <ToolButton action="align-center" disabled={disabled} icon="align-center" label="Align center" onAction={onAction} />
            <ToolButton action="align-right" disabled={disabled} icon="align-right" label="Align right" onAction={onAction} />
          </RibbonGroup>
          <Divider orientation="vertical" className="mt-1" />
          <RibbonGroup label="Insert">
            <ToolButton action="table" disabled={disabled} icon="table" label="Insert table" onAction={onAction} />
            <ToolButton action="chart" disabled={disabled} icon="chart" label="Insert chart" onAction={onAction} />
            <ToolButton action="comment" disabled={disabled} icon="comment" label="Add comment" onAction={onAction} />
          </RibbonGroup>
          <Divider orientation="vertical" className="mt-1" />
          <RibbonGroup label="Data">
            <ToolButton action="sort" disabled={disabled} icon="sort" label="Sort range" onAction={onAction} />
            <ToolButton action="filter" disabled={disabled} icon="filter" label="Create filter" onAction={onAction} />
            <ToolButton action="freeze" disabled={disabled} icon="freeze" label="Freeze panes" onAction={onAction} />
          </RibbonGroup>
          <Divider orientation="vertical" className="mt-1" />
          <RibbonGroup label="View">
            <ToolButton action="view" disabled={disabled} icon="eye" label="View options" onAction={onAction} />
            <ToolButton action="settings" disabled={disabled} icon="settings" label="Workbook settings" onAction={onAction} />
          </RibbonGroup>
        </Inline>
      </Box>
    </Tabs>
  );
}
