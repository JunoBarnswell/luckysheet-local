import type { ReactNode } from 'react';
import { Box, Button, Heading, Icon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { SaveState, WorkspacePhase } from '../state/workspace';

export interface AppShellProps {
  children: ReactNode;
  formulaBar: ReactNode;
  isBusy: boolean;
  notice: string;
  onMenu: () => void;
  onShare: () => void;
  ribbon: ReactNode;
  saveState: SaveState;
  sheetTabs: ReactNode;
  statusBar: ReactNode;
  title: string;
  workspacePhase: WorkspacePhase;
}

const saveStateCopy: Record<SaveState, { label: string; tone: string }> = {
  saved: { label: 'Saved', tone: 'text-emerald-300' },
  saving: { label: 'Saving', tone: 'text-amber-300' },
  offline: { label: 'Offline', tone: 'text-rose-300' },
  syncing: { label: 'Syncing', tone: 'text-sky-300' },
};

export function AppShell({ children, formulaBar, isBusy, notice, onMenu, onShare, ribbon, saveState, sheetTabs, statusBar, title, workspacePhase }: AppShellProps) {
  const saveCopy = saveStateCopy[saveState];
  return (
    <Box
      aria-busy={isBusy}
      as="main"
      className="flex h-screen min-h-[680px] min-w-[960px] flex-col overflow-hidden bg-canvas text-ink"
      data-workspace-phase={workspacePhase}
    >
      <Box as="header" className="flex h-16 shrink-0 items-center justify-between border-b border-slate-800 bg-slate-950 px-5 text-white">
        <Inline gap="lg" className="min-w-0">
          <Inline gap="sm" className="shrink-0">
            <Box className="flex h-9 w-9 items-center justify-center rounded-xl bg-accent text-white shadow-lg shadow-accent/30">
              <Icon name="grid" size="lg" />
            </Box>
            <Stack gap="none" className="hidden sm:flex">
              <Text size="xs" tone="inverse" weight="bold" className="tracking-[0.18em] text-slate-300">SHEETS</Text>
              <Text size="xs" tone="subtle">Workspace</Text>
            </Stack>
          </Inline>
          <Box className="hidden h-8 w-px bg-slate-700 md:block" />
          <Stack gap="none" className="min-w-0">
            <Text size="xs" tone="subtle" weight="medium" className="uppercase tracking-[0.14em]">Planning workbook</Text>
            <Heading as="h1" size="md" className="truncate text-white">{title}</Heading>
          </Stack>
          <Inline gap="xs" className="hidden rounded-lg border border-slate-700 bg-slate-900 px-2.5 py-1.5 text-slate-300 lg:flex">
            <Text size="xs" tone="subtle">Personal</Text>
            <Icon name="chevron-down" size="xs" />
          </Inline>
        </Inline>

        <Inline gap="sm" className="shrink-0">
          <Inline gap="xs" className="hidden md:flex">
            <Box className="h-1.5 w-1.5 rounded-full bg-emerald-400" />
            <Text size="xs" className={saveCopy.tone}>{saveCopy.label}</Text>
          </Inline>
          <Text size="xs" tone="subtle" className="hidden xl:inline">{notice}</Text>
          <Button aria-label="Share workbook" disabled={isBusy} icon="share" onClick={onShare} size="sm" variant="soft">Share</Button>
          <Button aria-label="Open workbook menu" disabled={isBusy} icon="more-horizontal" iconOnly onClick={onMenu} size="sm" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white" />
        </Inline>
      </Box>

      <Box as="section" className="z-10 shrink-0 border-b border-line bg-white shadow-sm">
        {ribbon}
        {formulaBar}
      </Box>

      <Box as="section" className="flex min-h-0 flex-1 overflow-hidden">
        {children}
      </Box>

      <Box as="section" className="z-10 shrink-0 border-t border-line bg-white">
        {sheetTabs}
      </Box>

      <Box as="footer" className="z-10 shrink-0 border-t border-slate-800 bg-slate-950 text-white">
        {statusBar}
      </Box>
    </Box>
  );
}
