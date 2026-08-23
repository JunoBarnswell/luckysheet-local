import { useState, type ReactNode } from 'react';
import { Box, Button, DropdownMenu, Heading, Icon, Inline, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { PeerCursor, SaveState, WorkspacePhase } from '../state/workspace';
import { localeLabels, translate, type Locale } from '../i18n';

export interface AppShellProps {
  children: ReactNode;
  formulaBar: ReactNode;
  isBusy: boolean;
  notice: string;
  locale: Locale;
  onLocaleChange: (locale: Locale) => void;
  onSearch?: (query: string) => void;
  onShare: () => void;
  peers: readonly PeerCursor[];
  workbookMenu?: ReactNode;
  ribbon: ReactNode;
  saveState: SaveState;
  sheetTabs: ReactNode;
  statusBar: ReactNode;
  title: string;
  workspacePhase: WorkspacePhase;
}

const peerColorClasses: Record<string, string> = {
  '#2563eb': 'bg-blue-600',
  '#10b981': 'bg-emerald-500',
  '#f59e0b': 'bg-amber-500',
  '#ef4444': 'bg-rose-500',
  '#8b5cf6': 'bg-violet-500',
  '#06b6d4': 'bg-cyan-500',
};

export function AppShell({ children, formulaBar, isBusy, locale, notice, onLocaleChange, onSearch, onShare, peers, ribbon, saveState, sheetTabs, statusBar, title, workbookMenu, workspacePhase }: AppShellProps) {
  const [search, setSearch] = useState('');
  const saveCopy = {
    label: translate(locale, saveState),
    tone: saveState === 'saved' ? 'text-emerald-300' : saveState === 'saving' ? 'text-amber-300' : saveState === 'offline' ? 'text-rose-300' : 'text-sky-300',
  };
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
              <Text size="xs" tone="subtle">{translate(locale, 'workspaceLabel')}</Text>
            </Stack>
          </Inline>
          <Box className="hidden h-8 w-px bg-slate-700 md:block" />
          <Stack gap="none" className="min-w-0">
            <Text size="xs" tone="subtle" weight="medium" className="uppercase tracking-[0.14em]">{translate(locale, 'planningWorkbook')}</Text>
            <Heading as="h1" size="md" className="truncate text-white">{title}</Heading>
          </Stack>
          <Inline gap="xs" className="hidden rounded-lg border border-slate-700 bg-slate-900 px-2.5 py-1.5 text-slate-300 lg:flex">
            <Text size="xs" tone="subtle">{translate(locale, 'personal')}</Text>
            <Icon name="chevron-down" size="xs" />
          </Inline>
        </Inline>

        <Inline gap="sm" className="shrink-0">
          <Box className="hidden w-64 lg:block">
            <TextInput
              aria-label={translate(locale, 'searchWorkbook')}
              placeholder={translate(locale, 'searchWorkbook')}
              leadingIcon="search"
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && search.trim()) onSearch?.(search.trim());
              }}
              className="border-slate-700 bg-slate-900 text-white placeholder:text-slate-500 focus:border-blue-400"
            />
          </Box>
          <Inline gap="xs" className="hidden md:flex">
            <Box className="h-1.5 w-1.5 rounded-full bg-emerald-400" />
            <Text size="xs" className={saveCopy.tone}>{saveCopy.label}</Text>
          </Inline>
          <Text size="xs" tone="subtle" className="hidden xl:inline">{notice}</Text>
          {peers.length > 0 ? (
            <Inline gap="none" className="hidden items-center md:flex">
              {peers.slice(0, 3).map((peer) => (
                <Box key={peer.actorId} title={peer.name} className={`-ml-1 flex h-7 w-7 items-center justify-center rounded-full border-2 border-slate-950 text-[10px] font-semibold text-white first:ml-0 ${peerColorClasses[peer.color] ?? 'bg-blue-600'}`}>
                  {peer.name.slice(0, 2).toUpperCase()}
                </Box>
              ))}
              {peers.length > 3 ? <Text size="xs" tone="subtle" className="ml-1">+{peers.length - 3}</Text> : null}
            </Inline>
          ) : null}
          <Button aria-label="Share workbook" disabled={isBusy} icon="share" onClick={onShare} size="sm" variant="soft">{translate(locale, 'share')}</Button>
          <DropdownMenu
            align="right"
            trigger={<Button aria-label={translate(locale, 'language')} disabled={isBusy} size="sm" variant="ghost" className="text-slate-300 hover:bg-slate-800 hover:text-white">{localeLabels[locale]}</Button>}
          >
            {({ close }) => (
              <Stack gap="xs" className="min-w-36">
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { onLocaleChange('en-US'); close(); }}>{translate(locale, 'english')}</Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={() => { onLocaleChange('zh-CN'); close(); }}>{translate(locale, 'simplifiedChinese')}</Button>
              </Stack>
            )}
          </DropdownMenu>
          {workbookMenu}
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
