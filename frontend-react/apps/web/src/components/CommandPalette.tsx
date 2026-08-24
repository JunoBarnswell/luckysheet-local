import { useEffect, useMemo, useState } from 'react';
import { Box, Button, Dialog, Stack, Text, TextInput } from '@react-sheets/ui-system';

export interface CommandPaletteEntry {
  id: string;
  label: string;
  group: string;
  keywords?: readonly string[];
  tip?: string;
  commandId?: string;
  enabled: boolean;
  execute: () => void;
}

export interface CommandPaletteProps {
  open: boolean;
  commands: readonly CommandPaletteEntry[];
  onClose: () => void;
}

const RECENT_COMMANDS_KEY = 'react-sheets:command-palette:recent';

function fuzzyMatch(value: string, needle: string): boolean {
  let position = 0;
  for (const character of needle) {
    position = value.indexOf(character, position);
    if (position < 0) return false;
    position += 1;
  }
  return true;
}

function score(entry: CommandPaletteEntry, query: string): number {
  const needle = query.trim().toLocaleLowerCase();
  if (!needle) return 0;
  const label = entry.label.toLocaleLowerCase();
  const keywords = (entry.keywords ?? []).map((keyword) => keyword.toLocaleLowerCase());
  const searchable = [label, entry.group.toLocaleLowerCase(), ...(entry.tip ? [entry.tip.toLocaleLowerCase()] : []), ...(entry.commandId ? [entry.commandId.toLocaleLowerCase()] : []), ...keywords];
  if (label === needle) return 400;
  if (label.startsWith(needle)) return 300;
  if (keywords.some((keyword) => keyword === needle)) return 250;
  if (keywords.some((keyword) => keyword.startsWith(needle))) return 200;
  if (searchable.some((value) => value.includes(needle))) return 100;
  if (searchable.some((value) => fuzzyMatch(value, needle))) return 50;
  return -1;
}

export function CommandPalette({ open, commands, onClose }: CommandPaletteProps) {
  const [query, setQuery] = useState('');
  const [activeIndex, setActiveIndex] = useState(0);
  const [recentIds, setRecentIds] = useState<string[]>([]);
  const results = useMemo(() => commands
    .map((entry, index) => ({ entry, index, score: score(entry, query), recent: recentIds.indexOf(entry.id) }))
    .filter((entry) => entry.score >= 0 || !query.trim())
    .sort((left, right) => right.score - left.score || (left.recent < 0 ? 1 : right.recent < 0 ? -1 : left.recent - right.recent) || left.entry.label.localeCompare(right.entry.label)), [commands, query, recentIds]);

  useEffect(() => {
    if (!open) return;
    setQuery('');
    setActiveIndex(0);
    try {
      const stored = JSON.parse(window.localStorage.getItem(RECENT_COMMANDS_KEY) ?? '[]') as unknown;
      if (Array.isArray(stored)) setRecentIds(stored.filter((id): id is string => typeof id === 'string').slice(0, 12));
    } catch {
      setRecentIds([]);
    }
  }, [open]);

  const execute = (entry: CommandPaletteEntry): void => {
    const next = [entry.id, ...recentIds.filter((id) => id !== entry.id)].slice(0, 12);
    setRecentIds(next);
    try { window.localStorage.setItem(RECENT_COMMANDS_KEY, JSON.stringify(next)); } catch { /* preference storage is optional */ }
    entry.execute();
    onClose();
  };

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'ArrowDown') {
        event.preventDefault();
        setActiveIndex((value) => Math.min(Math.max(0, results.length - 1), value + 1));
      } else if (event.key === 'ArrowUp') {
        event.preventDefault();
        setActiveIndex((value) => Math.max(0, value - 1));
      } else if (event.key === 'Enter') {
        event.preventDefault();
        const current = results[activeIndex]?.entry;
        if (current?.enabled) {
          execute(current);
        }
      }
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [activeIndex, onClose, open, results]);

  return (
    <Dialog open={open} title="Command Palette" description="Search commands by name, keyword or context." onClose={onClose} maxWidth="md" testId="command-palette" bodyClassName="p-0">
      <Stack gap="sm" className="p-4">
        <TextInput aria-label="Command search" autoFocus leadingIcon="search" placeholder="Search commands" value={query} onChange={(event) => { setQuery(event.target.value); setActiveIndex(0); }} />
        <Box className="max-h-[360px] overflow-y-auto rounded-lg border border-slate-200 p-1">
          {results.length === 0 ? <Text size="sm" tone="muted" className="block px-3 py-6 text-center">No matching commands</Text> : null}
          {results.map(({ entry }, index) => (
            <Button
              key={entry.id}
              aria-disabled={!entry.enabled}
              disabled={!entry.enabled}
              onClick={() => execute(entry)}
              size="sm"
              variant={index === activeIndex ? 'soft' : 'ghost'}
              className="h-10 w-full justify-between px-3 text-left"
            >
              <Text as="span" className="min-w-0 truncate text-left">{entry.label}</Text>
              <Text as="span" size="xs" tone="subtle" className="ml-3 shrink-0">{entry.group}</Text>
            </Button>
          ))}
        </Box>
      </Stack>
    </Dialog>
  );
}
