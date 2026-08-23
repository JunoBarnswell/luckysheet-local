import React, { useEffect, useState } from 'react';
import { Box, Button, Inline, Stack, TextInput, Text } from '@react-sheets/ui-system';

export interface FindReplaceDialogProps {
  open: boolean;
  initialFind?: string;
  onClose: () => void;
  onReplaceAll: (params: { find: string; replace: string; matchCase: boolean; entireCell: boolean; scope: 'sheet' | 'workbook' }) => number;
}

/** 查找替换对话框:返回替换命中数由调用方展示 */
export function FindReplaceDialog({ initialFind = '', open, onClose, onReplaceAll }: FindReplaceDialogProps): React.ReactElement | null {
  const [find, setFind] = useState('');
  const [replace, setReplace] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [entireCell, setEntireCell] = useState(false);
  const [scope, setScope] = useState<'sheet' | 'workbook'>('sheet');
  const [result, setResult] = useState<number | null>(null);

  useEffect(() => {
    if (!open) return;
    setFind(initialFind);
    setResult(null);
  }, [initialFind, open]);

  if (!open) return null;

  return (
    <Box className="fixed inset-0 z-50 flex items-start justify-center bg-slate-900/30 pt-24">
      <Box className="w-96 rounded-xl border border-slate-200 bg-white p-4 shadow-2xl">
        <Stack gap="sm">
          <Text size="sm" weight="semibold">Find and replace</Text>
          <TextInput aria-label="Find" placeholder="Find" value={find} onChange={(event) => setFind(event.target.value)} />
          <TextInput
            aria-label="Replace with"
            placeholder="Replace with"
            value={replace}
            onChange={(event) => setReplace(event.target.value)}
          />
          <Inline gap="md" className="text-xs text-slate-600">
            <label className="flex items-center gap-1.5">
              <input checked={matchCase} type="checkbox" onChange={(event) => setMatchCase(event.target.checked)} />
              Match case
            </label>
            <label className="flex items-center gap-1.5">
              <input checked={entireCell} type="checkbox" onChange={(event) => setEntireCell(event.target.checked)} />
              Entire cell
            </label>
            <select
              aria-label="Scope"
              className="rounded border border-slate-200 px-1 py-0.5"
              value={scope}
              onChange={(event) => setScope(event.target.value as 'sheet' | 'workbook')}
            >
              <option value="sheet">This sheet</option>
              <option value="workbook">All sheets</option>
            </select>
          </Inline>
          {result !== null ? (
            <Text size="xs" tone="subtle">{result} cell(s) replaced</Text>
          ) : null}
          <Inline gap="sm" className="justify-end">
            <Button size="sm" variant="ghost" onClick={onClose}>Close</Button>
            <Button
              size="sm"
              variant="primary"
              disabled={!find}
              onClick={() => {
                const count = onReplaceAll({ find, replace, matchCase, entireCell, scope });
                setResult(count);
              }}
            >
              Replace all
            </Button>
          </Inline>
        </Stack>
      </Box>
    </Box>
  );
}
