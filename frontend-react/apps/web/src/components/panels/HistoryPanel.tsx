import React from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Stack, Text } from '@react-sheets/ui-system';
import type { HistoryEntry } from '@react-sheets/command-runtime';

export interface HistoryPanelProps {
  entries: readonly HistoryEntry[];
  onUndoTo?: (index: number) => void;
  onClose?: () => void;
}

export function HistoryPanel({ entries, onUndoTo, onClose }: HistoryPanelProps) {
  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Revision & Mutation Log</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Text size="xs" tone="subtle">
            All workbook state mutations are tracked as reversible transactions.
          </Text>

          {entries.length === 0 ? (
            <div className="rounded-lg border border-dashed border-slate-200 p-6 text-center text-xs text-slate-400">
              No edit history in this session yet.
            </div>
          ) : (
            <Stack gap="xs">
              {entries.map((entry, idx) => (
                <div
                  key={entry.operationId}
                  className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2.5 text-xs shadow-2xs"
                >
                  <div>
                    <div className="font-semibold text-slate-800">
                      {entry.description || 'Workbook Mutation'}
                    </div>
                    <div className="text-[10px] text-slate-400">
                      {new Date(entry.timestamp).toLocaleTimeString()} · {entry.undo.length} inverse ops
                    </div>
                  </div>
                  <span className="rounded bg-slate-100 px-1.5 py-0.5 text-[10px] font-mono text-slate-600">
                    #{idx + 1}
                  </span>
                </div>
              ))}
            </Stack>
          )}
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>
            Close Panel
          </Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
