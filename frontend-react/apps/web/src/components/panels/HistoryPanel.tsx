import React from 'react';
import { Box, Button, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Stack, Text } from '@react-sheets/ui-system';
import type { HistoryEntry } from '@react-sheets/command-runtime';
import type { RevisionRecord } from '@react-sheets/protocol';

export interface HistoryPanelProps {
  entries: readonly HistoryEntry[];
  remoteRevisions?: readonly RevisionRecord[];
  previewRevision?: number | null;
  canRestore?: boolean;
  onUndoTo?: (index: number) => void;
  onRestoreRevision?: (revision: number) => void;
  onPreviewRevision?: (revision: number) => void;
  onClearPreview?: () => void;
  onRefreshRevisions?: () => void;
  onClose?: () => void;
}

export function HistoryPanel({
  entries,
  remoteRevisions = [],
  previewRevision = null,
  canRestore = true,
  onUndoTo,
  onRestoreRevision,
  onPreviewRevision,
  onClearPreview,
  onRefreshRevisions,
  onClose,
}: HistoryPanelProps) {
  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <Inline gap="sm" className="w-full items-center justify-between">
          <PanelTitle size="sm">Revision & Mutation Log</PanelTitle>
          {onRefreshRevisions ? (
            <Button variant="ghost" size="sm" onClick={onRefreshRevisions}>
              Refresh
            </Button>
          ) : null}
        </Inline>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Text size="xs" tone="subtle">
            All workbook state mutations are tracked as reversible transactions.
          </Text>

          {previewRevision != null ? (
            <Inline gap="sm" className="items-center rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-900">
              <Text size="xs" weight="semibold">Previewing revision #{previewRevision}</Text>
              {onClearPreview ? (
                <Button variant="ghost" size="sm" onClick={onClearPreview}>
                  Exit preview
                </Button>
              ) : null}
            </Inline>
          ) : null}

          {remoteRevisions.length > 0 ? (
            <Stack gap="xs">
              <Text size="xs" weight="semibold">Server revisions ({remoteRevisions.length})</Text>
              {remoteRevisions.slice(0, 20).map((revision) => (
                <Inline
                  key={`${revision.operationId}-${revision.revision}`}
                  gap="sm"
                  className="items-start justify-between rounded-lg border border-blue-100 bg-blue-50/50 p-2 text-xs"
                >
                  <Stack gap="none" className="min-w-0 flex-1">
                    <Inline gap="sm" className="items-center">
                      <Text size="xs" weight="semibold" className="text-blue-700">#{revision.revision}</Text>
                      {previewRevision === revision.revision ? (
                        <Text size="xs" className="text-amber-700">Previewing</Text>
                      ) : null}
                    </Inline>
                    <Text size="xs" className="truncate">{revision.payload.mutations.length} mutation(s)</Text>
                    <Text size="xs" tone="subtle" className="truncate">
                      {revision.payload.mutations.slice(0, 3).map((mutation) => mutation.id).join(' · ') || 'Workbook metadata'}
                    </Text>
                    <Text size="xs" tone="subtle">{new Date(revision.createdAt).toLocaleString()}</Text>
                  </Stack>
                  <Stack gap="xs" className="shrink-0">
                    {onPreviewRevision ? (
                      <Button variant="ghost" size="sm" onClick={() => onPreviewRevision(revision.revision)}>
                        Preview
                      </Button>
                    ) : null}
                    {onRestoreRevision && canRestore ? (
                      <Button variant="secondary" size="sm" onClick={() => onRestoreRevision(revision.revision)}>
                        Restore
                      </Button>
                    ) : null}
                  </Stack>
                </Inline>
              ))}
            </Stack>
          ) : null}

          {entries.length === 0 ? (
            <Box className="rounded-lg border border-dashed border-slate-200 p-6 text-center text-xs text-slate-400">
              No edit history in this session yet.
            </Box>
          ) : (
            <Stack gap="xs">
              <Text size="xs" weight="semibold">Session undo stack ({entries.length})</Text>
              {entries.map((entry, idx) => (
                <Box
                  key={entry.operationId}
                  className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2.5 text-xs shadow-2xs"
                >
                  <Stack gap="none">
                    <Text size="sm" weight="semibold" className="text-slate-800">
                      {entry.description || 'Workbook Mutation'}
                    </Text>
                    <Text size="xs" tone="subtle">
                      {new Date(entry.timestamp).toLocaleTimeString()} · {entry.inversePlan.length} inverse ops · {entry.status}
                    </Text>
                  </Stack>
                  <Inline gap="sm" className="items-center">
                    <Text size="xs" weight="medium" className="rounded bg-slate-100 px-1.5 py-0.5 font-mono text-slate-600">
                      #{idx + 1}
                    </Text>
                    {onUndoTo ? (
                      <Button variant="ghost" size="sm" onClick={() => onUndoTo(idx)}>
                        Restore
                      </Button>
                    ) : null}
                  </Inline>
                </Box>
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
