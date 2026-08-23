import React, { useEffect, useState } from 'react';
import { SAMPLE_AUTOMATION_SCRIPT } from '@react-sheets/spreadsheet-app';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Stack, Text, Textarea } from '@react-sheets/ui-system';

export interface AutomationPanelProps {
  recording: boolean;
  recordedScript: string;
  lastResult: { ok: boolean; durationMs: number; error?: string } | null;
  canRunScripts: boolean;
  onRunScript: (source: string) => void;
  onStartRecording: () => void;
  onStopRecording: () => void;
  onClose?: () => void;
}

export function AutomationPanel({
  canRunScripts,
  lastResult,
  onClose,
  onRunScript,
  onStartRecording,
  onStopRecording,
  recordedScript,
  recording,
}: AutomationPanelProps) {
  const [source, setSource] = useState(SAMPLE_AUTOMATION_SCRIPT);

  useEffect(() => {
    if (recordedScript) setSource(recordedScript);
  }, [recordedScript]);

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Automation Script</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Text size="xs" tone="muted">
            Scripts run through the Facade API and execute real workbook commands. Network and filesystem access are blocked.
          </Text>

          <Textarea
            value={source}
            onChange={(event) => setSource(event.target.value)}
            rows={12}
            disabled={!canRunScripts || recording}
            className="font-mono text-xs"
          />

          {recording ? (
            <Box className="rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-800">
              Recording commands… perform edits in the workbook, then stop recording.
            </Box>
          ) : null}

          {lastResult ? (
            <Box className={`rounded-lg border px-3 py-2 text-xs ${lastResult.ok ? 'border-emerald-100 bg-emerald-50 text-emerald-800' : 'border-red-100 bg-red-50 text-red-800'}`}>
              {lastResult.ok
                ? `Last run completed in ${lastResult.durationMs}ms`
                : `Last run failed: ${lastResult.error ?? 'unknown error'}`}
            </Box>
          ) : null}

          <Stack gap="sm">
            {recording ? (
              <Button variant="primary" size="sm" disabled={!canRunScripts} onClick={onStopRecording}>
                Stop recording
              </Button>
            ) : (
              <Button variant="outline" size="sm" icon="history" disabled={!canRunScripts} onClick={onStartRecording}>
                Start recording
              </Button>
            )}
            <Button variant="primary" size="sm" icon="refresh" disabled={!canRunScripts || recording} onClick={() => onRunScript(source)}>
              Run script
            </Button>
          </Stack>
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>Close Panel</Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
