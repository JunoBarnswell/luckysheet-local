import React from 'react';
import { Box, Button, Panel, PanelBody, PanelHeader, PanelTitle, Stack, Text } from '@react-sheets/ui-system';
import type { CompatibilityReport } from '@react-sheets/exchange-xlsx';
import type { XlsxLayoutRepairPlan } from '@react-sheets/exchange-xlsx';

export interface CompatibilityReportPanelProps {
  report: CompatibilityReport | null;
  onClear?: () => void;
  needsLayoutRepair?: boolean;
  repairPreview?: XlsxLayoutRepairPlan | null;
  onPreviewRepair?: () => void;
  onApplyRepair?: () => void;
}

export function CompatibilityReportPanel({ report, onClear, needsLayoutRepair = false, repairPreview, onPreviewRepair, onApplyRepair }: CompatibilityReportPanelProps) {
  if (!report && !needsLayoutRepair) {
    return (
      <Panel className="shadow-none">
        <PanelBody>
          <Text size="xs" tone="subtle">No compatibility report yet. Import or export an XLSX file to see feature mapping.</Text>
        </PanelBody>
      </Panel>
    );
  }

  return (
    <Panel className="shadow-none">
      <PanelHeader className="border-b border-slate-200">
        <Stack gap="xs" className="w-full">
          <PanelTitle size="sm">XLSX Compatibility</PanelTitle>
          {report ? <Text size="xs" tone="subtle">{report.fileName} · {report.dateSystem} date system · level {report.importLevel}</Text> : <Text size="xs" tone="subtle">Stored source artifact</Text>}
        </Stack>
      </PanelHeader>
      <PanelBody>
        <Stack gap="sm">
          {needsLayoutRepair ? (
            <Stack gap="xs" className="rounded-lg border border-amber-300 bg-amber-50 p-3">
              <Text size="sm" weight="semibold">Imported layout needs repair</Text>
              <Text size="xs" tone="muted">The original XLSX is available. Repair only updates row/column geometry, font units and pane metadata.</Text>
              {repairPreview ? <Text size="xs">Preview: {repairPreview.summary.columns} columns · {repairPreview.summary.rows} rows · {repairPreview.summary.fonts} fonts · {repairPreview.summary.panes} panes</Text> : null}
              {repairPreview ? <Button size="sm" variant="primary" onClick={onApplyRepair}>Apply layout repair</Button> : <Button size="sm" variant="secondary" onClick={onPreviewRepair}>Preview layout repair</Button>}
            </Stack>
          ) : null}
          {report ? (
            <>
          <Text size="xs">
            Editable {report.summary.editableFeatures} · Preserved {report.summary.preservedOnly} · Unsupported {report.summary.unsupported}
          </Text>
          {report.issues.length === 0 ? (
            <Text size="xs" tone="subtle">All detected features are fully supported.</Text>
          ) : (
            <Stack gap="xs">
              {report.issues.slice(0, 12).map((issue, index) => (
                <Box key={`${issue.feature}-${index}`} className="rounded-lg border border-slate-200 bg-white p-2 text-xs">
                  <Text size="xs" weight="semibold">{issue.feature}</Text>
                  <Text size="xs" tone="subtle">{issue.message}</Text>
                  {issue.location ? <Text size="xs" tone="subtle">{issue.location}</Text> : null}
                </Box>
              ))}
            </Stack>
          )}
          {onClear ? (
            <Button variant="ghost" size="sm" onClick={onClear}>
              Dismiss report
            </Button>
          ) : null}
            </>
          ) : null}
        </Stack>
      </PanelBody>
    </Panel>
  );
}
