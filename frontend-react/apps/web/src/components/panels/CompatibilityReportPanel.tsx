import React from 'react';
import { Button, Panel, PanelBody, PanelHeader, PanelTitle, Stack, Text } from '@react-sheets/ui-system';
import type { CompatibilityReport } from '@react-sheets/exchange-xlsx';

export interface CompatibilityReportPanelProps {
  report: CompatibilityReport | null;
  onClear?: () => void;
}

export function CompatibilityReportPanel({ report, onClear }: CompatibilityReportPanelProps) {
  if (!report) {
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
          <Text size="xs" tone="subtle">{report.fileName} · {report.dateSystem} date system · level {report.importLevel}</Text>
        </Stack>
      </PanelHeader>
      <PanelBody>
        <Stack gap="sm">
          <Text size="xs">
            Editable {report.summary.editableFeatures} · Preserved {report.summary.preservedOnly} · Unsupported {report.summary.unsupported}
          </Text>
          {report.issues.length === 0 ? (
            <Text size="xs" tone="subtle">All detected features are fully supported.</Text>
          ) : (
            <Stack gap="xs">
              {report.issues.slice(0, 12).map((issue, index) => (
                <div key={`${issue.feature}-${index}`} className="rounded-lg border border-slate-200 bg-white p-2 text-xs">
                  <Text size="xs" weight="semibold">{issue.feature}</Text>
                  <Text size="xs" tone="subtle">{issue.message}</Text>
                  {issue.location ? <Text size="xs" tone="subtle">{issue.location}</Text> : null}
                </div>
              ))}
            </Stack>
          )}
          {onClear ? (
            <Button variant="ghost" size="sm" onClick={onClear}>
              Dismiss report
            </Button>
          ) : null}
        </Stack>
      </PanelBody>
    </Panel>
  );
}
