import React from 'react';
import {
  Box,
  Button,
  Heading,
  Icon,
  Inline,
  Panel,
  PanelBody,
  PanelHeader,
  PanelTitle,
  Stack,
  StatePanel,
  Tab,
  TabList,
  Tabs,
  Text,
} from '@react-sheets/ui-system';
import type {
  ChartModel,
  ConditionalFormatRule,
  DataValidationRule,
  PivotModel,
  ShapeModel,
  SparklineModel,
} from '@react-sheets/core-model';
import type { HistoryEntry } from '@react-sheets/command-runtime';
import type { PrintLayout } from '@react-sheets/pro-features';
import type { SheetView, SidebarPanelId, WorkspacePhase } from '../state/workspace';
import { ChartPanel } from './panels/ChartPanel';
import { PivotPanel } from './panels/PivotPanel';
import { ShapeEditorPanel } from './panels/ShapeEditorPanel';
import { SparklinePanel } from './panels/SparklinePanel';
import { ConditionalFormatPanel } from './panels/ConditionalFormatPanel';
import { DataValidationPanel } from './panels/DataValidationPanel';
import { PrintPanel } from './panels/PrintPanel';
import { HistoryPanel } from './panels/HistoryPanel';

export interface FeatureSidebarProps {
  activePanel: SidebarPanelId;
  activeCell: string;
  onPanelChange: (panel: SidebarPanelId) => void;
  onRetry: () => void;
  phase: WorkspacePhase;
  sheet: SheetView;
  sheetId: string;
  charts: ChartModel[];
  pivots: PivotModel[];
  shapes: ShapeModel[];
  sparklines: SparklineModel[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  historyEntries: readonly HistoryEntry[];
  onAddChart: (chart: ChartModel) => void;
  onRemoveChart: (id: string) => void;
  onAddPivot: (pivot: PivotModel) => void;
  onRemovePivot: (id: string) => void;
  onAddShape: (shape: ShapeModel) => void;
  onRemoveShape: (id: string) => void;
  onAddSparkline: (sparkline: SparklineModel) => void;
  onRemoveSparkline: (id: string) => void;
  onAddConditionalFormat: (rule: ConditionalFormatRule) => void;
  onRemoveConditionalFormat: (id: string) => void;
  onAddDataValidation: (rule: DataValidationRule) => void;
  onRemoveDataValidation: (id: string) => void;
  onPrint: (layout: PrintLayout) => void;
  onExportPdf: (layout: PrintLayout) => void;
}

const panels: Array<{ icon: React.ComponentProps<typeof Icon>['name']; id: SidebarPanelId; label: string }> = [
  { id: 'inspector', label: 'Inspect', icon: 'sliders' },
  { id: 'chart', label: 'Chart', icon: 'chart' },
  { id: 'pivot', label: 'Pivot', icon: 'table-pivot' },
  { id: 'shape', label: 'Shape', icon: 'shape-square' },
  { id: 'sparkline', label: 'Spark', icon: 'sparkline' },
  { id: 'conditionalFormat', label: 'Format', icon: 'sparkles' },
  { id: 'dataValidation', label: 'Validate', icon: 'check-circle' },
  { id: 'print', label: 'Print', icon: 'printer' },
  { id: 'history', label: 'History', icon: 'history' },
];

function InsightRow({ label, tone = 'muted', value }: { label: string; tone?: 'accent' | 'muted' | 'success'; value: string }) {
  return (
    <Inline gap="sm" className="justify-between border-b border-slate-100 py-2.5 last:border-0">
      <Text size="xs" tone="muted">{label}</Text>
      <Text size="xs" tone={tone} weight="semibold">{value}</Text>
    </Inline>
  );
}

function InspectorPanel({ activeCell, sheet }: { activeCell: string; sheet: SheetView }) {
  const cells = sheet.rows.flatMap((row) => row.cells).filter((cell) => cell.value !== '');
  const selected = cells.find((cell) => cell.address === activeCell);
  const numericValues = cells
    .map((cell) => Number(cell.value.replace(/[$,%]/g, '')))
    .filter((value) => Number.isFinite(value));
  const average =
    numericValues.length > 0
      ? Math.round(numericValues.reduce((sum, value) => sum + value, 0) / numericValues.length).toLocaleString('en-US')
      : '—';

  return (
    <Stack gap="md">
      <Panel tone="accent" className="overflow-hidden shadow-none">
        <PanelBody>
          <Inline gap="sm" className="mb-3">
            <Box className="flex h-9 w-9 items-center justify-center rounded-xl bg-white text-blue-600 shadow-xs">
              <Icon name="sparkles" size="md" />
            </Box>
            <Stack gap="none">
              <PanelTitle as="h3" size="sm">Active Cell Insight</PanelTitle>
              <Text size="xs" tone="muted">Address · {activeCell}</Text>
            </Stack>
          </Inline>
          <Text size="sm" tone="default" weight="semibold">
            {selected?.value || 'Empty cell selected'}
          </Text>
          <Inline gap="xs" className="mt-3">
            <span className="rounded bg-blue-100 px-2 py-0.5 text-[10px] font-bold text-blue-700">
              {selected?.formula ? 'Formula: ' + selected.formula : 'Direct Value'}
            </span>
          </Inline>
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader>
          <Inline gap="sm">
            <Icon name="chart" size="sm" className="text-blue-600" />
            <PanelTitle as="h3" size="sm">Worksheet Statistics</PanelTitle>
          </Inline>
        </PanelHeader>
        <PanelBody className="py-2">
          <InsightRow label="Numeric average" value={average} tone="accent" />
          <InsightRow label="Occupied cells" value={String(cells.length)} />
          <InsightRow label="Worksheet name" value={sheet.name} />
          <InsightRow label="Columns count" value={String(sheet.columns.length)} />
        </PanelBody>
      </Panel>
    </Stack>
  );
}

export function FeatureSidebar({
  activePanel,
  activeCell,
  onPanelChange,
  onRetry,
  phase,
  sheet,
  sheetId,
  charts,
  pivots,
  shapes,
  sparklines,
  conditionalFormats,
  dataValidations,
  historyEntries,
  onAddChart,
  onRemoveChart,
  onAddPivot,
  onRemovePivot,
  onAddShape,
  onRemoveShape,
  onAddSparkline,
  onRemoveSparkline,
  onAddConditionalFormat,
  onRemoveConditionalFormat,
  onAddDataValidation,
  onRemoveDataValidation,
  onPrint,
  onExportPdf,
}: FeatureSidebarProps) {
  const disabled = phase !== 'ready';
  const activePanelLabel = panels.find((panel) => panel.id === activePanel)?.label ?? 'Inspect';

  return (
    <Box
      as="aside"
      aria-label="Feature sidebar"
      className="hidden w-[310px] shrink-0 flex-col border-l border-slate-200 bg-slate-50/70 lg:flex"
    >
      <Tabs className="shrink-0 border-b border-slate-200 bg-white px-2 pt-2">
        <TabList label="Feature panels" className="grid grid-cols-5 gap-0.5">
          {panels.slice(0, 5).map((panel) => (
            <Tab
              key={panel.id}
              active={panel.id === activePanel}
              disabled={disabled}
              onClick={() => onPanelChange(panel.id)}
              className="flex-col gap-0.5 px-0.5 py-1.5"
            >
              <Icon name={panel.icon} size="xs" />
              <span className="text-[10px] font-medium leading-none">{panel.label}</span>
            </Tab>
          ))}
        </TabList>
        <TabList label="Pro feature panels" className="grid grid-cols-4 gap-0.5 border-t border-slate-100 py-1">
          {panels.slice(5).map((panel) => (
            <Tab
              key={panel.id}
              active={panel.id === activePanel}
              disabled={disabled}
              onClick={() => onPanelChange(panel.id)}
              className="flex-col gap-0.5 px-0.5 py-1"
            >
              <Icon name={panel.icon} size="xs" />
              <span className="text-[10px] font-medium leading-none">{panel.label}</span>
            </Tab>
          ))}
        </TabList>
      </Tabs>

      <Box className="min-h-0 flex-1 overflow-auto p-3">
        <Inline gap="sm" className="mb-3">
          <Stack gap="none" className="min-w-0">
            <Text size="xs" tone="subtle" weight="bold" className="uppercase tracking-wider">
              Tool Panel
            </Text>
            <Heading as="h2" size="sm">{activePanelLabel}</Heading>
          </Stack>
        </Inline>

        {phase === 'loading' ? <StatePanel kind="loading" description="Preparing panel data." /> : null}
        {phase === 'error' ? <StatePanel actionLabel="Try again" description="Panel data could not be loaded." kind="error" onAction={onRetry} /> : null}
        {phase === 'empty' ? <StatePanel actionLabel="Try again" description="Open a workbook to inspect." kind="empty" onAction={onRetry} /> : null}

        {phase === 'ready' && activePanel === 'inspector' ? <InspectorPanel activeCell={activeCell} sheet={sheet} /> : null}
        {phase === 'ready' && activePanel === 'chart' ? (
          <ChartPanel
            sheetId={sheetId}
            charts={charts}
            onAddChart={onAddChart}
            onRemoveChart={onRemoveChart}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'pivot' ? (
          <PivotPanel
            sheetId={sheetId}
            pivots={pivots}
            onAddPivot={onAddPivot}
            onRemovePivot={onRemovePivot}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'shape' ? (
          <ShapeEditorPanel
            sheetId={sheetId}
            shapes={shapes}
            onAddShape={onAddShape}
            onRemoveShape={onRemoveShape}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'sparkline' ? (
          <SparklinePanel
            sheetId={sheetId}
            sparklines={sparklines}
            onAddSparkline={onAddSparkline}
            onRemoveSparkline={onRemoveSparkline}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'conditionalFormat' ? (
          <ConditionalFormatPanel
            sheetId={sheetId}
            rules={conditionalFormats}
            onAddRule={onAddConditionalFormat}
            onRemoveRule={onRemoveConditionalFormat}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'dataValidation' ? (
          <DataValidationPanel
            sheetId={sheetId}
            rules={dataValidations}
            onAddRule={onAddDataValidation}
            onRemoveRule={onRemoveDataValidation}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'print' ? (
          <PrintPanel onPrint={onPrint} onExportPdf={onExportPdf} />
        ) : null}
        {phase === 'ready' && activePanel === 'history' ? (
          <HistoryPanel entries={historyEntries} />
        ) : null}
      </Box>
    </Box>
  );
}
