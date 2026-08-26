import React, { useEffect, useState } from 'react';
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
  TabList,
  Tabs,
  Text,
  TextInput,
  Textarea,
} from '@react-sheets/ui-system';
import type {
  CellComment,
  ConditionalFormatRule,
  DataValidationRule,
  DrawingObject,
  DrawingPayload,
  PivotFieldDefinition,
  PivotModel,
  SparklineModel,
  SparklineGroup,
  DefinedNameModel,
  DataRelationship,
  DataSourceManifest,
  TableSheetDefinition,
  GanttSheetDefinition,
  ReportSheetDefinition,
  SheetTableModel,
  RangeRef,
} from '@react-sheets/core-model';
import type { HistoryEntry } from '@react-sheets/command-runtime';
import type { RevisionRecord } from '@react-sheets/protocol';
import type { WorkbookTableModel } from '@react-sheets/core-model';
import type { PrintLayout } from '@react-sheets/spreadsheet-app';
import type { QueryDefinition } from '@react-sheets/spreadsheet-app';
import { parseAddress, type CanvasSheetSnapshot, type SidebarPanelId, type AppPhase } from '@react-sheets/spreadsheet-app';
import { localizeText, type Locale } from '../i18n';
import type { PivotPanelCallbacks, PivotPanelState, PivotSlicerControl, PivotTimelineControl } from './pivot/pivot-contract';
import { ChartPanel } from './panels/ChartPanel';
import { DataChartPanel } from './panels/DataChartPanel';
import { BarcodePanel } from './panels/BarcodePanel';
import { PivotPanel } from './panels/PivotPanel';
import { SlicerEditorPanel } from './panels/SlicerEditorPanel';
import { ShapeEditorPanel } from './panels/ShapeEditorPanel';
import { TextBoxEditorPanel } from './panels/TextBoxEditorPanel';
import { FormControlPanel } from './panels/FormControlPanel';
import { PicturePanel } from './panels/PicturePanel';
import { SparklinePanel } from './panels/SparklinePanel';
import { ConditionalFormatPanel } from './panels/ConditionalFormatPanel';
import { DataValidationPanel } from './panels/DataValidationPanel';
import { PrintPanel } from './panels/PrintPanel';
import { QueryPanel } from './panels/QueryPanel';
import { AutomationPanel } from './panels/AutomationPanel';
import { ExtendedPanel } from './panels/ExtendedPanel';
import { HistoryPanel } from './panels/HistoryPanel';
import { CompatibilityReportPanel } from './panels/CompatibilityReportPanel';
import { DataSourcePanel } from './panels/DataSourcePanel';
import { TableSheetDesignerPanel } from './panels/TableSheetDesignerPanel';
import { GanttDesignerPanel } from './panels/GanttDesignerPanel';
import { ReportDesignerPanel } from './panels/ReportDesignerPanel';
import { TableDesignPanel } from './panels/TableDesignPanel';
import { DefinedNamesPanel } from './panels/DefinedNamesPanel';
import { SelectionPane, type DrawingSelectionMode } from './home/SelectionPane';
import {
  FormulaAuditPanel,
  type FormulaAuditPanelCallbacks,
  type FormulaAuditPanelProps,
  type FormulaAuditSectionStates,
} from './panels/FormulaAuditPanel';
import type { CommandDescriptor } from '@react-sheets/command-runtime';

export interface FeatureSidebarProps {
  activePanel: SidebarPanelId;
  locale: Locale;
  activeCell: string;
  /** 主选区(供面板默认范围) */
  selectedRange?: { startRow: number; endRow: number; startColumn: number; endColumn: number };
  onPanelChange: (panel: SidebarPanelId) => void;
  onClosePanel?: () => void;
  onRetry: () => void;
  phase: AppPhase;
  sheet: CanvasSheetSnapshot;
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  initialBarcodeSymbology: import('@react-sheets/core-model').BarcodeSymbology;
  onSelectDrawing: (drawingId: string, mode: DrawingSelectionMode) => void;
  onSetDrawingVisibility: (drawingId: string, visible: boolean) => void;
  onRenameDrawing: (drawingId: string, name: string) => void;
  onReorderDrawing: (drawingId: string, direction: 'forward' | 'backward') => void;
  pivot?: PivotModel;
  pivotList?: readonly { id: string; label: string }[];
  activePivotId?: string;
  pivotFieldCatalog?: readonly PivotFieldDefinition[];
  pivotSlicerControls?: readonly PivotSlicerControl[];
  pivotTimelineControls?: readonly PivotTimelineControl[];
  pivotPanelState?: PivotPanelState;
  pivotCallbacks?: PivotPanelCallbacks;
  formulaAudit?: FormulaAuditPanelProps['projection'];
  formulaAuditState?: FormulaAuditPanelProps['state'];
  formulaAuditError?: string;
  formulaAuditSectionStates?: FormulaAuditSectionStates;
  formulaAuditCallbacks?: FormulaAuditPanelCallbacks;
  definedNames?: readonly DefinedNameModel[];
  onSaveDefinedName: (input: DefinedNameModel) => void;
  onRemoveDefinedName: (input: DefinedNameModel) => void;
  sparklines: SparklineModel[];
  sparklineGroups: SparklineGroup[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
  historyPreviewRevision?: number | null;
  canRestoreHistory?: boolean;
  onUndoToHistory: (index: number) => void;
  onRestoreRevision: (revision: number) => void;
  onPreviewRevision: (revision: number) => void;
  onClearHistoryPreview: () => void;
  onRefreshRevisions: () => void;
  compatibilityReport?: import('@react-sheets/exchange-excel-ooxml').CompatibilityReport | null;
  onClearCompatibilityReport: () => void;
  tables: readonly WorkbookTableModel[];
  dataSources: readonly DataSourceManifest[];
  relationships: readonly DataRelationship[];
  onUpdateTableSheet: (definition: TableSheetDefinition) => void;
  onUpdateGanttSheet: (definition: GanttSheetDefinition) => void;
  onUpdateReportSheet: (definition: ReportSheetDefinition) => void;
  activeTable?: SheetTableModel;
  onTableNameChange: (name: string) => void;
  onToggleTableOption: (option: 'hasHeaderRow' | 'showFirstColumn' | 'showLastColumn' | 'showBandedRows' | 'showBandedColumns' | 'showFilterButton') => void;
  onResizeTable: (range: RangeRef) => void;
  onTableStyleChange: (styleName: string) => void;
  onConvertTableToRange: () => void;
  onRemoveDataSource: (sourceId: string) => void;
  onRemoveDataRegion: (regionId: string) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onAddSparkline: (sparkline: SparklineModel) => void;
  onCreateSparklineGroup: (sparklineIds: string[]) => void;
  onRemoveSparkline: (id: string) => void;
  onInsertChart: (type: import('@react-sheets/core-model').ChartDrawingPayload['chartType'], sourceRange: RangeRef, title: string, stacked: NonNullable<import('@react-sheets/core-model').ChartDrawingPayload['stacked']>) => void;
  onAddConditionalFormat: (rule: ConditionalFormatRule) => void;
  onRemoveConditionalFormat: (id: string) => void;
  onAddDataValidation: (rule: DataValidationRule) => void;
  onRemoveDataValidation: (id: string) => void;
  onPrint: (layout: PrintLayout) => void;
  onExportPdf: (layout: PrintLayout) => void;
  printPageCount?: number;
  queryConnectors?: readonly string[];
  loadedQueries?: readonly { queryId: string; queryName: string; columns: readonly string[]; rowCount: number; loadedAt: string }[];
  lastQueryResult?: { queryId: string; queryName: string; columns: readonly string[]; rowCount: number; loadedAt: string } | null;
  canQuery: boolean;
  onLoadQuery: (query: QueryDefinition) => Promise<void>;
  onRefreshQuery: (queryId: string) => Promise<void>;
  onTestQueryConnection: (connectorId: string, config: Record<string, unknown>) => Promise<{ ok: boolean; message?: string }>;
  automationRecording?: boolean;
  recordedScript?: string;
  lastScriptResult?: { ok: boolean; durationMs: number; error?: string } | null;
  canRunScripts: boolean;
  onRunAutomationScript: (source: string) => void;
  onStartAutomationRecording: () => void;
  onStopAutomationRecording: () => void;
  lastWhatIfMessage?: string | null;
  canRunExtended: boolean;
  onGoalSeek: (params: { setRow: number; setColumn: number; targetValue: number; changingRow: number; changingColumn: number }) => void;
  onRunScenario: (params: {
    name: string;
    changingCell: { row: number; column: number };
    changingValue: number;
    resultCell: { row: number; column: number };
  }) => void;
  onAddComment: (text: string) => void;
  onReplyComment: (text: string) => void;
  onResolveComment: () => void;
  onRemoveComment: () => void;
  onAddNote: (text: string) => void;
  onRemoveNote: () => void;
}

const panels: Array<{ icon: React.ComponentProps<typeof Icon>['name']; id: SidebarPanelId; label: string }> = [
  { id: 'inspector', label: 'Inspect', icon: 'sliders' },
  { id: 'chart', label: 'Chart', icon: 'chart' },
  { id: 'dataChart', label: 'Data Chart', icon: 'data-chart' },
  { id: 'barcode', label: 'Barcode', icon: 'barcode' },
  { id: 'pivot', label: 'Pivot', icon: 'table-pivot' },
  { id: 'slicer', label: 'Slicer', icon: 'sliders' },
  { id: 'formulaAudit', label: 'Formula Audit', icon: 'function' },
  { id: 'definedNames', label: 'Names', icon: 'function' },
  { id: 'shape', label: 'Shape', icon: 'shape-square' },
  { id: 'textbox', label: 'Text Box', icon: 'textbox' },
  { id: 'formControl', label: 'Form Control', icon: 'sliders' },
  { id: 'picture', label: 'Picture', icon: 'picture' },
  { id: 'selectionPane', label: 'Selection', icon: 'shape-square' },
  { id: 'sparkline', label: 'Spark', icon: 'sparkline' },
  { id: 'conditionalFormat', label: 'Format', icon: 'sparkles' },
  { id: 'dataValidation', label: 'Validate', icon: 'check-circle' },
  { id: 'print', label: 'Print', icon: 'printer' },
  { id: 'query', label: 'Query', icon: 'table' },
  { id: 'automate', label: 'Automate', icon: 'function' },
  { id: 'extended', label: 'Extended', icon: 'sparkles' },
  { id: 'history', label: 'History', icon: 'history' },
  { id: 'data', label: 'Tables', icon: 'table' },
];

function InsightRow({ label, tone = 'muted', value }: { label: string; tone?: 'accent' | 'muted' | 'success'; value: string }) {
  return (
    <Inline gap="sm" className="justify-between border-b border-slate-100 py-2.5 last:border-0">
      <Text size="xs" tone="muted">{label}</Text>
      <Text size="xs" tone={tone} weight="semibold">{value}</Text>
    </Inline>
  );
}

function InspectorPanel({
  activeCell,
  sheet,
  compatibilityReport,
  onClearCompatibilityReport,
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onAddNote,
  onRemoveNote,
}: {
  activeCell: string;
  sheet: CanvasSheetSnapshot;
  compatibilityReport?: import('@react-sheets/exchange-excel-ooxml').CompatibilityReport | null;
  onClearCompatibilityReport: () => void;
  onAddComment: (text: string) => void;
  onReplyComment: (text: string) => void;
  onResolveComment: () => void;
  onRemoveComment: () => void;
  onAddNote: (text: string) => void;
  onRemoveNote: () => void;
}) {
  const selectedAddress = parseAddress(activeCell);
  const selected = selectedAddress ? sheet.getCell(selectedAddress.row, selectedAddress.column) : undefined;
  const selectedCell = selected;
  const average = '—';

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
            <Text as="span" size="xs" weight="bold" className="rounded bg-blue-100 px-2 py-0.5 text-blue-700">
              {selected?.formula ? 'Formula: ' + selected.formula : 'Direct Value'}
            </Text>
          </Inline>
        </PanelBody>
      </Panel>

      <CommentHyperlinkForms
        comment={selectedCell?.comment}
        commentText={selectedCell?.commentText ?? ''}
        note={selectedCell?.note}
        onAddComment={onAddComment}
        onReplyComment={onReplyComment}
        onResolveComment={onResolveComment}
        onRemoveComment={onRemoveComment}
        onAddNote={onAddNote}
        onRemoveNote={onRemoveNote}
      />

      <CompatibilityReportPanel report={compatibilityReport ?? null} onClear={onClearCompatibilityReport} />

      <Panel className="shadow-none">
        <PanelHeader>
          <Inline gap="sm">
            <Icon name="chart" size="sm" className="text-blue-600" />
            <PanelTitle as="h3" size="sm">Worksheet Statistics</PanelTitle>
          </Inline>
        </PanelHeader>
        <PanelBody className="py-2">
          <InsightRow label="Numeric average" value={average} tone="accent" />
          <InsightRow label="Occupied cells" value={String(sheet.occupiedCellCount)} />
          <InsightRow label="Worksheet name" value={sheet.name} />
          <InsightRow label="Columns count" value={String(sheet.columnCount)} />
        </PanelBody>
      </Panel>
    </Stack>
  );
}

export function FeatureSidebar({
  activePanel,
  locale,
  activeCell,
  selectedRange,
  onPanelChange,
  onClosePanel,
  onRetry,
  phase,
  sheet,
  sheetId,
  drawings,
  drawingPayloads,
  selectedDrawingIds = [],
  initialBarcodeSymbology,
  onSelectDrawing,
  onSetDrawingVisibility,
  onRenameDrawing,
  onReorderDrawing,
  pivot,
  pivotList,
  activePivotId,
  pivotFieldCatalog,
  pivotSlicerControls,
  pivotTimelineControls,
  pivotPanelState,
  pivotCallbacks,
  formulaAudit,
  formulaAuditState,
  formulaAuditError,
  formulaAuditSectionStates,
  formulaAuditCallbacks,
  definedNames = [],
  onSaveDefinedName,
  onRemoveDefinedName,
  sparklines,
  sparklineGroups,
  conditionalFormats,
  dataValidations,
  historyEntries,
  remoteRevisions,
  historyPreviewRevision = null,
  canRestoreHistory = true,
  onUndoToHistory,
  onRestoreRevision,
  onPreviewRevision,
  onClearHistoryPreview,
  onRefreshRevisions,
  compatibilityReport = null,
  onClearCompatibilityReport,
  tables,
  dataSources,
  relationships,
  onUpdateTableSheet,
  onUpdateGanttSheet,
  onUpdateReportSheet,
  activeTable,
  onTableNameChange,
  onToggleTableOption,
  onResizeTable,
  onTableStyleChange,
  onConvertTableToRange,
  onRemoveDataSource,
  onRemoveDataRegion,
  onCommand,
  onAddSparkline,
  onCreateSparklineGroup,
  onRemoveSparkline,
  onInsertChart,
  onAddConditionalFormat,
  onRemoveConditionalFormat,
  onAddDataValidation,
  onRemoveDataValidation,
  onPrint,
  onExportPdf,
  printPageCount = 0,
  queryConnectors = [],
  loadedQueries = [],
  lastQueryResult = null,
  canQuery,
  onLoadQuery,
  onRefreshQuery,
  onTestQueryConnection,
  automationRecording = false,
  recordedScript = '',
  lastScriptResult = null,
  canRunScripts,
  onRunAutomationScript,
  onStartAutomationRecording,
  onStopAutomationRecording,
  lastWhatIfMessage = null,
  canRunExtended,
  onGoalSeek,
  onRunScenario,
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onAddNote,
  onRemoveNote,
}: FeatureSidebarProps) {
  const tableResizeRange = activeTable && selectedRange
    && (selectedRange.startRow !== activeTable.range.startRow
      || selectedRange.endRow !== activeTable.range.endRow
      || selectedRange.startColumn !== activeTable.range.startColumn
      || selectedRange.endColumn !== activeTable.range.endColumn)
    ? { sheetId: activeTable.sheetId, ...selectedRange }
    : undefined;
  const disabled = phase !== 'ready';
  const activePanelLabel = sheet.tableSheet && activePanel === 'data'
    ? localizeText(locale, 'TableSheet Designer')
    : sheet.ganttSheet && activePanel === 'data'
      ? localizeText(locale, 'GanttSheet Designer')
      : sheet.reportSheet && activePanel === 'data'
        ? localizeText(locale, 'ReportSheet Designer')
    : localizeText(locale, panels.find((panel) => panel.id === activePanel)?.label ?? 'Inspect');

  const columnLabelOf = (column: number): string => {
    let label = '';
    let remaining = column + 1;
    while (remaining > 0) {
      const modulo = (remaining - 1) % 26;
      label = String.fromCharCode(65 + modulo) + label;
      remaining = Math.floor((remaining - 1) / 26);
    }
    return label;
  };
  const selectionText = selectedRange
    ? columnLabelOf(selectedRange.startColumn) + (selectedRange.startRow + 1)
      + ':'
      + columnLabelOf(selectedRange.endColumn) + (selectedRange.endRow + 1)
    : undefined;

  if (activePanel === 'pivot' && phase === 'ready') {
    return (
      <Box as="aside" aria-label="Feature sidebar" className="flex h-full min-h-0 flex-1 flex-col bg-white">
        <PivotPanel locale={locale} pivot={pivot} pivotList={pivotList} activePivotId={activePivotId} fieldCatalog={pivotFieldCatalog} slicerControls={pivotSlicerControls} timelineControls={pivotTimelineControls} state={pivotPanelState} callbacks={pivotCallbacks} onClose={onClosePanel} />
      </Box>
    );
  }

  return (
    <Box
      as="aside"
      aria-label="Feature sidebar"
      className="flex min-h-0 flex-1 flex-col"
    >
      <Tabs className="shrink-0 border-b border-slate-200 bg-white px-2 pt-2">
        <TabList label="Feature panels" className="grid grid-cols-5 gap-0.5">
          {panels.slice(0, 5).map((panel) => (
            <Button
              key={panel.id}
              aria-pressed={panel.id === activePanel}
              disabled={disabled}
              onClick={() => onPanelChange(panel.id)}
              onPointerDown={() => onPanelChange(panel.id)}
              onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onPanelChange(panel.id); } }}
              variant={panel.id === activePanel ? 'soft' : 'ghost'}
              className="flex-col gap-0.5 px-0.5 py-1.5"
            >
              <Icon name={panel.icon} size="xs" />
              <Text as="span" size="xs" className="font-medium leading-none">{localizeText(locale, panel.label)}</Text>
            </Button>
          ))}
        </TabList>
        <TabList label="Pro feature panels" className="grid grid-cols-4 gap-0.5 border-t border-slate-100 py-1">
          {panels.slice(5).map((panel) => (
            <Button
              key={panel.id}
              aria-pressed={panel.id === activePanel}
              disabled={disabled}
              onClick={() => onPanelChange(panel.id)}
              onPointerDown={() => onPanelChange(panel.id)}
              onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); onPanelChange(panel.id); } }}
              variant={panel.id === activePanel ? 'soft' : 'ghost'}
              className="flex-col gap-0.5 px-0.5 py-1"
            >
              <Icon name={panel.icon} size="xs" />
              <Text as="span" size="xs" className="font-medium leading-none">{localizeText(locale, panel.label)}</Text>
            </Button>
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

        {phase === 'ready' && activePanel === 'inspector' ? (
          <InspectorPanel
            activeCell={activeCell}
            sheet={sheet}
            compatibilityReport={compatibilityReport}
            onClearCompatibilityReport={onClearCompatibilityReport}
            onAddComment={onAddComment}
            onReplyComment={onReplyComment}
            onResolveComment={onResolveComment}
            onRemoveComment={onRemoveComment}
            onAddNote={onAddNote}
            onRemoveNote={onRemoveNote}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'slicer' ? (
          <SlicerEditorPanel sheetId={sheetId} drawings={drawings} drawingPayloads={drawingPayloads} selectedDrawingIds={selectedDrawingIds} onCommand={onCommand} />
        ) : null}
        {phase === 'ready' && activePanel === 'chart' ? (
          <ChartPanel
            sheetId={sheetId}
            drawings={drawings}
            drawingPayloads={drawingPayloads}
            selectedDrawingIds={selectedDrawingIds}
            defaultRange={selectionText}
            onInsertChart={onInsertChart}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'dataChart' ? (
          <DataChartPanel
            sheetId={sheetId}
            sheet={sheet}
            drawings={drawings}
            drawingPayloads={drawingPayloads}
            selectedDrawingIds={selectedDrawingIds}
            tables={tables}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'barcode' ? (
          <BarcodePanel
            sheetId={sheetId}
            sheet={sheet}
            activeCell={activeCell}
            selectedRange={selectedRange}
            initialSymbology={initialBarcodeSymbology}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'formulaAudit' ? (
          <FormulaAuditPanel
            activeCell={activeCell}
            callbacks={formulaAuditCallbacks}
            errorMessage={formulaAuditError}
            locale={locale}
            projection={formulaAudit}
            sectionStates={formulaAuditSectionStates}
            state={formulaAuditState}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'definedNames' ? (
          <DefinedNamesPanel
            sheetId={sheetId}
            names={definedNames}
            onSave={onSaveDefinedName}
            onRemove={onRemoveDefinedName}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'shape' ? (
          <ShapeEditorPanel
            sheetId={sheetId}
            drawings={drawings}
            drawingPayloads={drawingPayloads}
            drawingGroups={sheet.drawingGroups}
            snapSettings={sheet.snapSettings}
            selectedDrawingIds={selectedDrawingIds}
            onSelectDrawing={(drawingId) => onSelectDrawing(drawingId, 'replace')}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'textbox' ? (
          <TextBoxEditorPanel
            sheetId={sheetId}
            drawings={drawings}
            drawingPayloads={drawingPayloads}
            selectedDrawingIds={selectedDrawingIds}
            onCommand={onCommand}
            onClose={onClosePanel}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'formControl' ? (
          <FormControlPanel
            sheetId={sheetId}
            drawings={drawings}
            drawingPayloads={drawingPayloads}
            selectedDrawingIds={selectedDrawingIds}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'selectionPane' ? (
          <SelectionPane
            locale={locale}
            disabled={disabled}
            items={drawings.map((drawing) => ({
              id: drawing.id,
              kind: drawing.kind === 'slicer' || drawing.kind === 'timeline'
                ? 'pivot-control'
                : drawing.kind,
              name: drawing.name,
              visible: drawing.visible !== false,
              zIndex: drawing.zIndex,
            }))}
            selectedIds={selectedDrawingIds}
            onSelect={onSelectDrawing}
            onVisibilityChange={onSetDrawingVisibility}
            onRename={onRenameDrawing}
            onReorder={onReorderDrawing}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'sparkline' ? (
          <SparklinePanel
            sheetId={sheetId}
            activeCell={activeCell}
            sparklines={sparklines}
            sparklineGroups={sparklineGroups}
            defaultRange={selectionText}
            onAddSparkline={onAddSparkline}
            onCreateSparklineGroup={onCreateSparklineGroup}
            onRemoveSparkline={onRemoveSparkline}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'picture' ? (
          <PicturePanel
            sheetId={sheetId}
            activeCell={activeCell}
            sheet={sheet}
            selectedDrawingIds={selectedDrawingIds}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'conditionalFormat' ? (
          <ConditionalFormatPanel
            sheetId={sheetId}
            range={selectedRange
              ? { sheetId, ...selectedRange }
              : { ...sheet.usedRange, sheetId }}
            locale={locale}
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
          <PrintPanel onPrint={onPrint} onExportPdf={onExportPdf} pageCount={printPageCount} />
        ) : null}
        {phase === 'ready' && activePanel === 'automate' ? (
          <AutomationPanel
            recording={automationRecording}
            recordedScript={recordedScript}
            lastResult={lastScriptResult}
            canRunScripts={canRunScripts}
            onRunScript={onRunAutomationScript}
            onStartRecording={onStartAutomationRecording}
            onStopRecording={onStopAutomationRecording}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'extended' ? (
          <ExtendedPanel
            lastWhatIfMessage={lastWhatIfMessage}
            canRunExtended={canRunExtended}
            sheetId={sheetId}
            onGoalSeek={onGoalSeek}
            onRunScenario={onRunScenario}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'query' ? (
          <QueryPanel
            connectors={queryConnectors}
            loadedQueries={loadedQueries}
            lastResult={lastQueryResult}
            canQuery={canQuery}
            onLoadQuery={onLoadQuery}
            onRefreshQuery={onRefreshQuery}
            onTestConnection={onTestQueryConnection}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'history' ? (
          <HistoryPanel
            entries={historyEntries}
            remoteRevisions={remoteRevisions}
            previewRevision={historyPreviewRevision}
            canRestore={canRestoreHistory}
            onUndoTo={onUndoToHistory}
            onRestoreRevision={onRestoreRevision}
            onPreviewRevision={onPreviewRevision}
            onClearPreview={onClearHistoryPreview}
            onRefreshRevisions={onRefreshRevisions}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'data' ? (
          sheet.tableSheet ? (
            <TableSheetDesignerPanel definition={sheet.tableSheet} tables={tables} relationships={relationships} onUpdate={onUpdateTableSheet} />
          ) : sheet.ganttSheet ? (
            <GanttDesignerPanel definition={sheet.ganttSheet} tables={tables} onUpdate={onUpdateGanttSheet} />
          ) : sheet.reportSheet ? (
            <ReportDesignerPanel definition={sheet.reportSheet} tables={tables} activeCell={activeCell} onUpdate={onUpdateReportSheet} />
          ) : activeTable ? (
            <TableDesignPanel table={activeTable} locale={locale} selectedRange={tableResizeRange} onNameChange={onTableNameChange} onToggle={onToggleTableOption} onResize={onResizeTable} onStyleChange={onTableStyleChange} onConvert={onConvertTableToRange} />
          ) : (
            <DataSourcePanel sources={dataSources} regions={sheet.dataRegions} onRemoveSource={onRemoveDataSource} onRemoveRegion={onRemoveDataRegion} />
          )
        ) : null}
      </Box>
    </Box>
  );
}


function CommentHyperlinkForms({
  comment,
  commentText: initialCommentText,
  note,
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onAddNote,
  onRemoveNote,
}: {
  comment?: CellComment;
  commentText: string;
  note?: import('@react-sheets/core-model').CellNote;
  onAddComment: (text: string) => void;
  onReplyComment: (text: string) => void;
  onResolveComment: () => void;
  onRemoveComment: () => void;
  onAddNote: (text: string) => void;
  onRemoveNote: () => void;
}) {
  const [commentText, setCommentText] = useState('');
  const [noteText, setNoteText] = useState('');
  const [replyText, setReplyText] = useState('');

  useEffect(() => setCommentText(initialCommentText), [initialCommentText]);
  useEffect(() => setNoteText(note?.text ?? ''), [note?.id, note?.text]);
  useEffect(() => setReplyText(''), [comment?.id]);

  return (
    <Stack gap="md">
      <Panel className="shadow-none">
        <PanelHeader>
          <Inline gap="sm">
            <Icon name="comment" size="sm" className="text-amber-600" />
            <PanelTitle as="h3" size="sm">Note</PanelTitle>
          </Inline>
        </PanelHeader>
        <PanelBody>
          <Stack gap="sm">
            <Textarea
              rows={2}
              placeholder="Add a cell note (separate from threaded comments)"
              value={noteText}
              onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) => setNoteText(event.target.value)}
            />
            <Inline gap="sm" className="justify-end">
              {note ? (
                <Button size="sm" variant="ghost" onClick={() => { onRemoveNote(); setNoteText(''); }}>
                  Remove
                </Button>
              ) : null}
              <Button
                size="sm"
                variant="primary"
                onClick={() => {
                  if (!noteText.trim()) return;
                  onAddNote(noteText.trim());
                }}
              >
                Save note
              </Button>
            </Inline>
          </Stack>
        </PanelBody>
      </Panel>

      <Panel className="shadow-none">
        <PanelHeader>
          <Inline gap="sm">
            <Icon name="comment" size="sm" className="text-blue-600" />
            <PanelTitle as="h3" size="sm">Comment</PanelTitle>
          </Inline>
        </PanelHeader>
        <PanelBody>
          <Stack gap="sm">
            <Textarea
              rows={3}
              placeholder="Add a comment for the selected cell"
              value={commentText}
              onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) => setCommentText(event.target.value)}
            />
            <Inline gap="sm" className="justify-end">
              <Button size="sm" variant="ghost" onClick={() => { onRemoveComment(); }}>
                Clear
              </Button>
              <Button
                size="sm"
                variant="primary"
                onClick={() => {
                  if (!commentText.trim()) return;
                  onAddComment(commentText.trim());
                  setCommentText('');
                }}
              >
                Save comment
              </Button>
            </Inline>
          </Stack>
        </PanelBody>
      </Panel>

      {comment ? (
        <Panel className="shadow-none">
          <PanelHeader>
            <Inline gap="sm">
              <Icon name="comment" size="sm" className="text-blue-600" />
              <PanelTitle as="h3" size="sm">Thread</PanelTitle>
              {comment.resolved ? <Text size="xs" tone="success">Resolved</Text> : null}
            </Inline>
          </PanelHeader>
          <PanelBody>
            <Stack gap="sm">
              {(comment.replies ?? []).map((reply) => (
                <Box key={reply.id} className="rounded-lg bg-slate-50 p-2">
                  <Text size="xs" weight="semibold">{reply.author}</Text>
                  <Text size="xs" className="mt-1 block">{reply.text}</Text>
                </Box>
              ))}
              <Textarea rows={2} aria-label="Reply to comment" placeholder="Reply to this comment" value={replyText} onChange={(event) => setReplyText(event.target.value)} />
              <Inline gap="sm" className="justify-end">
                {!comment.resolved ? <Button size="xs" variant="ghost" onClick={onResolveComment}>Resolve</Button> : null}
                <Button size="xs" variant="primary" disabled={!replyText.trim()} onClick={() => { onReplyComment(replyText.trim()); setReplyText(''); }}>Reply</Button>
              </Inline>
            </Stack>
          </PanelBody>
        </Panel>
      ) : null}

    </Stack>
  );
}
