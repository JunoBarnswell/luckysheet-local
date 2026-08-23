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
  DefinedNameModel,
} from '@react-sheets/core-model';
import type { HistoryEntry } from '@react-sheets/command-runtime';
import type { RevisionRecord, TableRowsResponse } from '@react-sheets/protocol';
import type { WorkbookTableModel } from '@react-sheets/core-model';
import type { PrintLayout } from '@react-sheets/spreadsheet-app';
import type { QueryDefinition } from '@react-sheets/spreadsheet-app';
import { parseAddress, type CanvasSheetSnapshot, type SidebarPanelId, type AppPhase } from '@react-sheets/spreadsheet-app';
import { localizeText, type Locale } from '../i18n';
import type { PivotPanelCallbacks, PivotPanelState, PivotSlicerControl, PivotTimelineControl } from './pivot/pivot-contract';
import { ChartPanel } from './panels/ChartPanel';
import { PivotPanel } from './panels/PivotPanel';
import { ShapeEditorPanel } from './panels/ShapeEditorPanel';
import { SparklinePanel } from './panels/SparklinePanel';
import { ConditionalFormatPanel } from './panels/ConditionalFormatPanel';
import { DataValidationPanel } from './panels/DataValidationPanel';
import { PrintPanel } from './panels/PrintPanel';
import { QueryPanel } from './panels/QueryPanel';
import { AutomationPanel } from './panels/AutomationPanel';
import { ExtendedPanel } from './panels/ExtendedPanel';
import { HistoryPanel } from './panels/HistoryPanel';
import { CompatibilityReportPanel } from './panels/CompatibilityReportPanel';
import { DataModelPanel } from './panels/DataModelPanel';
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
  onRetry: () => void;
  phase: AppPhase;
  sheet: CanvasSheetSnapshot;
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  onSelectDrawing?: (drawingId: string, mode: DrawingSelectionMode) => void;
  onSetDrawingVisibility?: (drawingId: string, visible: boolean) => void;
  onRenameDrawing?: (drawingId: string, name: string) => void;
  onReorderDrawing?: (drawingId: string, direction: 'forward' | 'backward') => void;
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
  onSaveDefinedName?: (input: DefinedNameModel) => void;
  onRemoveDefinedName?: (input: DefinedNameModel) => void;
  sparklines: SparklineModel[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
  historyPreviewRevision?: number | null;
  canRestoreHistory?: boolean;
  onUndoToHistory?: (index: number) => void;
  onRestoreRevision?: (revision: number) => void;
  onPreviewRevision?: (revision: number) => void;
  onClearHistoryPreview?: () => void;
  onRefreshRevisions?: () => void;
  compatibilityReport?: import('@react-sheets/exchange-xlsx').CompatibilityReport | null;
  onClearCompatibilityReport?: () => void;
  tables: readonly WorkbookTableModel[];
  onReadDataRows: (tableId: string, offset?: number, limit?: number) => Promise<TableRowsResponse>;
  onRemoveDataTable: (tableId: string) => Promise<void>;
  onCommand: (descriptor: CommandDescriptor) => void;
  onAddSparkline: (sparkline: SparklineModel) => void;
  onRemoveSparkline: (id: string) => void;
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
  canQuery?: boolean;
  onLoadQuery?: (query: QueryDefinition) => Promise<void>;
  onRefreshQuery?: (queryId: string) => Promise<void>;
  onTestQueryConnection?: (connectorId: string, config: Record<string, unknown>) => Promise<{ ok: boolean; message?: string }>;
  automationRecording?: boolean;
  recordedScript?: string;
  lastScriptResult?: { ok: boolean; durationMs: number; error?: string } | null;
  canRunScripts?: boolean;
  onRunAutomationScript?: (source: string) => void;
  onStartAutomationRecording?: () => void;
  onStopAutomationRecording?: () => void;
  lastWhatIfMessage?: string | null;
  canRunExtended?: boolean;
  onGoalSeek?: (params: { setRow: number; setColumn: number; targetValue: number; changingRow: number; changingColumn: number }) => void;
  onRunDataTable?: (params: {
    inputMode: 'row' | 'column';
    inputCell: { row: number; column: number };
    tableRange: { startRow: number; startColumn: number; endRow: number; endColumn: number };
  }) => void;
  onRunScenario?: (params: {
    name: string;
    changingCell: { row: number; column: number };
    changingValue: number;
    resultCell: { row: number; column: number };
  }) => void;
  onAddComment?: (text: string) => void;
  onReplyComment?: (text: string) => void;
  onResolveComment?: () => void;
  onRemoveComment?: () => void;
  onAddNote?: (text: string) => void;
  onRemoveNote?: () => void;
  onSetHyperlink?: (url: string) => void;
  onRemoveHyperlink?: () => void;
}

const panels: Array<{ icon: React.ComponentProps<typeof Icon>['name']; id: SidebarPanelId; label: string }> = [
  { id: 'inspector', label: 'Inspect', icon: 'sliders' },
  { id: 'chart', label: 'Chart', icon: 'chart' },
  { id: 'pivot', label: 'Pivot', icon: 'table-pivot' },
  { id: 'formulaAudit', label: 'Formula Audit', icon: 'function' },
  { id: 'definedNames', label: 'Names', icon: 'function' },
  { id: 'shape', label: 'Shape', icon: 'shape-square' },
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
  onSetHyperlink,
  onRemoveHyperlink,
}: {
  activeCell: string;
  sheet: CanvasSheetSnapshot;
  compatibilityReport?: import('@react-sheets/exchange-xlsx').CompatibilityReport | null;
  onClearCompatibilityReport?: () => void;
  onAddComment?: (text: string) => void;
  onReplyComment?: (text: string) => void;
  onResolveComment?: () => void;
  onRemoveComment?: () => void;
  onAddNote?: (text: string) => void;
  onRemoveNote?: () => void;
  onSetHyperlink?: (url: string) => void;
  onRemoveHyperlink?: () => void;
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
        hyperlinkUrl={selectedCell?.hyperlink ?? ''}
        onAddComment={onAddComment}
        onReplyComment={onReplyComment}
        onResolveComment={onResolveComment}
        onRemoveComment={onRemoveComment}
        onAddNote={onAddNote}
        onRemoveNote={onRemoveNote}
        onSetHyperlink={onSetHyperlink}
        onRemoveHyperlink={onRemoveHyperlink}
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
  onRetry,
  phase,
  sheet,
  sheetId,
  drawings,
  drawingPayloads,
  selectedDrawingIds = [],
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
  onReadDataRows,
  onRemoveDataTable,
  onCommand,
  onAddSparkline,
  onRemoveSparkline,
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
  canQuery = true,
  onLoadQuery,
  onRefreshQuery,
  onTestQueryConnection,
  automationRecording = false,
  recordedScript = '',
  lastScriptResult = null,
  canRunScripts = true,
  onRunAutomationScript,
  onStartAutomationRecording,
  onStopAutomationRecording,
  lastWhatIfMessage = null,
  canRunExtended = true,
  onGoalSeek,
  onRunDataTable,
  onRunScenario,
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onAddNote,
  onRemoveNote,
  onSetHyperlink,
  onRemoveHyperlink,
}: FeatureSidebarProps) {
  const disabled = phase !== 'ready';
  const activePanelLabel = localizeText(locale, panels.find((panel) => panel.id === activePanel)?.label ?? 'Inspect');

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
            onSetHyperlink={onSetHyperlink}
            onRemoveHyperlink={onRemoveHyperlink}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'chart' ? (
          <ChartPanel
            sheetId={sheetId}
            drawings={drawings}
            drawingPayloads={drawingPayloads}
            defaultRange={selectionText}
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'pivot' ? (
          <PivotPanel
            pivot={pivot}
            pivotList={pivotList}
            activePivotId={activePivotId}
            fieldCatalog={pivotFieldCatalog}
            slicerControls={pivotSlicerControls}
            timelineControls={pivotTimelineControls}
            state={pivotPanelState}
            callbacks={pivotCallbacks}
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
        {phase === 'ready' && activePanel === 'definedNames' && onSaveDefinedName && onRemoveDefinedName ? (
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
            onCommand={onCommand}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'selectionPane' && onSelectDrawing && onSetDrawingVisibility ? (
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
            sparklines={sparklines}
            defaultRange={selectionText}
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
          <PrintPanel onPrint={onPrint} onExportPdf={onExportPdf} pageCount={printPageCount} />
        ) : null}
        {phase === 'ready' && activePanel === 'automate' ? (
          <AutomationPanel
            recording={automationRecording}
            recordedScript={recordedScript}
            lastResult={lastScriptResult}
            canRunScripts={canRunScripts}
            onRunScript={onRunAutomationScript ?? (() => undefined)}
            onStartRecording={onStartAutomationRecording ?? (() => undefined)}
            onStopRecording={onStopAutomationRecording ?? (() => undefined)}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'extended' ? (
          <ExtendedPanel
            lastWhatIfMessage={lastWhatIfMessage}
            canRunExtended={canRunExtended}
            sheetId={sheetId}
            onGoalSeek={onGoalSeek ?? (() => undefined)}
            onRunDataTable={onRunDataTable ?? (() => undefined)}
            onRunScenario={onRunScenario ?? (() => undefined)}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'query' ? (
          <QueryPanel
            connectors={queryConnectors}
            loadedQueries={loadedQueries}
            lastResult={lastQueryResult}
            canQuery={canQuery}
            onLoadQuery={onLoadQuery ?? (async () => undefined)}
            onRefreshQuery={onRefreshQuery ?? (async () => undefined)}
            onTestConnection={onTestQueryConnection ?? (async () => ({ ok: false, message: 'Query unavailable' }))}
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
          <DataModelPanel tables={tables} onReadRows={onReadDataRows} onRemove={onRemoveDataTable} />
        ) : null}
      </Box>
    </Box>
  );
}


function CommentHyperlinkForms({
  comment,
  commentText: initialCommentText,
  note,
  hyperlinkUrl: initialHyperlinkUrl,
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onAddNote,
  onRemoveNote,
  onSetHyperlink,
  onRemoveHyperlink,
}: {
  comment?: CellComment;
  commentText: string;
  note?: import('@react-sheets/core-model').CellNote;
  hyperlinkUrl: string;
  onAddComment?: (text: string) => void;
  onReplyComment?: (text: string) => void;
  onResolveComment?: () => void;
  onRemoveComment?: () => void;
  onAddNote?: (text: string) => void;
  onRemoveNote?: () => void;
  onSetHyperlink?: (url: string) => void;
  onRemoveHyperlink?: () => void;
}) {
  const [commentText, setCommentText] = useState('');
  const [noteText, setNoteText] = useState('');
  const [replyText, setReplyText] = useState('');
  const [hyperlinkUrl, setHyperlinkUrl] = useState('');

  useEffect(() => setCommentText(initialCommentText), [initialCommentText]);
  useEffect(() => setNoteText(note?.text ?? ''), [note?.id, note?.text]);
  useEffect(() => setReplyText(''), [comment?.id]);
  useEffect(() => setHyperlinkUrl(initialHyperlinkUrl), [initialHyperlinkUrl]);

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
              {onRemoveNote && note ? (
                <Button size="sm" variant="ghost" onClick={() => { onRemoveNote(); setNoteText(''); }}>
                  Remove
                </Button>
              ) : null}
              <Button
                size="sm"
                variant="primary"
                onClick={() => {
                  if (!onAddNote || !noteText.trim()) return;
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
              {onRemoveComment ? (
                <Button size="sm" variant="ghost" onClick={() => { onRemoveComment(); }}>
                  Clear
                </Button>
              ) : null}
              <Button
                size="sm"
                variant="primary"
                onClick={() => {
                  if (!onAddComment || !commentText.trim()) return;
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
                {onResolveComment && !comment.resolved ? <Button size="xs" variant="ghost" onClick={onResolveComment}>Resolve</Button> : null}
                <Button size="xs" variant="primary" disabled={!replyText.trim() || !onReplyComment} onClick={() => { onReplyComment?.(replyText.trim()); setReplyText(''); }}>Reply</Button>
              </Inline>
            </Stack>
          </PanelBody>
        </Panel>
      ) : null}

      <Panel className="shadow-none">
        <PanelHeader>
          <Inline gap="sm">
            <Icon name="share" size="sm" className="text-blue-600" />
            <PanelTitle as="h3" size="sm">Hyperlink</PanelTitle>
          </Inline>
        </PanelHeader>
        <PanelBody>
          <Stack gap="sm">
            <TextInput
              type="url"
              placeholder="https://example.com"
              value={hyperlinkUrl}
              onChange={(event) => setHyperlinkUrl(event.target.value)}
            />
            <Inline gap="sm" className="justify-end">
              {onRemoveHyperlink ? (
                <Button size="sm" variant="ghost" onClick={() => { onRemoveHyperlink(); }}>
                  Remove
                </Button>
              ) : null}
              <Button
                size="sm"
                variant="primary"
                onClick={() => {
                  if (!onSetHyperlink || !hyperlinkUrl.trim()) return;
                  onSetHyperlink(hyperlinkUrl.trim());
                }}
              >
                Apply link
              </Button>
            </Inline>
          </Stack>
        </PanelBody>
      </Panel>
    </Stack>
  );
}
