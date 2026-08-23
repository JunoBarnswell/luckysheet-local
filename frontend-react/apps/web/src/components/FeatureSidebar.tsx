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
  ChartModel,
  ConditionalFormatRule,
  DataValidationRule,
  ShapeModel,
  SparklineModel,
} from '@react-sheets/core-model';
import type { HistoryEntry } from '@react-sheets/command-runtime';
import type { RevisionRecord, TableRowsResponse } from '@react-sheets/protocol';
import type { WorkbookTableModel } from '@react-sheets/core-model';
import type { PrintLayout } from '@react-sheets/pro-features';
import { parseAddress, type CanvasSheetSnapshot, type SidebarPanelId, type AppPhase } from '@react-sheets/spreadsheet-app';
import type { PivotModel } from '@react-sheets/core-model';
import { localizeText, type Locale } from '../i18n';
import type { PivotDefinition as PivotUiDefinition, PivotFieldDefinition as PivotUiFieldDefinition, PivotPanelCallbacks, PivotPanelState, PivotResult as PivotUiResult } from './pivot/types';
import { ChartPanel } from './panels/ChartPanel';
import { PivotPanel } from './panels/PivotPanel';
import { ShapeEditorPanel } from './panels/ShapeEditorPanel';
import { SparklinePanel } from './panels/SparklinePanel';
import { ConditionalFormatPanel } from './panels/ConditionalFormatPanel';
import { DataValidationPanel } from './panels/DataValidationPanel';
import { PrintPanel } from './panels/PrintPanel';
import { HistoryPanel } from './panels/HistoryPanel';
import { DataModelPanel } from './panels/DataModelPanel';

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
  charts: ChartModel[];
  pivot?: PivotModel;
  pivotList?: readonly { id: string; label: string }[];
  activePivotId?: string;
  pivotFieldCatalog?: readonly PivotUiFieldDefinition[];
  pivotResult?: PivotUiResult;
  onShowPivotDetails?: (paths: import('@react-sheets/core-model').PivotSourceRowPath[]) => void;
  pivotPanelState?: PivotPanelState;
  pivotCallbacks?: PivotPanelCallbacks;
  shapes: ShapeModel[];
  sparklines: SparklineModel[];
  conditionalFormats: ConditionalFormatRule[];
  dataValidations: DataValidationRule[];
  historyEntries: readonly HistoryEntry[];
  remoteRevisions: readonly RevisionRecord[];
  tables: readonly WorkbookTableModel[];
  onReadDataRows: (tableId: string, offset?: number, limit?: number) => Promise<TableRowsResponse>;
  onRemoveDataTable: (tableId: string) => Promise<void>;
  onAddChart: (chart: ChartModel) => void;
  onRemoveChart: (id: string) => void;
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
  onAddComment?: (text: string) => void;
  onReplyComment?: (text: string) => void;
  onResolveComment?: () => void;
  onRemoveComment?: () => void;
  onSetHyperlink?: (url: string) => void;
  onRemoveHyperlink?: () => void;
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
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onSetHyperlink,
  onRemoveHyperlink,
}: {
  activeCell: string;
  sheet: CanvasSheetSnapshot;
  onAddComment?: (text: string) => void;
  onReplyComment?: (text: string) => void;
  onResolveComment?: () => void;
  onRemoveComment?: () => void;
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
        hyperlinkUrl={selectedCell?.hyperlink ?? ''}
        onAddComment={onAddComment}
        onReplyComment={onReplyComment}
        onResolveComment={onResolveComment}
        onRemoveComment={onRemoveComment}
        onSetHyperlink={onSetHyperlink}
        onRemoveHyperlink={onRemoveHyperlink}
      />

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
  charts,
  pivot,
  pivotList,
  activePivotId,
  pivotFieldCatalog,
  pivotResult,
  onShowPivotDetails,
  pivotPanelState,
  pivotCallbacks,
  shapes,
  sparklines,
  conditionalFormats,
  dataValidations,
  historyEntries,
  remoteRevisions,
  tables,
  onReadDataRows,
  onRemoveDataTable,
  onAddChart,
  onRemoveChart,
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
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
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
            onAddComment={onAddComment}
            onReplyComment={onReplyComment}
            onResolveComment={onResolveComment}
            onRemoveComment={onRemoveComment}
            onSetHyperlink={onSetHyperlink}
            onRemoveHyperlink={onRemoveHyperlink}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'chart' ? (
          <ChartPanel
            sheetId={sheetId}
            charts={charts}
            defaultRange={selectionText}
            onAddChart={onAddChart}
            onRemoveChart={onRemoveChart}
          />
        ) : null}
        {phase === 'ready' && activePanel === 'pivot' ? (
          <PivotPanel
            pivot={pivot}
            pivotList={pivotList}
            activePivotId={activePivotId}
            fieldCatalog={pivotFieldCatalog}
            result={pivotResult}
            onShowDetails={onShowPivotDetails}
            state={pivotPanelState}
            callbacks={pivotCallbacks}
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
          <PrintPanel onPrint={onPrint} onExportPdf={onExportPdf} />
        ) : null}
        {phase === 'ready' && activePanel === 'history' ? (
          <HistoryPanel entries={historyEntries} remoteRevisions={remoteRevisions} />
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
  hyperlinkUrl: initialHyperlinkUrl,
  onAddComment,
  onReplyComment,
  onResolveComment,
  onRemoveComment,
  onSetHyperlink,
  onRemoveHyperlink,
}: {
  comment?: CellComment;
  commentText: string;
  hyperlinkUrl: string;
  onAddComment?: (text: string) => void;
  onReplyComment?: (text: string) => void;
  onResolveComment?: () => void;
  onRemoveComment?: () => void;
  onSetHyperlink?: (url: string) => void;
  onRemoveHyperlink?: () => void;
}) {
  const [commentText, setCommentText] = useState('');
  const [replyText, setReplyText] = useState('');
  const [hyperlinkUrl, setHyperlinkUrl] = useState('');

  useEffect(() => setCommentText(initialCommentText), [initialCommentText]);
  useEffect(() => setReplyText(''), [comment?.id]);
  useEffect(() => setHyperlinkUrl(initialHyperlinkUrl), [initialHyperlinkUrl]);

  return (
    <Stack gap="md">
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




