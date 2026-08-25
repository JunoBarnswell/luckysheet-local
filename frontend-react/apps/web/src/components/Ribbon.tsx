import React, { createContext, useContext } from 'react';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import {
  Button,
  Icon,
  RibbonShell,
  Text,
  type RibbonTabId,
} from '@react-sheets/ui-system';
import {
  SAMPLE_AUTOMATION_SCRIPT,
  buildRibbonCommand,
  RIBBON_COMMAND_CATALOG,
  getRibbonGroupDefinition,
  getRibbonCommandDefinition,
  isRibbonCommandEnabled,
  RIBBON_TEXT,
  type AppPhase,
  type HomeRibbonState,
  type HomeStyleKey,
  type RibbonCommandActions,
  type RibbonCommandContext,
  type RibbonCommandId,
  type RibbonCommandResult,
  type UiSessionIntent,
} from '@react-sheets/spreadsheet-app';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { translate, translateRibbonTab, translateRibbonText, type Locale } from '../i18n';
import { homeText } from './home/home-localization';
import { CommandPalette, type CommandPaletteEntry } from './CommandPalette';
import { HomeRibbon, type HomeRibbonCommandOptions } from './HomeRibbon';
import { InsertRibbon } from './InsertRibbon';
import { RibbonTabPresenter } from './RibbonTabPresenter';
import type { BarcodeSymbology, ChartDrawingPayload, DataChartPlotType, FormControlType, ShapeDrawingPayload, SheetTableModel, SparklineModel } from '@react-sheets/core-model';

export interface RibbonProps {
  activeTab: RibbonTabId;
  locale: Locale;
  onCommand: (descriptor: CommandDescriptor) => void;
  onSessionIntent: (intent: UiSessionIntent) => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onBeginFormatPainter: (locked?: boolean) => void;
  formatPainterActive?: boolean;
  onMergeCells: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onSave: () => void;
  onExportXlsx: () => void;
  onImportXlsx: () => void;
  onRecalculate: () => void;
  onTracePrecedents: () => void;
  onTraceDependents: () => void;
  onRemoveArrows: () => void;
  onToggleShowFormulas: () => void;
  onScanFormulaErrors: () => void;
  onEvaluateFormula: () => void;
  onOpenPrintLayout: () => void;
  onSetPrintArea: () => void;
  onClearPrintArea: () => void;
  onSetPrintTitleRows: () => void;
  onSetPrintTitleColumns: () => void;
  onSetPrintScale: (scale: number) => void;
  onToggleViewGridlines: () => void;
  onTogglePrintGridlines: () => void;
  onToggleViewHeadings: () => void;
  onTogglePrintHeadings: () => void;
  onAutoSum: () => void;
  onFill: (direction: 'down' | 'right') => void;
  onFreezeAtPrimary: () => void;
  onOpenColumnWidth: () => void;
  onAutoFitColumns: () => void;
  onHideColumns: () => void;
  onUnhideColumns: () => void;
  onOpenDefaultColumnWidth: () => void;
  /** Host-owned Create PivotTable dialog entry point. */
  onCreatePivotDialog?: () => void;
  /** Host-owned selection-aware sort builder. */
  buildSortDescriptor?: (ascending: boolean) => CommandDescriptor | undefined;
  onCreateSheetTable: () => void;
  onOpenTableSettings: () => void;
  onToggleTableOption: (option: 'hasHeaderRow' | 'showFirstColumn' | 'showLastColumn' | 'showBandedRows' | 'showBandedColumns' | 'showFilterButton') => void;
  onConvertActiveTableToRange: () => void;
  onCreateDataTable: () => void;
  onToggleSheetTableTotalRow: () => CommandDescriptor | undefined;
  onApplyFilterSelection: () => CommandDescriptor | undefined;
  onClearFilter: () => CommandDescriptor | undefined;
  onGroupRows: () => CommandDescriptor | undefined;
  onUngroupRows: () => CommandDescriptor | undefined;
  onGroupColumns: () => CommandDescriptor | undefined;
  onUngroupColumns: () => CommandDescriptor | undefined;
  onSubtotal: () => CommandDescriptor | undefined;
  onRemoveDuplicates: () => CommandDescriptor | undefined;
  onTextToColumns: () => CommandDescriptor | undefined;
  onResolveComment: () => void;
  onProtectSelection: () => void;
  onUnprotectSelection: () => void;
  onShowOutlineLevel: (level: 1 | 2 | 3) => void;
  onTransposeSelection: () => void;
  onFlipSelection: (axis: 'h' | 'v') => void;
  onSplitByDelimiter: () => void;
  onToggleBandedRows: () => void;
  onSetRecalculationMode: (mode: 'automatic' | 'manual') => void;
  onOpenDefinedNames: () => void;
  onCreateAdvancedSheet: (kind: 'table-sheet' | 'gantt-sheet' | 'report-sheet') => void;
  onApplyBarcode: (symbology?: BarcodeSymbology) => void;
  onCreateDataChart: (type?: DataChartPlotType) => void;
  onCreateCamera: () => void;
  onCreateFormControl: (type?: FormControlType) => void;
  onApplyCheckbox: () => void;
  onCreateTextBox: () => void;
  onInsertChartType: (type: ChartDrawingPayload['chartType']) => void;
  onInsertSparklineType: (type: SparklineModel['type']) => void;
  onInsertShapeType: (type: ShapeDrawingPayload['type']) => void;
  onTabChange: (tab: RibbonTabId) => void;
  phase: AppPhase;
  activePivot?: { sheetId: string; pivotId: string };
  activeTableSheet?: { sheetId: string; viewId: string };
  activeGanttSheet?: { sheetId: string; viewId: string };
  activeReportSheet?: { sheetId: string; tableId?: string };
  activeTable?: { sheetId: string; tableId: string; table: SheetTableModel; resizeRange?: SheetTableModel['range'] };
  activeChart?: { sheetId: string; chartId: string };
  /** Canonical, selection-derived Home state. All Home controls read this one source. */
  homeState: HomeRibbonState;
  canExecute?: (commandId: string, params?: unknown) => boolean;
  commandPaletteOpen?: boolean;
  onCloseCommandPalette: () => void;
}

const RibbonLocaleContext = createContext<Locale>('en-US');
type RibbonLayoutMode = 'wide' | 'compact' | 'narrow';
const RibbonLayoutContext = createContext<RibbonLayoutMode>('wide');

function CatalogButton({
  id,
  context,
  onExecute,
  iconOnly = false,
  textBelow = false,
  labelOverride,
  variant = 'ghost',
  className,
  testId,
  mixed = false,
}: {
  id: RibbonCommandId;
  context: RibbonCommandContext;
  onExecute: (result: RibbonCommandResult) => void;
  iconOnly?: boolean;
  textBelow?: boolean;
  labelOverride?: string;
  variant?: 'danger' | 'ghost' | 'outline' | 'primary' | 'secondary' | 'soft';
  className?: string;
  testId?: string;
  mixed?: boolean;
}) {
  const locale = useContext(RibbonLocaleContext);
  const layout = useContext(RibbonLayoutContext);
  const definition = getRibbonCommandDefinition(id);
  const enabled = isRibbonCommandEnabled(definition, context);
  const label = labelOverride ?? translateRibbonText(locale, definition.labelKey);
  const compactIcon = layout !== 'wide' && definition.display === 'small';
  const active = !mixed && Boolean(definition.active?.(context));
  const mixedLabel = mixed ? `${label} (${homeText(locale, 'mixed')})` : label;
  return (
    <Button
      aria-label={mixedLabel}
      aria-pressed={definition.active ? active : undefined}
      title={mixedLabel}
      data-testid={testId}
      data-mixed={mixed || undefined}
      disabled={!enabled}
      icon={definition.icon}
      iconOnly={iconOnly || compactIcon}
      onClick={() => {
        const result = buildRibbonCommand(id, context);
        if (result) onExecute(result);
      }}
      size="sm"
      variant={active ? 'primary' : variant}
      className={[
        textBelow ? '!h-[64px] !min-h-0 !w-[68px] flex-col gap-0 rounded-none px-1 text-[10px] leading-3 [&>svg]:!h-6 [&>svg]:!w-6' : undefined,
        className,
        mixed ? 'border border-dashed border-slate-400 bg-slate-50 text-slate-600' : undefined,
      ].filter(Boolean).join(' ')}
    >
      {iconOnly || compactIcon ? null : label}
    </Button>
  );
}

export function Ribbon({
  activeTab,
  locale,
  onCommand,
  onSessionIntent,
  onCopy,
  onCut,
  onPaste,
  onBeginFormatPainter,
  formatPainterActive = false,
  onMergeCells,
  onUndo,
  onRedo,
  onSave,
  onExportXlsx,
  onImportXlsx,
  onRecalculate,
  onTracePrecedents,
  onTraceDependents,
  onRemoveArrows,
  onToggleShowFormulas,
  onScanFormulaErrors,
  onEvaluateFormula,
  onOpenPrintLayout,
  onSetPrintArea,
  onClearPrintArea,
  onSetPrintTitleRows,
  onSetPrintTitleColumns,
  onSetPrintScale,
  onToggleViewGridlines,
  onTogglePrintGridlines,
  onToggleViewHeadings,
  onTogglePrintHeadings,
  onAutoSum,
  onFill,
  onFreezeAtPrimary,
  onOpenColumnWidth,
  onAutoFitColumns,
  onHideColumns,
  onUnhideColumns,
  onOpenDefaultColumnWidth,
  onCreatePivotDialog,
  buildSortDescriptor,
  onCreateSheetTable,
  onOpenTableSettings,
  onToggleTableOption,
  onConvertActiveTableToRange,
  onCreateDataTable,
  onToggleSheetTableTotalRow,
  onApplyFilterSelection,
  onClearFilter,
  onGroupRows,
  onUngroupRows,
  onGroupColumns,
  onUngroupColumns,
  onSubtotal,
  onRemoveDuplicates,
  onTextToColumns,
  onResolveComment,
  onProtectSelection,
  onUnprotectSelection,
  onShowOutlineLevel,
  onTransposeSelection,
  onFlipSelection,
  onSplitByDelimiter,
  onToggleBandedRows,
  onSetRecalculationMode,
  onOpenDefinedNames,
  onCreateAdvancedSheet,
  onApplyBarcode,
  onCreateDataChart,
  onCreateCamera,
  onCreateFormControl,
  onApplyCheckbox,
  onCreateTextBox,
  onInsertChartType,
  onInsertSparklineType,
  onInsertShapeType,
  onTabChange,
  phase,
  activePivot,
  activeTableSheet,
  activeGanttSheet,
  activeReportSheet,
  activeTable,
  activeChart,
  homeState,
  canExecute,
  commandPaletteOpen = false,
  onCloseCommandPalette,
}: RibbonProps) {
  const disabled = phase !== 'ready';
  const cellStyle = homeState.style;
  const canFormat = (style: Record<string, unknown>): boolean => !disabled && homeState.canFormat
    && (!canExecute || canExecute('sheet.style.set', { style }));
  const emitStyle = (style: Record<string, unknown>): void => {
    if (canFormat(style)) onCommand({ commandId: 'sheet.style.set', params: { style } });
  };
  const isMixed = (field: HomeStyleKey): boolean => homeState.mixedStyleKeys.includes(field);
  const catalogActions: RibbonCommandActions = {
    onCopy,
    onCut,
    onPaste,
    onUndo,
    onRedo,
    onSave,
    onExportXlsx,
    onImportXlsx,
    onRecalculate,
    onTracePrecedents,
    onTraceDependents,
    onRemoveArrows,
    onToggleShowFormulas,
    onScanFormulaErrors,
    onEvaluateFormula,
    onOpenPrintLayout,
    onSetPrintArea,
    onClearPrintArea,
    onSetPrintTitleRows,
    onSetPrintTitleColumns,
    onSetPrintScale,
    onToggleViewGridlines,
    onTogglePrintGridlines,
    onToggleViewHeadings,
    onTogglePrintHeadings,
    onAutoSum,
    onFill,
    onFreezeAtPrimary,
    onCreateSheetTable,
    onOpenTableSettings,
    onToggleTableOption,
    onConvertActiveTableToRange,
    onCreateDataTable,
    onToggleSheetTableTotalRow,
    onApplyFilterSelection,
    onClearFilter,
    onGroupRows,
    onUngroupRows,
    onGroupColumns,
    onUngroupColumns,
    onSubtotal,
    onRemoveDuplicates,
    onTextToColumns,
    onResolveComment,
    onProtectSelection,
    onUnprotectSelection,
    onShowOutlineLevel,
    onTransposeSelection,
    onFlipSelection,
    onSplitByDelimiter,
    onToggleBandedRows,
    onSetRecalculationMode,
    onOpenDefinedNames,
    onCreateAdvancedSheet,
    onApplyBarcode,
    onCreateDataChart,
    onCreateCamera,
    onCreateFormControl: () => onCreateFormControl('button'),
    onApplyCheckbox,
    onCreateTextBox,
  };
  const catalogContext: RibbonCommandContext = {
    phase,
    disabled,
    cellStyle,
    canExecute,
    buildSortDescriptor,
    openCreatePivotDialog: onCreatePivotDialog,
    activePivot,
    activeTableSheet,
    activeGanttSheet,
    activeReportSheet,
    activeTable,
    activeChart,
    actions: catalogActions,
    dispatchSessionIntent: onSessionIntent,
    sampleAutomationScript: SAMPLE_AUTOMATION_SCRIPT,
  };
  const executeCatalogResult = (result: RibbonCommandResult) => {
    if (result.type === 'command') onCommand(result.descriptor);
    else if (result.type === 'intent') onSessionIntent(result.intent);
    else result.invoke();
  };
  const renderHomeCommand = (id: RibbonCommandId, options: HomeRibbonCommandOptions = {}) => (
    <CatalogButton
      key={id}
      id={id}
      context={catalogContext}
      onExecute={executeCatalogResult}
      iconOnly={options.iconOnly}
      textBelow={options.tile}
      className={options.className}
      testId={options.testId}
    />
  );

  const commandPaletteEntries: CommandPaletteEntry[] = RIBBON_COMMAND_CATALOG.map((definition) => {
    const result = isRibbonCommandEnabled(definition, catalogContext) ? buildRibbonCommand(definition.id, catalogContext) : undefined;
    const group = getRibbonGroupDefinition(definition.group);
    return {
      id: definition.id,
      label: translateRibbonText(locale, definition.labelKey),
      group: translateRibbonText(locale, group.labelKey),
      keywords: [definition.id, definition.commandId ?? ''],
      tip: definition.tooltipKey ? translateRibbonText(locale, definition.tooltipKey) : undefined,
      commandId: definition.commandId,
      enabled: result !== undefined,
      execute: () => { if (result) executeCatalogResult(result); },
    };
  });

  return (
    <RibbonLocaleContext.Provider value={locale}>
      <RibbonShell
        activeTab={activeTab}
        contextualTabs={[
          ...(activePivot ? ['pivotAnalyze', 'pivotDesign'] as const : []),
          ...(activeTableSheet ? ['tableSheetDesign'] as const : []),
          ...(activeGanttSheet ? ['ganttTask', 'ganttProject', 'ganttView', 'ganttFormat'] as const : []),
          ...(activeReportSheet ? ['reportSheetDesign'] as const : []),
          ...(activeTable ? ['tableDesign'] as const : []),
          ...(activeChart ? ['chartDesign', 'chartFormat'] as const : []),
        ]}
        disabled={disabled}
        onFileEntry={() => onSessionIntent({ type: 'backstage.open', panel: 'info' })}
        onTabChange={onTabChange}
        tabLabel={(tab) => translateRibbonTab(locale, tab)}
        status={(
          <>
            <Icon name="cloud-check" size="sm" className={disabled ? 'text-slate-300' : 'text-emerald-500'} />
            <Text size="xs" tone="muted">{translate(locale, 'engineConnected')}</Text>
          </>
        )}
      >
        {(layout) => (
          <RibbonLayoutContext.Provider value={layout.mode}>
        {activeTab !== 'home' && activeTab !== 'insert' ? <RibbonTabPresenter tab={activeTab} locale={locale} layout={layout} renderCommand={renderHomeCommand} /> : null}

        {activeTab === 'home' ? (
          <HomeRibbon
            context={catalogContext}
            disabled={disabled}
            homeState={homeState}
            formatPainterActive={formatPainterActive}
            layout={layout}
            locale={locale}
            onAutoFitColumns={onAutoFitColumns}
            onBeginFormatPainter={onBeginFormatPainter}
            onEmitStyle={emitStyle}
            onHideColumns={onHideColumns}
            onMergeCells={onMergeCells}
            onOpenColumnWidth={onOpenColumnWidth}
            onOpenDefaultColumnWidth={onOpenDefaultColumnWidth}
            onUnhideColumns={onUnhideColumns}
            renderCommand={renderHomeCommand}
          />
        ) : null}
        {activeTab === 'insert' ? (
      <InsertRibbon
            locale={locale}
            layout={layout}
            disabled={disabled}
            renderCommand={renderHomeCommand}
        onInsertChart={onInsertChartType}
        onInsertDataChart={(type) => onCreateDataChart(type)}
        onInsertBarcode={(symbology) => onApplyBarcode(symbology)}
            onInsertSparkline={onInsertSparklineType}
            onInsertShape={onInsertShapeType}
            onInsertFormControl={onCreateFormControl}
          />
        ) : null}

          </RibbonLayoutContext.Provider>
        )}
      </RibbonShell>
      <CommandPalette commands={commandPaletteEntries} onClose={onCloseCommandPalette} open={commandPaletteOpen} />
    </RibbonLocaleContext.Provider>
  );
}
