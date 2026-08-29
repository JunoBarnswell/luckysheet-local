import React, { createContext, useContext, useEffect, useMemo, useState } from 'react';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import {
  Button,
  Icon,
  Inline,
  RibbonShell,
  Text,
  type RibbonTabId,
} from '@react-sheets/ui-system';
import {
  buildRibbonCommand,
  RIBBON_COMMAND_CATALOG,
  getRibbonGroupDefinition,
  getRibbonCommandDefinition,
  getRibbonCommandDisabledReason,
  isRibbonCommandEnabled,
  RIBBON_TEXT,
  type AppPhase,
  type HomeRibbonState,
  type HomeStyleKey,
  type RibbonCommandActions,
  type RibbonCommandContext,
  type RibbonCommandId,
  type RibbonCommandResult,
  type RibbonMergeOperation,
  type RibbonPivotActions,
  type UiSessionIntent,
  type CompiledFeatureSurfaceSchema,
  type WorkbookSession,
  EXCEL_KEY_TIP_BINDINGS,
  INITIAL_KEY_TIP_STATE,
  keyTipTransition,
  type KeyTipState,
} from '@react-sheets/spreadsheet-app';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { translate, translateRibbonTab, translateRibbonText, type Locale } from '../i18n';
import { homeText } from './home/home-localization';
import { CommandPalette, type CommandPaletteEntry } from './CommandPalette';
import { HomeRibbon, type HomeRibbonCommandOptions } from './HomeRibbon';
import { InsertRibbon } from './InsertRibbon';
import { RibbonTabPresenter } from './RibbonTabPresenter';
import { RibbonLayoutRenderer } from './RibbonLayoutRenderer';
import type { BarcodeSymbology, ChartDrawingPayload, DrawingConnectorType, FormControlType, ShapeDrawingPayload, SheetTableModel, SparklineModel } from '@react-sheets/core-model';

export interface RibbonProps {
  activeTab: RibbonTabId;
  locale: Locale;
  session: WorkbookSession;
  onCommand: (descriptor: CommandDescriptor) => void;
  onSessionIntent: (intent: UiSessionIntent) => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onBeginFormatPainter: (locked?: boolean) => void;
  formatPainterActive?: boolean;
  onMergeCells: (operation: RibbonMergeOperation) => void;
  onUndo: () => void;
  onRedo: () => void;
  onSave: () => void;
  onExportDocument: () => void;
  onImportDocument: () => void;
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
  onAutoSum: (functionName?: 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN') => void;
  onFill: (direction: 'down' | 'up' | 'right' | 'left', mode?: 'copy' | 'series') => void;
  onFreezeAtPrimary: () => void;
  onOpenColumnWidth: () => void;
  onAutoFitColumns: () => void;
  onHideColumns: () => void;
  onUnhideColumns: () => void;
  onOpenDefaultColumnWidth: () => void;
  onOpenRowHeight: () => void;
  onAutoFitRows: () => void;
  onHideRows: () => void;
  onUnhideRows: () => void;
  /** Host-owned Create PivotTable dialog entry point. */
  onCreatePivotDialog?: () => void;
  /** Host-owned selection-aware sort builder. */
  buildSortDescriptor?: (ascending: boolean) => CommandDescriptor | undefined;
  onCreateSheetTable: () => void;
  onOpenTableSettings: () => void;
  onToggleTableOption: (option: 'hasHeaderRow' | 'showFirstColumn' | 'showLastColumn' | 'showBandedRows' | 'showBandedColumns' | 'showFilterButton') => void;
  onConvertActiveTableToRange: () => void;
  onCreateDataSource: () => void;
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
  onSetRecalculationMode: (mode: 'automatic' | 'manual' | 'partial') => void;
  onOpenDefinedNames: () => void;
  onCreateAdvancedSheet: (kind: 'table-sheet' | 'gantt-sheet' | 'report-sheet') => void;
  onApplyBarcode: (symbology?: BarcodeSymbology) => void;
  onCreateCamera: () => void;
  onCaptureScreenshot: () => Promise<void>;
  onCreateFormControl: (type?: FormControlType) => void;
  onApplyCheckbox: () => void;
  onCreateTextBox: () => void;
  onInsertChartType: (type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype']) => void;
  onInsertSparklineType: (type: SparklineModel['type']) => void;
  onInsertShapeType: (type: ShapeDrawingPayload['type']) => void;
  onInsertConnectorType: (type: DrawingConnectorType) => void;
  onTabChange: (tab: RibbonTabId) => void;
  phase: AppPhase;
  activePivot?: { sheetId: string; pivotId: string };
  pivotActions?: RibbonPivotActions;
  activeTableSheet?: { sheetId: string; viewId: string };
  activeGanttSheet?: { sheetId: string; viewId: string };
  activeReportSheet?: { sheetId: string; tableId?: string };
  activeTable?: { sheetId: string; tableId: string; table: SheetTableModel; resizeRange?: SheetTableModel['range'] };
  activeChart?: { sheetId: string; chartId: string };
  activePicture?: { sheetId: string; drawingId: string };
  activeShape?: {
    sheetId: string;
    drawingIds: readonly string[];
    transforms: readonly { drawingId: string; transform: { x: number; y: number; width: number; height: number; rotation?: number } }[];
  };
  activeSparkline?: { sheetId: string; sparklineId: string };
  /** Canonical, selection-derived Home state. All Home controls read this one source. */
  homeState: HomeRibbonState;
  /** FeatureRuntime-compiled availability; command catalog is presentation only. */
  featureSurfaceSchema: CompiledFeatureSurfaceSchema;
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
  ribbonLayoutNodeId,
  ribbonSurfaceId,
  mixed = false,
  iconNode,
  iconOverride,
  trailingNode,
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
  ribbonLayoutNodeId?: string;
  ribbonSurfaceId?: string;
  iconNode?: React.ReactNode;
  iconOverride?: import('@react-sheets/ui-system').IconName;
  trailingNode?: React.ReactNode;
  mixed?: boolean;
}) {
  const locale = useContext(RibbonLocaleContext);
  const layout = useContext(RibbonLayoutContext);
  const definition = getRibbonCommandDefinition(id);
  const enabled = isRibbonCommandEnabled(definition, context);
  const disabledReason = getRibbonCommandDisabledReason(definition, context);
  const label = labelOverride ?? translateRibbonText(locale, definition.labelKey);
  const isNarrow = layout === 'narrow';
  const compactIcon = isNarrow && definition.display === 'small';
  const compactTile = isNarrow && textBelow;
  const active = !mixed && Boolean(definition.active?.(context));
  const mixedLabel = mixed ? `${label} (${homeText(locale, 'mixed')})` : label;
  const keyTip = EXCEL_KEY_TIP_BINDINGS.find((binding) => binding.target.kind === 'command' && binding.target.id === id)?.sequence;
  return (
    <Button
      aria-label={mixedLabel}
      aria-pressed={definition.active ? active : undefined}
      title={disabledReason ?? mixedLabel}
      data-testid={testId}
      data-ribbon-command={id}
      data-ribbon-keytip={keyTip}
      data-ribbon-layout-node={ribbonLayoutNodeId}
      data-ribbon-surface={ribbonSurfaceId}
      data-mixed={mixed || undefined}
      data-disabled-reason={disabledReason}
      disabled={!enabled}
      icon={iconNode ? undefined : iconOverride ?? definition.icon}
      iconNode={iconNode}
      iconOnly={iconOnly || compactIcon}
      onClick={() => {
        const result = buildRibbonCommand(id, context);
        if (result) onExecute(result);
      }}
      size="sm"
      variant={active ? 'primary' : variant}
      className={[
        textBelow ? compactTile ? '!h-6 !min-h-0 !w-6 rounded-none px-0 [&>svg]:!h-3 [&>svg]:!w-3' : '!h-[104px] !min-h-0 min-w-[42px] max-w-[64px] flex-col gap-1 overflow-hidden rounded-none px-1 text-center text-[13px] leading-4 !whitespace-normal break-words [&>svg]:!h-8 [&>svg]:!w-8 [&>svg]:!shrink-0' : undefined,
        className,
        mixed ? 'border border-dashed border-slate-400 bg-slate-50 text-slate-600' : undefined,
      ].filter(Boolean).join(' ')}
    >
      {iconOnly || compactIcon || compactTile ? null : trailingNode ? <Inline gap="none" className="gap-0.5">{label}{trailingNode}</Inline> : label}
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
  onExportDocument,
  onImportDocument,
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
  onOpenRowHeight,
  onAutoFitRows,
  onHideRows,
  onUnhideRows,
  onCreatePivotDialog,
  buildSortDescriptor,
  onCreateSheetTable,
  onOpenTableSettings,
  onToggleTableOption,
  onConvertActiveTableToRange,
  onCreateDataSource,
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
  onCreateCamera,
  onCaptureScreenshot,
  onCreateFormControl,
  onApplyCheckbox,
  onCreateTextBox,
  onInsertChartType,
  onInsertSparklineType,
  onInsertShapeType,
  onTabChange,
  phase,
  activePivot,
  pivotActions,
  activeTableSheet,
  activeGanttSheet,
  activeReportSheet,
  activeTable,
  activeChart,
  activePicture,
  activeShape,
  activeSparkline,
  homeState,
  featureSurfaceSchema,
  session,
  canExecute,
  commandPaletteOpen = false,
  onCloseCommandPalette,
  onInsertConnectorType,
}: RibbonProps) {
  const [keyTipState, setKeyTipState] = useState<KeyTipState>(INITIAL_KEY_TIP_STATE);
  const availableCommandIds = useMemo(() => new Set([
    ...featureSurfaceSchema.ribbon.flatMap((surface) => surface.commandId ? [surface.commandId] : []),
    ...featureSurfaceSchema.contextualTabs.flatMap((surface) => surface.commandId ? [surface.commandId] : []),
  ]), [featureSurfaceSchema]);
  const isSurfaceAvailable = (id: RibbonCommandId): boolean => availableCommandIds.has(id);
  useEffect(() => {
    const featurePhase = session.getFeatureLifecyclePhase();
    if (featurePhase === 'ready') session.advanceFeatureLifecycle('rendered');
    if (session.getFeatureLifecyclePhase() !== 'rendered') return;

    let cancelled = false;
    const advanceSteady = (): void => {
      if (!cancelled) session.advanceFeatureLifecycle('steady');
    };
    const hasFrameScheduler = typeof globalThis.requestAnimationFrame === 'function';
    const scheduled = hasFrameScheduler
      ? globalThis.requestAnimationFrame(advanceSteady)
      : globalThis.setTimeout(advanceSteady, 0);
    return () => {
      cancelled = true;
      if (hasFrameScheduler) globalThis.cancelAnimationFrame(scheduled as number);
      else globalThis.clearTimeout(scheduled);
    };
  }, [session, featureSurfaceSchema]);
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent): void => {
      const target = event.target as HTMLElement | null;
      const editingSurface = target?.closest('input, textarea, [contenteditable="true"]');
      if (!keyTipState.active) {
        if (editingSurface) return;
        if (event.key !== 'Alt' && !(event.key === 'F10' && !event.shiftKey)) return;
        event.preventDefault();
        event.stopPropagation();
        setKeyTipState({ active: true, prefix: '' });
        return;
      }
      if (event.key === 'Alt' || (event.key === 'F10' && !event.shiftKey)) {
        event.preventDefault();
        event.stopPropagation();
        setKeyTipState(INITIAL_KEY_TIP_STATE);
        return;
      }
      const transition = keyTipTransition(keyTipState, event.key);
      if (event.key === 'Escape' || transition.state !== keyTipState || transition.action) {
        event.preventDefault();
        event.stopPropagation();
      }
      setKeyTipState(transition.state);
      if (!transition.action) return;
      if (transition.action.kind === 'tab') {
        onTabChange(transition.action.id as RibbonTabId);
        return;
      }
      const sequence = `${keyTipState.prefix}${event.key.toLocaleUpperCase()}`;
      const node = document.querySelector<HTMLElement>(`[data-ribbon-keytip="${sequence}"]`);
      if (node) {
        node.click();
        return;
      }
      // Menu members are rendered in a portal only after their menu opens.
      // Resolve those commands through the same catalog builder instead of
      // silently dropping a valid KeyTip when its visual member is not mounted.
      const commandId = transition.action.id as RibbonCommandId;
      if (RIBBON_COMMAND_CATALOG.some((definition) => definition.id === commandId) && isSurfaceAvailable(commandId)) {
        const result = buildRibbonCommand(commandId, catalogContext);
        if (result) executeCatalogResult(result);
      }
    };
    document.addEventListener('keydown', handleKeyDown, true);
    return () => document.removeEventListener('keydown', handleKeyDown, true);
  }, [keyTipState, onTabChange, availableCommandIds]);
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
    onExportDocument,
    onImportDocument,
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
    onMerge: onMergeCells,
    onFill,
    onFreezeAtPrimary,
    onCreateSheetTable,
    onOpenTableSettings,
    onToggleTableOption,
    onConvertActiveTableToRange,
    onCreateDataSource,
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
    onCreateCamera,
    onCaptureScreenshot,
    onCreateFormControl,
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
    pivotActions,
    activeTableSheet,
    activeGanttSheet,
    activeReportSheet,
    activeTable,
    activeChart,
    activePicture,
    activeShape,
    activeSparkline,
    actions: catalogActions,
    dispatchSessionIntent: onSessionIntent,
  };
  const executeCatalogResult = (result: RibbonCommandResult) => {
    if (result.type === 'command') onCommand(result.descriptor);
    else if (result.type === 'intent') onSessionIntent(result.intent);
    else {
      try {
        const pending = result.invoke();
        if (pending) void Promise.resolve(pending).catch((error) => session.notify(error instanceof Error ? error.message : 'Ribbon action failed'));
      } catch (error) {
        session.notify(error instanceof Error ? error.message : 'Ribbon action failed');
      }
    }
  };
  const renderHomeCommand = (id: RibbonCommandId, options: HomeRibbonCommandOptions = {}) => !isSurfaceAvailable(id) ? null : (
    <CatalogButton
      key={id}
      id={id}
      context={catalogContext}
      onExecute={executeCatalogResult}
      iconOnly={options.iconOnly}
      iconNode={options.iconNode}
      iconOverride={options.iconOverride}
      trailingNode={options.trailingNode}
      labelOverride={options.labelOverride}
      ribbonLayoutNodeId={options.ribbonLayoutNodeId}
      textBelow={options.tile}
      className={options.className}
      ribbonSurfaceId={options.ribbonSurfaceId}
      testId={options.testId}
    />
  );

  const commandPaletteEntries: CommandPaletteEntry[] = RIBBON_COMMAND_CATALOG.filter((definition) => isSurfaceAvailable(definition.id)).map((definition) => {
    const result = isRibbonCommandEnabled(definition, catalogContext) ? buildRibbonCommand(definition.id, catalogContext) : undefined;
    const groups = [...new Set(definition.placements.map((placement) => translateRibbonText(locale, getRibbonGroupDefinition(placement.group).labelKey)))];
    return {
      id: definition.id,
      label: translateRibbonText(locale, definition.labelKey),
      group: groups.join(' / '),
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
          ...(activeTableSheet ? ['tableSheetDesign'] as const : []),
          ...(activeGanttSheet ? ['ganttTask', 'ganttProject', 'ganttView', 'ganttFormat'] as const : []),
          ...(activeReportSheet ? ['reportSheetDesign'] as const : []),
          ...(activeTable ? ['tableDesign'] as const : []),
          ...(activeChart ? ['chartDesign', 'chartFormat'] as const : []),
          ...(activePicture ? ['pictureFormat'] as const : []),
          ...(activeShape ? ['shapeFormat'] as const : []),
          ...(activeSparkline ? ['sparklineDesign'] as const : []),
        ]}
        disabled={disabled}
        onFileEntry={() => onSessionIntent({ type: 'backstage.open', panel: 'info' })}
        onTabChange={onTabChange}
        tabLabel={(tab) => translateRibbonTab(locale, tab)}
        keyTipState={keyTipState}
        keyTipBindings={EXCEL_KEY_TIP_BINDINGS}
        status={(
          <>
            <Icon name="cloud-check" size="sm" className={disabled ? 'text-slate-300' : 'text-emerald-500'} />
            <Text size="xs" tone="muted">{translate(locale, 'engineConnected')}</Text>
          </>
        )}
      >
        {(layout) => (
          <RibbonLayoutContext.Provider value={layout.mode}>
        {activeTab === 'pageLayout' || activeTab === 'formulas' || activeTab === 'data'
          ? <RibbonLayoutRenderer tab={activeTab} locale={locale} layout={layout} renderCommand={renderHomeCommand} featureSurfaceSchema={featureSurfaceSchema} />
          : activeTab !== 'home' && activeTab !== 'insert'
            ? <RibbonTabPresenter tab={activeTab} locale={locale} layout={layout} renderCommand={renderHomeCommand} featureSurfaceSchema={featureSurfaceSchema} />
            : null}

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
            onOpenRowHeight={onOpenRowHeight}
            onAutoFitRows={onAutoFitRows}
            onHideRows={onHideRows}
            onUnhideRows={onUnhideRows}
            onUnhideColumns={onUnhideColumns}
            renderCommand={renderHomeCommand}
            featureSurfaceSchema={featureSurfaceSchema}
          />
        ) : null}
        {activeTab === 'insert' ? (
      <InsertRibbon
            locale={locale}
            layout={layout}
            disabled={disabled}
            renderCommand={renderHomeCommand}
            featureSurfaceSchema={featureSurfaceSchema}
        onInsertChart={onInsertChartType}
            onInsertSparkline={onInsertSparklineType}
            onInsertShape={onInsertShapeType}
            onInsertConnector={onInsertConnectorType}
            onInsertFormControl={onCreateFormControl}
            onOpenMoreCharts={() => onSessionIntent({ type: 'dialog.open', dialog: 'recommended-charts' })}
            canExecute={(commandId) => !canExecute || canExecute(commandId)}
            canInsertConnector={(activeShape?.drawingIds.length ?? 0) >= 2}
          />
        ) : null}

          </RibbonLayoutContext.Provider>
        )}
      </RibbonShell>
      <CommandPalette commands={commandPaletteEntries} onClose={onCloseCommandPalette} open={commandPaletteOpen} />
    </RibbonLocaleContext.Provider>
  );
}
