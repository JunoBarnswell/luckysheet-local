import React, { createContext, useContext } from 'react';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import {
  Box,
  Button,
  ColorPicker,
  Divider,
  DropdownMenu,
  Icon,
  Inline,
  RibbonShell,
  Select,
  Stack,
  Text,
  TextInput,
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
  type RibbonGroupId,
  type UiSessionIntent,
} from '@react-sheets/spreadsheet-app';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { localizeText, translate, translateRibbonTab, translateRibbonText, type Locale } from '../i18n';
import { homeText } from './home/home-localization';
import { CommandPalette, type CommandPaletteEntry } from './CommandPalette';
import { HomeRibbon, type HomeRibbonCommandOptions } from './HomeRibbon';

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
  onCreatePivot: () => CommandDescriptor | undefined;
  onCreateChart: () => CommandDescriptor | undefined;
  onCreateSparkline: () => CommandDescriptor | undefined;
  onCreateShape: () => CommandDescriptor | undefined;
  onBringDrawingForward: () => CommandDescriptor | undefined;
  onSendDrawingBackward: () => CommandDescriptor | undefined;
  onRemoveDrawing: () => CommandDescriptor | undefined;
  onCreateSheetTable: () => void;
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
  onTabChange: (tab: RibbonTabId) => void;
  phase: AppPhase;
  activePivot?: { sheetId: string; pivotId: string };
  /** Canonical, selection-derived Home state. All Home controls read this one source. */
  homeState: HomeRibbonState;
  canExecute?: (commandId: string, params?: unknown) => boolean;
  commandPaletteOpen?: boolean;
  onCloseCommandPalette: () => void;
}

const RibbonLocaleContext = createContext<Locale>('en-US');
type RibbonLayoutMode = 'wide' | 'compact' | 'narrow';
const RibbonLayoutContext = createContext<RibbonLayoutMode>('wide');

function RibbonGroup({ children, label, group, className }: { children: React.ReactNode; label?: string; group?: RibbonGroupId; className?: string }) {
  const locale = useContext(RibbonLocaleContext);
  const layout = useContext(RibbonLayoutContext);
  const definition = group ? getRibbonGroupDefinition(group) : undefined;
  const localizedLabel = definition ? translateRibbonText(locale, definition.labelKey) : localizeText(locale, label ?? '');
  const collapsed = definition
    ? layout === 'narrow' ? definition.priority > 10 : layout === 'compact' ? definition.priority > 40 : false
    : false;
  if (collapsed) {
    return (
      <DropdownMenu
        align="left"
        trigger={(
          <Button
            aria-label={localizedLabel}
            title={localizedLabel}
            size="sm"
            variant="ghost"
            icon="more-horizontal"
            className="max-w-[9rem] justify-start"
          >
            {localizedLabel}
          </Button>
        )}
      >
        <Stack gap="xs" className="min-w-[12rem] p-1">
          {children}
        </Stack>
      </DropdownMenu>
    );
  }
  return (
    <Stack gap="xs" className={`h-[98px] min-w-0 shrink-0 justify-between ${className ?? ''}`}>
      <Inline gap="xs" className="min-h-8 flex-wrap items-start">
        {children}
      </Inline>
      <Text size="xs" tone="subtle" className="pointer-events-none text-center text-[10px] font-semibold uppercase tracking-wider text-slate-400">
        {localizedLabel}
      </Text>
    </Stack>
  );
}

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
  onCreatePivot,
  onCreateChart,
  onCreateSparkline,
  onCreateShape,
  onBringDrawingForward,
  onSendDrawingBackward,
  onRemoveDrawing,
  onCreateSheetTable,
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
  onTabChange,
  phase,
  activePivot,
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
    onCreatePivot,
    onCreateChart,
    onCreateSparkline,
    onCreateShape,
    onBringDrawingForward,
    onSendDrawingBackward,
    onRemoveDrawing,
    onCreateSheetTable,
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
  };
  const catalogContext: RibbonCommandContext = {
    phase,
    disabled,
    cellStyle,
    canExecute,
    buildSortDescriptor,
    openCreatePivotDialog: onCreatePivotDialog,
    activePivot,
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
        contextualTabs={activePivot ? ['pivotAnalyze', 'pivotDesign'] : []}
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
        {activeTab === 'file' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="workbook">
              <CatalogButton id="save" context={catalogContext} onExecute={executeCatalogResult} variant="secondary" />
              <CatalogButton id="exportXlsx" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="importXlsx" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'settings' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="settings">
              <CatalogButton id="settings" context={catalogContext} onExecute={executeCatalogResult} variant="secondary" />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'automate' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="scripts">
              <CatalogButton id="openAutomate" context={catalogContext} onExecute={executeCatalogResult} variant="primary" />
              <CatalogButton id="runSampleScript" context={catalogContext} onExecute={executeCatalogResult} variant="outline" />
              <CatalogButton id="startRecording" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="stopRecording" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'pageLayout' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="pageSetup">
              <CatalogButton id="pageSetup" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="setPrintArea" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="clearPrintArea" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="printTitleRows" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="printTitleColumns" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />
            <RibbonGroup group="scaleToFit">
              <CatalogButton id="setScale100" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />
            <RibbonGroup group="sheetOptions">
              <CatalogButton id="viewGridlines" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="printGridlines" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="viewHeadings" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="printHeadings" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'formulas' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="calculation">
              <CatalogButton id="calculateNow" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="calculationAutomatic" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="calculationManual" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="goalSeek" context={catalogContext} onExecute={executeCatalogResult} variant="outline" />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />
            <RibbonGroup group="formulaAudit">
              <CatalogButton id="tracePrecedents" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="traceDependents" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="removeArrows" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="showFormulas" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="errorChecking" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="evaluateFormula" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />
            <RibbonGroup group="definedNames">
              <CatalogButton id="definedNames" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'pivotAnalyze' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="pivotAnalyze">
              <CatalogButton id="pivotRefresh" context={catalogContext} onExecute={executeCatalogResult} variant="primary" />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'pivotDesign' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="pivotDesign">
              <CatalogButton id="pivotFieldList" context={catalogContext} onExecute={executeCatalogResult} variant="primary" />
            </RibbonGroup>
          </Inline>
        ) : null}

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
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="tablesPivots">
              <CatalogButton id="pivotTable" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="quickPivot" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="chartsVisuals">
              <CatalogButton id="chartBuilder" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="columnChart" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="sparkline" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="quickSparkline" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="illustrations">
              <CatalogButton id="shapesLines" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="rectangle" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="bringDrawingForward" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="sendDrawingBackward" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="removeDrawing" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="functions">
              <CatalogButton id="insertFunction" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="cells">
              <CatalogButton id="insertRow" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="insertColumn" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="deleteRow" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="deleteColumn" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'data' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="sortFilter">
              <CatalogButton id="sortAscending" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="sortDescending" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="customSort" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="dataTools">
              <CatalogButton id="dataModel" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="createDataTable" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="formatAsTable" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="totalRow" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="dataValidation" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="filterSelection" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="clearFilter" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="outline">
              <CatalogButton id="groupRows" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="ungroupRows" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="groupColumns" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="ungroupColumns" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="showLevel1" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="showLevel2" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="showLevel3" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="subtotal" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="removeDuplicates" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="textToColumns" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="findTransform">
              <CatalogButton id="findReplace" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="goTo" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="transpose" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="flipHorizontal" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="flipVertical" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="splitByDelimiter" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'review' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="comments">
              <CatalogButton id="newComment" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="resolveComment" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="showComments" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup group="notesLinks">
              <CatalogButton id="newNote" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="insertLink" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup group="protection">
              <CatalogButton id="protectSelection" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="unprotect" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup group="historyAudit">
              <CatalogButton id="revisionLog" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'view' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="freezePanes">
              <CatalogButton id="freezeTopRow" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="freezeFirstColumn" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="freezeAtSelection" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="unfreezeAll" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="zoom">
              <CatalogButton id="zoomIn" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="zoomOut" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="zoomReset" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="commandPalette" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="printLayout">
              <CatalogButton id="printPdf" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="appearanceFiles">
              <CatalogButton id="bandedRows" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="exportXlsxView" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="importXlsxView" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}
          </RibbonLayoutContext.Provider>
        )}
      </RibbonShell>
      <CommandPalette commands={commandPaletteEntries} onClose={onCloseCommandPalette} open={commandPaletteOpen} />
    </RibbonLocaleContext.Provider>
  );
}
