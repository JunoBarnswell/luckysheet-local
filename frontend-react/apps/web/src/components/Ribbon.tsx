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
        licenseEntry={(
          <Button
            className="h-full rounded-none px-2 text-[11px] font-normal text-slate-600 hover:bg-[#f3f8f4] hover:text-[#217345]"
            onClick={() => onSessionIntent({ type: 'backstage.open', panel: 'options' })}
            size="sm"
            variant="ghost"
          >
            在线表格编辑器授权信息
          </Button>
        )}
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
          <Inline gap="none" className="flex-nowrap items-start overflow-hidden" data-testid="home-ribbon-groups">
            <RibbonGroup className="!w-[141px] overflow-hidden" group="clipboard">
              <Stack gap="none" className="w-full">
                <Inline gap="none" className="h-[68px] flex-nowrap items-start">
                  <Stack gap="none" className="w-[58px] shrink-0">
                    <Inline gap="none" className="h-8 flex-nowrap items-start">
                      <CatalogButton id="undo" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
                      <CatalogButton id="redo" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
                    </Inline>
                    <Inline gap="none" className="h-8 flex-nowrap items-start">
                      <CatalogButton id="cut" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
                      <CatalogButton id="copy" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
                    </Inline>
                  </Stack>
                  <CatalogButton id="paste" context={catalogContext} onExecute={executeCatalogResult} textBelow labelOverride={homeText(locale, 'pasteAll')} className="!w-[58px]" />
                  <Button
                    size="sm"
                    variant={formatPainterActive ? 'secondary' : 'ghost'}
                    icon="palette"
                    iconOnly
                    disabled={disabled || !homeState.canFormat}
                    data-testid="home-format-painter"
                    aria-pressed={formatPainterActive}
                    title={homeText(locale, 'formatPainterHint')}
                    onClick={() => onBeginFormatPainter(false)}
                    onDoubleClick={() => onBeginFormatPainter(true)}
                  />
                  <CatalogButton id="pasteSpecial" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
                </Inline>
              </Stack>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />

            <RibbonGroup className="!w-[216px] overflow-hidden" group="font">
              <Stack gap="xs" className="w-[205px] pt-3">
              <Inline gap="xs" className="h-8 flex-nowrap items-start">
              <Box className="w-[96px] shrink-0">
                <Select
                  aria-label={homeText(locale, 'fontFamily')}
                  className="w-full"
                  disabled={!homeState.canFormat || disabled}
                  sizeVariant="sm"
                  value={isMixed('fontFamily') ? '__mixed__' : cellStyle.fontFamily ?? 'Microsoft YaHei'}
                  onChange={(event) => {
                    if (event.target.value !== '__mixed__') emitStyle({ fontFamily: event.target.value });
                  }}
                >
                  {isMixed('fontFamily') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
                  <option value="Microsoft YaHei">微软雅黑</option>
                  <option value="Arial">Arial</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Segoe UI">Segoe UI</option>
                  <option value="Times New Roman">Times New Roman</option>
                </Select>
              </Box>
              <TextInput
                aria-label={homeText(locale, 'fontSize')}
                className="!w-[42px]"
                disabled={!homeState.canFormat || disabled}
                inputMode="decimal"
                value={isMixed('fontSizePx') ? '' : String(Math.round(pixelsToPoints(cellStyle.fontSizePx ?? pointsToPixels(11))))}
                onChange={(event) => {
                  const value = Number(event.target.value);
                  if (Number.isFinite(value) && value >= 1 && value <= 409) emitStyle({ fontSizePx: pointsToPixels(value) });
                }}
              />
              <Button
                aria-label="Increase font size"
                disabled={!homeState.canFormat || disabled}
                onClick={() => emitStyle({ fontSizePx: Math.min(pointsToPixels(409), (cellStyle.fontSizePx ?? pointsToPixels(11)) + pointsToPixels(1)) })}
                size="sm"
                variant="ghost"
                className="!h-8 !min-h-0 !w-7 rounded-none px-0 text-base font-semibold text-[#4e82ab]"
              >A</Button>
              <Button
                aria-label="Decrease font size"
                disabled={!homeState.canFormat || disabled}
                onClick={() => emitStyle({ fontSizePx: Math.max(pointsToPixels(1), (cellStyle.fontSizePx ?? pointsToPixels(11)) - pointsToPixels(1)) })}
                size="sm"
                variant="ghost"
                className="!h-8 !min-h-0 !w-7 rounded-none px-0 text-xs font-semibold text-[#4e82ab]"
              >A</Button>
              </Inline>
              <Inline gap="xs" className="h-8 flex-nowrap items-start">
              <CatalogButton id="bold" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('bold')} />
              <CatalogButton id="italic" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('italic')} />
              <CatalogButton id="underline" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('underline')} />
              <CatalogButton id="strikethrough" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('strikethrough')} />
              <DropdownMenu
                disabled={!canFormat({ textColor: '#1e293b' })}
                trigger={(
                  <Button
                    variant="ghost"
                    size="sm"
                    icon="type"
                    iconOnly
                    title={`${homeText(locale, 'textColor')}${isMixed('textColor') ? ` (${homeText(locale, 'mixed')})` : ''}`}
                    aria-label={homeText(locale, 'textColor')}
                    disabled={!canFormat({ textColor: '#1e293b' })}
                    className={isMixed('textColor') ? 'border border-dashed border-slate-400' : undefined}
                  />
                )}
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.textColor || '#1e293b'}
                    onChange={(color) => {
                      emitStyle({ textColor: color });
                      close();
                    }}
                  />
                )}
              </DropdownMenu>
              <DropdownMenu
                disabled={!canFormat({ background: '#ffffff' })}
                trigger={(
                  <Button
                    variant="ghost"
                    size="sm"
                    icon="paint-bucket"
                    iconOnly
                    title={`${homeText(locale, 'fillBackground')}${isMixed('background') ? ` (${homeText(locale, 'mixed')})` : ''}`}
                    aria-label={homeText(locale, 'fillBackground')}
                    disabled={!canFormat({ background: '#ffffff' })}
                    className={isMixed('background') ? 'border border-dashed border-slate-400' : undefined}
                  />
                )}
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.background || '#ffffff'}
                    onChange={(color) => {
                      emitStyle({ background: color });
                      close();
                    }}
                  />
                )}
              </DropdownMenu>
              <CatalogButton id="allBorders" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              </Inline>
              </Stack>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />

            <RibbonGroup className="!w-[260px] overflow-hidden" group="alignment">
              <CatalogButton id="alignLeft" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('horizontalAlignment')} />
              <CatalogButton id="alignCenter" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('horizontalAlignment')} />
              <CatalogButton id="alignRight" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('horizontalAlignment')} />
              <DropdownMenu
                disabled={!canFormat({ verticalAlignment: 'middle' })}
                trigger={(
                  <Button
                    variant="ghost"
                    size="sm"
                    icon="layout"
                    iconOnly
                    title={`${homeText(locale, 'vertical')}${isMixed('verticalAlignment') ? ` (${homeText(locale, 'mixed')})` : ''}`}
                    aria-label={homeText(locale, 'vertical')}
                    disabled={!canFormat({ verticalAlignment: 'middle' })}
                    className={isMixed('verticalAlignment') ? 'border border-dashed border-slate-400' : undefined}
                  />
                )}
              >
                {({ close }) => (
                  <Stack gap="xs" className="min-w-32 p-1">
                    {(['top', 'middle', 'bottom'] as const).map((verticalAlignment) => (
                      <Button
                        key={verticalAlignment}
                        size="sm"
                        variant={cellStyle.verticalAlignment === verticalAlignment ? 'secondary' : 'ghost'}
                        className="justify-start"
                        onClick={() => {
                          emitStyle({ verticalAlignment });
                          close();
                        }}
                      >
                        {homeText(locale, verticalAlignment)}
                      </Button>
                    ))}
                  </Stack>
                )}
              </DropdownMenu>
              <CatalogButton id="wrapText" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('wrapText')} />
              <Button
                size="sm"
                variant={homeState.merge === 'full' ? 'secondary' : 'ghost'}
                icon="merge-cells"
                iconOnly
                disabled={disabled || !homeState.canFormat}
                aria-label={translateRibbonText(locale, RIBBON_TEXT.commands.mergeCells)}
                title={translateRibbonText(locale, RIBBON_TEXT.commands.mergeCells)}
                data-mixed={homeState.merge === 'mixed' || undefined}
                className={homeState.merge === 'mixed' ? 'border border-dashed border-slate-400' : undefined}
                onClick={onMergeCells}
              />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />

            <RibbonGroup className="!w-[123px] overflow-hidden" group="number">
              <Box className="w-28 pt-3">
                <Select
                  aria-label={translateRibbonText(locale, RIBBON_TEXT.groups.number)}
                  title={isMixed('numberFormat') ? homeText(locale, 'mixed') : undefined}
                  sizeVariant="sm"
                  disabled={!canFormat({ numberFormat: 'general' })}
                  value={isMixed('numberFormat') ? '__mixed__' : cellStyle.numberFormat || 'general'}
                  onChange={(event) => {
                    if (event.target.value !== '__mixed__') emitStyle({ numberFormat: event.target.value });
                  }}
                >
                  {isMixed('numberFormat') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
                  <option value="general">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatGeneral)}</option>
                  <option value="$#,##0">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatCurrency)}</option>
                  <option value="0%">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatPercent)}</option>
                  <option value="#,##0">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatComma)}</option>
                  <option value="0.00">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatDecimal)}</option>
                </Select>
              </Box>
              <CatalogButton id="numberFormatCurrency" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('numberFormat')} />
              <CatalogButton id="numberFormatPercent" context={catalogContext} onExecute={executeCatalogResult} iconOnly mixed={isMixed('numberFormat')} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />

            <RibbonGroup className="!w-[353px] overflow-hidden" label={homeText(locale, 'styles')}>
              <CatalogButton id="conditionalFormat" context={catalogContext} onExecute={executeCatalogResult} textBelow />
              <CatalogButton id="formatAsTable" context={catalogContext} onExecute={executeCatalogResult} textBelow />
              <CatalogButton id="formatCells" context={catalogContext} onExecute={executeCatalogResult} textBelow testId="ribbon-format-cells" />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />

            <RibbonGroup className="!w-[54px] overflow-hidden" group="cells">
              <CatalogButton id="insertRowHome" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="insertColumnHome" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="shiftCells" context={catalogContext} onExecute={executeCatalogResult} />
              <DropdownMenu
                disabled={disabled}
                trigger={(
                  <Button size="sm" variant="ghost" disabled={disabled} icon="trash" title={homeText(locale, 'clear')}>
                    {homeText(locale, 'clear')}
                  </Button>
                )}
              >
                {({ close }) => (
                  <Stack gap="xs" className="p-1">
                    <CatalogButton id="clearContents" context={catalogContext} onExecute={(result) => { close(); executeCatalogResult(result); }} className="w-full justify-start" />
                    <CatalogButton id="clearFormats" context={catalogContext} onExecute={(result) => { close(); executeCatalogResult(result); }} className="w-full justify-start" />
                    <CatalogButton id="clearAll" context={catalogContext} onExecute={(result) => { close(); executeCatalogResult(result); }} className="w-full justify-start" />
                  </Stack>
                )}
              </DropdownMenu>
              <DropdownMenu
                disabled={disabled}
                trigger={<Button size="sm" variant="ghost" disabled={disabled} icon="columns" title="Format">Format</Button>}
              >
                {({ close }) => (
                  <Stack gap="xs" className="min-w-48 p-1">
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onOpenColumnWidth(); }}>Column Width…</Button>
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onAutoFitColumns(); }}>AutoFit Column Width</Button>
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onOpenDefaultColumnWidth(); }}>Default Column Width…</Button>
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onHideColumns(); }}>Hide Columns</Button>
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onUnhideColumns(); }}>Unhide Columns</Button>
                  </Stack>
                )}
              </DropdownMenu>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-[96px]" />

            <RibbonGroup className="!w-[111px] overflow-hidden" group="editing">
              <CatalogButton id="autoSum" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="sortRange" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="filterSelection" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="clearFilter" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="findReplace" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="goTo" context={catalogContext} onExecute={executeCatalogResult} />
              <Button
                size="sm"
                variant="ghost"
                icon="shape-square"
                disabled={disabled}
                data-testid="home-selection-pane"
                title={homeText(locale, 'selectionPane')}
                onClick={() => onSessionIntent({ type: 'panel.open', panel: 'selectionPane' })}
              >
                {homeText(locale, 'selectionPane')}
              </Button>
            </RibbonGroup>
          </Inline>
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
