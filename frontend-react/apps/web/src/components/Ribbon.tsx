import React, { createContext, useContext } from 'react';
import {
  Box,
  Button,
  ColorPicker,
  Divider,
  DropdownMenu,
  Icon,
  Inline,
  RibbonEmptyState,
  RibbonShell,
  Select,
  Stack,
  Text,
  type RibbonTabId,
} from '@react-sheets/ui-system';
import {
  SAMPLE_AUTOMATION_SCRIPT,
  buildRibbonCommand,
  getRibbonCommandDefinition,
  getRibbonGroupDefinition,
  isRibbonCommandEnabled,
  RIBBON_TEXT,
  type AppPhase,
  type RibbonCommandActions,
  type RibbonCommandContext,
  type RibbonCommandId,
  type RibbonCommandResult,
  type RibbonGroupId,
  type UiSessionIntent,
} from '@react-sheets/spreadsheet-app';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { localizeText, translate, translateRibbonTab, translateRibbonText, type Locale } from '../i18n';

export interface RibbonProps {
  activeTab: RibbonTabId;
  locale: Locale;
  onCommand: (descriptor: CommandDescriptor) => void;
  onSessionIntent: (intent: UiSessionIntent) => void;
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
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
  onAutoSum: () => void;
  onFreezeAtPrimary: () => void;
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
  onTabChange: (tab: RibbonTabId) => void;
  phase: AppPhase;
  cellStyle?: {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    align?: 'left' | 'center' | 'right';
    verticalAlign?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
    numberFormat?: string;
    background?: string;
    textColor?: string;
  };
  canExecute?: (commandId: string, params?: unknown) => boolean;
}

const RibbonLocaleContext = createContext<Locale>('en-US');
type RibbonLayoutMode = 'wide' | 'compact' | 'narrow';
const RibbonLayoutContext = createContext<RibbonLayoutMode>('wide');

function RibbonGroup({ children, label, group }: { children: React.ReactNode; label?: string; group?: RibbonGroupId }) {
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
    <Stack gap="xs" className="min-w-0 shrink-0">
      <Inline gap="xs" className="min-h-8 items-center">
        {children}
      </Inline>
      <Text size="xs" tone="subtle" className="text-center text-[10px] font-semibold uppercase tracking-wider text-slate-400">
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
  variant = 'ghost',
  className,
  testId,
}: {
  id: RibbonCommandId;
  context: RibbonCommandContext;
  onExecute: (result: RibbonCommandResult) => void;
  iconOnly?: boolean;
  variant?: 'danger' | 'ghost' | 'outline' | 'primary' | 'secondary' | 'soft';
  className?: string;
  testId?: string;
}) {
  const locale = useContext(RibbonLocaleContext);
  const layout = useContext(RibbonLayoutContext);
  const definition = getRibbonCommandDefinition(id);
  const enabled = isRibbonCommandEnabled(definition, context);
  const label = translateRibbonText(locale, definition.labelKey);
  const compactIcon = layout !== 'wide' && definition.display === 'small';
  return (
    <Button
      aria-label={label}
      title={label}
      data-testid={testId}
      disabled={!enabled}
      icon={definition.icon}
      iconOnly={iconOnly || compactIcon}
      onClick={() => {
        const result = buildRibbonCommand(id, context);
        if (result) onExecute(result);
      }}
      size="sm"
      variant={definition.active?.(context) ? 'primary' : variant}
      className={className}
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
  onAutoSum,
  onFreezeAtPrimary,
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
  onTabChange,
  phase,
  cellStyle = {},
  canExecute,
}: RibbonProps) {
  const disabled = phase !== 'ready';
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
  };
  const catalogContext: RibbonCommandContext = {
    phase,
    disabled,
    cellStyle,
    canExecute,
    buildSortDescriptor,
    openCreatePivotDialog: onCreatePivotDialog,
    actions: catalogActions,
    dispatchSessionIntent: onSessionIntent,
    sampleAutomationScript: SAMPLE_AUTOMATION_SCRIPT,
  };
  const executeCatalogResult = (result: RibbonCommandResult) => {
    if (result.type === 'command') onCommand(result.descriptor);
    else if (result.type === 'intent') onSessionIntent(result.intent);
    else result.invoke();
  };

  return (
    <RibbonLocaleContext.Provider value={locale}>
      <RibbonShell
        activeTab={activeTab}
        disabled={disabled}
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
          <RibbonEmptyState message={locale === 'zh-CN' ? '此选项卡的命令将在此显示。' : 'Commands for this tab will appear here.'} />
        ) : null}

        {activeTab === 'formulas' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="calculation">
              <CatalogButton id="calculateNow" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="goalSeek" context={catalogContext} onExecute={executeCatalogResult} variant="outline" />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup group="formulaAudit">
              <CatalogButton id="tracePrecedents" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="traceDependents" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="removeArrows" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="showFormulas" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="errorChecking" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="evaluateFormula" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'home' ? (
          <Inline gap="md" className="flex-wrap items-start">
            <RibbonGroup group="history">
              <CatalogButton id="undo" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="redo" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="clipboard">
              <CatalogButton id="cut" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="copy" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="paste" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="pasteSpecial" context={catalogContext} onExecute={executeCatalogResult} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="font">
              <CatalogButton id="bold" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="italic" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="underline" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="strikethrough" context={catalogContext} onExecute={executeCatalogResult} iconOnly />

              <DropdownMenu
                trigger={
                  <Button variant="ghost" size="sm" icon="type" iconOnly title={localizeText(locale, 'Text Color')} disabled={disabled} className="relative" />
                }
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.textColor || '#1e293b'}
                    onChange={(c) => {
                      onCommand({ commandId: 'sheet.style.set', params: { style: { textColor: c } } });
                      close();
                    }}
                  />
                )}
              </DropdownMenu>

              <DropdownMenu
                trigger={
                  <Button variant="ghost" size="sm" icon="paint-bucket" iconOnly title={localizeText(locale, 'Fill Background')} disabled={disabled} />
                }
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.background || '#ffffff'}
                    onChange={(c) => {
                      onCommand({ commandId: 'sheet.style.set', params: { style: { background: c } } });
                      close();
                    }}
                  />
                )}
              </DropdownMenu>

              <CatalogButton id="allBorders" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="alignment">
              <CatalogButton id="alignLeft" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="alignCenter" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="alignRight" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="wrapText" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="mergeCells" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="number">
              <CatalogButton id="formatCells" context={catalogContext} onExecute={executeCatalogResult} testId="ribbon-format-cells" />
              <Box className="w-28">
                <Select
                  sizeVariant="sm"
                  disabled={disabled}
                  value={cellStyle.numberFormat || 'general'}
                  onChange={(e) => onCommand({ commandId: 'sheet.style.set', params: { style: { numberFormat: e.target.value } } })}
                >
                  <option value="general">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatGeneral)}</option>
                  <option value="$#,##0">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatCurrency)}</option>
                  <option value="0%">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatPercent)}</option>
                  <option value="#,##0">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatComma)}</option>
                  <option value="0.00">{translateRibbonText(locale, RIBBON_TEXT.commands.numberFormatDecimal)}</option>
                </Select>
              </Box>
              <CatalogButton id="numberFormatCurrency" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="numberFormatPercent" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="cells">
              <CatalogButton id="insertRow" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="insertColumn" context={catalogContext} onExecute={executeCatalogResult} iconOnly />
              <CatalogButton id="shiftCells" context={catalogContext} onExecute={executeCatalogResult} />
              <DropdownMenu
                trigger={
                  <Button size="sm" variant="ghost" disabled={disabled} icon="trash">
                    {localizeText(locale, 'Clear')}
                  </Button>
                }
              >
                {({ close }) => (
                  <Stack gap="xs" className="p-1">
                    <CatalogButton id="clearContents" context={catalogContext} onExecute={(result) => { close(); executeCatalogResult(result); }} className="w-full justify-start" />
                    <CatalogButton id="clearFormats" context={catalogContext} onExecute={(result) => { close(); executeCatalogResult(result); }} className="w-full justify-start" />
                    <CatalogButton id="clearAll" context={catalogContext} onExecute={(result) => { close(); executeCatalogResult(result); }} className="w-full justify-start" />
                  </Stack>
                )}
              </DropdownMenu>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup group="editing">
              <CatalogButton id="autoSum" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="sortRange" context={catalogContext} onExecute={executeCatalogResult} />
              <CatalogButton id="conditionalFormat" context={catalogContext} onExecute={executeCatalogResult} />
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
    </RibbonLocaleContext.Provider>
  );
}
