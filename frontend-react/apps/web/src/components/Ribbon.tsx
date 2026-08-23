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
import { SAMPLE_AUTOMATION_SCRIPT, type AppPhase, type SidebarPanelId, type UiSessionIntent } from '@react-sheets/spreadsheet-app';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { localizeText, translate, translateRibbonTab, type Locale } from '../i18n';

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
  onAutoSum: () => void;
  onFreezeAtPrimary: () => void;
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

function RibbonGroup({ children, label }: { children: React.ReactNode; label: string }) {
  const locale = useContext(RibbonLocaleContext);
  return (
    <Stack gap="xs" className="shrink-0">
      <Inline gap="xs" className="min-h-8 items-center">
        {children}
      </Inline>
      <Text size="xs" tone="subtle" className="text-center text-[10px] font-semibold uppercase tracking-wider text-slate-400">
        {localizeText(locale, label)}
      </Text>
    </Stack>
  );
}

function ToolBtn({
  commandId,
  params,
  disabled,
  icon,
  label,
  active,
  onCommand,
  className,
}: {
  commandId: string;
  params?: unknown;
  disabled: boolean;
  icon: React.ComponentProps<typeof Icon>['name'];
  label: string;
  active?: boolean;
  onCommand: (descriptor: CommandDescriptor) => void;
  className?: string;
}) {
  const locale = useContext(RibbonLocaleContext);
  const localizedLabel = localizeText(locale, label);
  return (
    <Button
      aria-label={localizedLabel}
      title={localizedLabel}
      disabled={disabled}
      icon={icon}
      iconOnly
      onClick={() => onCommand({ commandId, params })}
      size="sm"
      variant={active ? 'primary' : 'ghost'}
      className={className}
    />
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
  onTabChange,
  phase,
  cellStyle = {},
  canExecute,
}: RibbonProps) {
  const disabled = phase !== 'ready';
  const blocked = (commandId: string, params?: unknown) => disabled || (canExecute ? !canExecute(commandId, params) : false);
  const openPanel = (panel: SidebarPanelId, notice?: string) => onSessionIntent({ type: 'panel.open', panel, notice });
  const openDialog = (dialog: Extract<UiSessionIntent, { type: 'dialog.open' }>['dialog'], findQuery?: string) => onSessionIntent({ type: 'dialog.open', dialog, findQuery });
  const dispatchBuiltCommand = (build: () => CommandDescriptor | undefined) => {
    const descriptor = build();
    if (descriptor) onCommand(descriptor);
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
        {activeTab === 'file' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label={locale === 'zh-CN' ? '工作簿' : 'Workbook'}>
              <Button size="sm" variant="secondary" disabled={disabled} onClick={onSave}>
                {locale === 'zh-CN' ? '保存' : 'Save'}
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={onExportXlsx}>
                {locale === 'zh-CN' ? '导出 XLSX' : 'Export XLSX'}
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={onImportXlsx}>
                {locale === 'zh-CN' ? '导入 XLSX' : 'Import XLSX'}
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'automate' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label={locale === 'zh-CN' ? '脚本' : 'Scripts'}>
              <Button size="sm" variant="primary" disabled={disabled} onClick={() => openPanel('automate')}>
                {locale === 'zh-CN' ? '打开自动化面板' : 'Open Automate Panel'}
              </Button>
              <Button size="sm" variant="outline" disabled={disabled} onClick={() => onCommand({ commandId: 'automation.run', params: { source: SAMPLE_AUTOMATION_SCRIPT } })}>
                {locale === 'zh-CN' ? '运行示例脚本' : 'Run Sample Script'}
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => onCommand({ commandId: 'automation.record.start', params: {} })}>
                {locale === 'zh-CN' ? '开始录制' : 'Start Recording'}
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => onCommand({ commandId: 'automation.record.stop', params: {} })}>
                {locale === 'zh-CN' ? '停止录制' : 'Stop Recording'}
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'pageLayout' ? (
          <RibbonEmptyState message={locale === 'zh-CN' ? '此选项卡的命令将在此显示。' : 'Commands for this tab will appear here.'} />
        ) : null}

        {activeTab === 'formulas' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Calculation">
              <Button size="sm" variant="ghost" icon="calculator" disabled={disabled} onClick={onRecalculate}>
                Calculate Now
              </Button>
              <Button size="sm" variant="outline" disabled={disabled} onClick={() => openPanel('extended')}>
                Goal Seek
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'home' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="History">
              <Button size="sm" variant="ghost" disabled={disabled} icon="undo" aria-label="Undo (Ctrl+Z)" onClick={onUndo} />
              <Button size="sm" variant="ghost" disabled={disabled} icon="redo" aria-label="Redo (Ctrl+Y)" onClick={onRedo} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Clipboard">
              <Button size="sm" variant="ghost" disabled={disabled} icon="scissors" aria-label="Cut (Ctrl+X)" onClick={onCut} />
              <Button size="sm" variant="ghost" disabled={disabled} icon="copy" aria-label="Copy (Ctrl+C)" onClick={onCopy} />
              <Button size="sm" variant="ghost" disabled={disabled} icon="clipboard" aria-label="Paste (Ctrl+V)" onClick={onPaste} />
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => openDialog('paste-special')}>
                Paste Special
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Font">
              <ToolBtn commandId="sheet.style.set" params={{ style: { bold: !cellStyle.bold } }} disabled={disabled} icon="bold" label="Bold (Ctrl+B)" active={cellStyle.bold} onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { italic: !cellStyle.italic } }} disabled={disabled} icon="italic" label="Italic (Ctrl+I)" active={cellStyle.italic} onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { underline: !cellStyle.underline } }} disabled={disabled} icon="underline" label="Underline (Ctrl+U)" active={cellStyle.underline} onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { strikethrough: !cellStyle.strikethrough } }} disabled={disabled} icon="strikethrough" label="Strikethrough" active={cellStyle.strikethrough} onCommand={onCommand} />

              <DropdownMenu
                trigger={
                  <Button variant="ghost" size="sm" icon="type" iconOnly title="Text Color" disabled={disabled} className="relative" />
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
                  <Button variant="ghost" size="sm" icon="paint-bucket" iconOnly title="Fill Background" disabled={disabled} />
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

              <ToolBtn commandId="sheet.style.set" params={{ style: { borders: { top: { style: 'thin', color: '#334155' }, right: { style: 'thin', color: '#334155' }, bottom: { style: 'thin', color: '#334155' }, left: { style: 'thin', color: '#334155' } } } }} disabled={disabled} icon="borders" label="All Borders" onCommand={onCommand} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Alignment">
              <ToolBtn commandId="sheet.style.set" params={{ style: { horizontalAlignment: 'left' } }} disabled={disabled} icon="align-left" label="Align Left" active={cellStyle.align === 'left'} onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { horizontalAlignment: 'center' } }} disabled={disabled} icon="align-center" label="Align Center" active={cellStyle.align === 'center'} onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { horizontalAlignment: 'right' } }} disabled={disabled} icon="align-right" label="Align Right" active={cellStyle.align === 'right'} onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { wrapText: !cellStyle.wrapText } }} disabled={disabled} icon="wrap-text" label="Wrap Text" active={cellStyle.wrapText} onCommand={onCommand} />
              <ToolBtn commandId="sheet.merge.set" disabled={disabled} icon="merge-cells" label="Merge & Center" onCommand={onCommand} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Number">
              <Button size="sm" variant="ghost" disabled={disabled} data-testid="ribbon-format-cells" onClick={() => openDialog('format-cells')}>
                Format Cells
              </Button>
              <div className="w-28">
                <Select
                  sizeVariant="sm"
                  disabled={disabled}
                  value={cellStyle.numberFormat || 'general'}
                  onChange={(e) => onCommand({ commandId: 'sheet.style.set', params: { style: { numberFormat: e.target.value } } })}
                >
                  <option value="general">General</option>
                  <option value="$#,##0">Currency ($)</option>
                  <option value="0%">Percent (%)</option>
                  <option value="#,##0">Comma (1,000)</option>
                  <option value="0.00">Decimal (0.00)</option>
                </Select>
              </div>
              <ToolBtn commandId="sheet.style.set" params={{ style: { numberFormat: '$#,##0' } }} disabled={disabled} icon="dollar-sign" label="Currency Format" onCommand={onCommand} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { numberFormat: '0%' } }} disabled={disabled} icon="percent" label="Percent Format" onCommand={onCommand} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Cells">
              <ToolBtn commandId="sheet.rows.insert" params={{ count: 1 }} disabled={disabled} icon="rows" label="Insert Row" onCommand={onCommand} />
              <ToolBtn commandId="sheet.columns.insert" params={{ count: 1 }} disabled={disabled} icon="columns" label="Insert Column" onCommand={onCommand} />
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => openDialog('shift-cells')}>
                Insert / Delete Cells
              </Button>
              <DropdownMenu
                trigger={
                  <Button size="sm" variant="ghost" disabled={disabled} icon="trash">
                    Clear
                  </Button>
                }
              >
                {({ close }) => (
                  <Stack gap="xs" className="p-1">
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onCommand({ commandId: 'sheet.range.clear' }); }}>
                      Clear Contents
                    </Button>
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onCommand({ commandId: 'sheet.range.clear', params: { mode: 'formats' } }); }}>
                      Clear Formats
                    </Button>
                    <Button size="sm" variant="ghost" className="justify-start" onClick={() => { close(); onCommand({ commandId: 'sheet.range.clear', params: { mode: 'all' } }); }}>
                      Clear All
                    </Button>
                  </Stack>
                )}
              </DropdownMenu>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Editing">
              <Button size="sm" variant="ghost" icon="calculator" disabled={disabled} onClick={onAutoSum}>AutoSum =SUM()</Button>
              <Button size="sm" variant="ghost" icon="sort" disabled={disabled} onClick={() => openDialog('sort-dialog')}>Sort Range</Button>
              <Button size="sm" variant="ghost" icon="sparkles" disabled={disabled} onClick={() => openPanel('conditionalFormat')}>Conditional Format</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'insert' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Tables & Pivots">
              <Button size="sm" variant="ghost" icon="table-pivot" onClick={() => openPanel('pivot')}>
                Pivot Table
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onCreatePivot)}>
                Quick Pivot
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Charts & Visuals">
              <Button size="sm" variant="ghost" icon="chart" onClick={() => openPanel('chart')}>
                Chart Builder
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onCreateChart)}>
                Column Chart
              </Button>
              <Button size="sm" variant="ghost" icon="sparkline" onClick={() => openPanel('sparkline')}>
                Sparkline
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onCreateSparkline)}>
                Quick Sparkline
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Illustrations">
              <Button size="sm" variant="ghost" icon="shape-square" onClick={() => openPanel('shape')}>
                Shapes & Lines
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onCreateShape)}>
                Rectangle
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onBringDrawingForward)}>
                Bring Forward
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onSendDrawingBackward)}>
                Send Backward
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onRemoveDrawing)}>
                Remove Drawing
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Functions">
              <Button size="sm" variant="ghost" icon="function" onClick={() => openDialog('function-wizard')}>
                Insert Function (fx)
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Cells">
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'sheet.rows.insert', params: { count: 1 } })}>Insert Row</Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'sheet.columns.insert', params: { count: 1 } })}>Insert Column</Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'sheet.rows.delete' })}>Delete Row</Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'sheet.columns.delete' })}>Delete Column</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'data' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Sort & Filter">
              <Button size="sm" variant="ghost" icon="sort" onClick={() => onCommand({ commandId: 'data.sort.rows', params: { criteria: [{ column: 0, ascending: true }] } })}>
                Sort A to Z
              </Button>
              <Button size="sm" variant="ghost" icon="sort" onClick={() => onCommand({ commandId: 'data.sort.rows', params: { criteria: [{ column: 0, ascending: false }] } })}>
                Sort Z to A
              </Button>
              <Button size="sm" variant="ghost" icon="sliders" onClick={() => openDialog('sort-dialog')}>
                Custom Sort...
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Data Tools">
              <Button size="sm" variant="ghost" icon="table" onClick={() => openPanel('data')}>
                Data Model
              </Button>
              <Button size="sm" variant="ghost" icon="table" onClick={onCreateDataTable}>
                Create Data Table
              </Button>
              <Button size="sm" variant="ghost" icon="table" disabled={disabled} onClick={onCreateSheetTable}>
                Format as Table
              </Button>
              <Button size="sm" variant="ghost" icon="table" disabled={disabled} onClick={() => dispatchBuiltCommand(onToggleSheetTableTotalRow)}>
                Total Row
              </Button>
              <Button size="sm" variant="ghost" icon="check-circle" onClick={() => openPanel('dataValidation')}>
                Data Validation
              </Button>
              <Button size="sm" variant="ghost" onClick={() => dispatchBuiltCommand(onApplyFilterSelection)}>
                Filter Selection
              </Button>
              <Button size="sm" variant="ghost" onClick={() => dispatchBuiltCommand(onClearFilter)}>
                Clear Filter
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Outline">
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onGroupRows)}>
                Group Rows
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onUngroupRows)}>
                Ungroup Rows
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onGroupColumns)}>
                Group Columns
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onUngroupColumns)}>
                Ungroup Columns
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => onCommand({ commandId: 'outline.showLevel', params: { level: 1 } })}>
                Show Level 1
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => onCommand({ commandId: 'outline.showLevel', params: { level: 2 } })}>
                Show Level 2
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => onCommand({ commandId: 'outline.showLevel', params: { level: 3 } })}>
                Show Level 3
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onSubtotal)}>
                Subtotal
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onRemoveDuplicates)}>
                Remove Duplicates
              </Button>
              <Button size="sm" variant="ghost" disabled={disabled} onClick={() => dispatchBuiltCommand(onTextToColumns)}>
                Text to Columns
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Find & Transform">
              <Button size="sm" variant="ghost" onClick={() => openDialog('find-replace')}>
                Find & Replace
              </Button>
              <Button size="sm" variant="ghost" onClick={() => openDialog('goto')}>
                Go To
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'matrix.transpose' })}>
                Transpose
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'matrix.flip', params: { direction: 'horizontal' } })}>Flip H</Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'matrix.flip', params: { direction: 'vertical' } })}>Flip V</Button>
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'data.textToColumns', params: { delimiter: ',', maxColumns: 8 } })}>Split by Delimiter</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'review' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Comments">
              <Button size="sm" variant="ghost" icon="comment" disabled={disabled} onClick={() => openPanel('inspector', 'Add a comment in the Inspector panel.')}>
                New Comment
              </Button>
              <Button size="sm" variant="ghost" icon="comment" disabled={blocked('comment.resolve')} onClick={() => onCommand({ commandId: 'comment.resolve' })}>
                Resolve
              </Button>
              <Button size="sm" variant="ghost" icon="comment" onClick={() => openPanel('inspector')}>
                Show Comments
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup label="Notes & Links">
              <Button size="sm" variant="ghost" icon="comment" onClick={() => openPanel('inspector', 'Add a cell note in the Inspector panel.')}>
                New Note
              </Button>
              <Button size="sm" variant="ghost" icon="share" onClick={() => openPanel('inspector', 'Insert a hyperlink in the Inspector panel.')}>
                Insert Link
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup label="Protection">
              <Button size="sm" variant="ghost" icon="lock" disabled={blocked('permission.protect.selection')} onClick={() => onCommand({ commandId: 'permission.protect.selection' })}>
                Protect Selection
              </Button>
              <Button size="sm" variant="ghost" icon="lock" disabled={blocked('permission.unprotect.selection')} onClick={() => onCommand({ commandId: 'permission.unprotect.selection' })}>
                Unprotect
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />
            <RibbonGroup label="History & Audit">
              <Button size="sm" variant="ghost" icon="history" onClick={() => openPanel('history')}>
                Revision Log
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'view' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Freeze Panes">
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onCommand({ commandId: 'sheet.freeze.set', params: { freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 } } })}>
                Freeze Top Row
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onCommand({ commandId: 'sheet.freeze.set', params: { freeze: { xSplit: 1, ySplit: 0, startRow: 0, startColumn: 1 } } })}>
                Freeze First Column
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={onFreezeAtPrimary}>
                Freeze at Selection
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onCommand({ commandId: 'sheet.freeze.set', params: { freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 } } })}>
                Unfreeze All
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Zoom">
              <Button size="sm" variant="ghost" icon="zoom-in" onClick={() => onSessionIntent({ type: 'zoom.adjust', delta: 10 })}>
                Zoom In
              </Button>
              <Button size="sm" variant="ghost" icon="zoom-out" onClick={() => onSessionIntent({ type: 'zoom.adjust', delta: -10 })}>
                Zoom Out
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onSessionIntent({ type: 'zoom.set', value: 100 })}>
                100%
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Print Layout">
              <Button size="sm" variant="ghost" icon="printer" onClick={() => openDialog('print-preview')}>
                Print & PDF
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Appearance & Files">
              <Button size="sm" variant="ghost" onClick={() => onCommand({ commandId: 'sheet.banded.set' })}>
                Banded Rows
              </Button>
              <Button size="sm" variant="ghost" onClick={onExportXlsx}>
                Export .xlsx
              </Button>
              <Button size="sm" variant="ghost" onClick={onImportXlsx}>
                Import .xlsx
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}
      </RibbonShell>
    </RibbonLocaleContext.Provider>
  );
}
