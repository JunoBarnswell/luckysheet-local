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
import type { AppPhase } from '@react-sheets/spreadsheet-app';
import { localizeText, translate, translateRibbonTab, type Locale } from '../i18n';

export interface RibbonProps {
  activeTab: RibbonTabId;
  locale: Locale;
  onExecute: (commandId: string, params?: unknown) => void;
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
  onExecute,
  className,
}: {
  commandId: string;
  params?: unknown;
  disabled: boolean;
  icon: React.ComponentProps<typeof Icon>['name'];
  label: string;
  active?: boolean;
  onExecute: (commandId: string, params?: unknown) => void;
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
      onClick={() => onExecute(commandId, params)}
      size="sm"
      variant={active ? 'primary' : 'ghost'}
      className={className}
    />
  );
}

export function Ribbon({ activeTab, locale, onExecute, onTabChange, phase, cellStyle = {} }: RibbonProps) {
  const disabled = phase !== 'ready';

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
        {activeTab === 'file' || activeTab === 'pageLayout' || activeTab === 'formulas' || activeTab === 'automate' ? (
          <RibbonEmptyState message={locale === 'zh-CN' ? '此选项卡的命令将在此显示。' : 'Commands for this tab will appear here.'} />
        ) : null}

        {activeTab === 'home' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="History">
              <ToolBtn commandId="ui.history.undo" disabled={disabled} icon="undo" label="Undo (Ctrl+Z)" onExecute={onExecute} />
              <ToolBtn commandId="ui.history.redo" disabled={disabled} icon="redo" label="Redo (Ctrl+Y)" onExecute={onExecute} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Clipboard">
              <ToolBtn commandId="ui.clipboard.cut" disabled={disabled} icon="scissors" label="Cut (Ctrl+X)" onExecute={onExecute} />
              <ToolBtn commandId="ui.clipboard.copy" disabled={disabled} icon="copy" label="Copy (Ctrl+C)" onExecute={onExecute} />
              <ToolBtn commandId="ui.clipboard.paste" disabled={disabled} icon="clipboard" label="Paste (Ctrl+V)" onExecute={onExecute} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Font">
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'bold' }} disabled={disabled} icon="bold" label="Bold (Ctrl+B)" active={cellStyle.bold} onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'italic' }} disabled={disabled} icon="italic" label="Italic (Ctrl+I)" active={cellStyle.italic} onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'underline' }} disabled={disabled} icon="underline" label="Underline (Ctrl+U)" active={cellStyle.underline} onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'strikethrough' }} disabled={disabled} icon="strikethrough" label="Strikethrough" active={cellStyle.strikethrough} onExecute={onExecute} />

              <DropdownMenu
                trigger={
                  <Button variant="ghost" size="sm" icon="type" iconOnly title="Text Color" disabled={disabled} className="relative" />
                }
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.textColor || '#1e293b'}
                    onChange={(c) => {
                      onExecute('sheet.style.set', { style: { textColor: c } });
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
                      onExecute('sheet.style.set', { style: { background: c } });
                      close();
                    }}
                  />
                )}
              </DropdownMenu>

              <ToolBtn commandId="sheet.style.set" params={{ style: { borders: { top: { style: 'thin', color: '#334155' }, right: { style: 'thin', color: '#334155' }, bottom: { style: 'thin', color: '#334155' }, left: { style: 'thin', color: '#334155' } } } }} disabled={disabled} icon="borders" label="All Borders" onExecute={onExecute} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Alignment">
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'horizontalAlignment', value: 'left' }} disabled={disabled} icon="align-left" label="Align Left" active={cellStyle.align === 'left'} onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'horizontalAlignment', value: 'center' }} disabled={disabled} icon="align-center" label="Align Center" active={cellStyle.align === 'center'} onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'horizontalAlignment', value: 'right' }} disabled={disabled} icon="align-right" label="Align Right" active={cellStyle.align === 'right'} onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.toggle" params={{ property: 'wrapText' }} disabled={disabled} icon="wrap-text" label="Wrap Text" active={cellStyle.wrapText} onExecute={onExecute} />
              <ToolBtn commandId="sheet.merge.set" disabled={disabled} icon="merge-cells" label="Merge & Center" onExecute={onExecute} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Number">
              <div className="w-28">
                <Select
                  sizeVariant="sm"
                  disabled={disabled}
                  value={cellStyle.numberFormat || 'general'}
                  onChange={(e) => onExecute('sheet.style.set', { style: { numberFormat: e.target.value } })}
                >
                  <option value="general">General</option>
                  <option value="$#,##0">Currency ($)</option>
                  <option value="0%">Percent (%)</option>
                  <option value="#,##0">Comma (1,000)</option>
                  <option value="0.00">Decimal (0.00)</option>
                </Select>
              </div>
              <ToolBtn commandId="sheet.style.set" params={{ style: { numberFormat: '$#,##0' } }} disabled={disabled} icon="dollar-sign" label="Currency Format" onExecute={onExecute} />
              <ToolBtn commandId="sheet.style.set" params={{ style: { numberFormat: '0%' } }} disabled={disabled} icon="percent" label="Percent Format" onExecute={onExecute} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Cells">
              <ToolBtn commandId="sheet.rows.insert" params={{ count: 1 }} disabled={disabled} icon="rows" label="Insert Row" onExecute={onExecute} />
              <ToolBtn commandId="sheet.columns.insert" params={{ count: 1 }} disabled={disabled} icon="columns" label="Insert Column" onExecute={onExecute} />
              <ToolBtn commandId="sheet.range.clear" disabled={disabled} icon="trash" label="Clear Content" onExecute={onExecute} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Editing">
              <ToolBtn commandId="sheet.formula.autosum" disabled={disabled} icon="calculator" label="AutoSum =SUM()" onExecute={onExecute} />
              <ToolBtn commandId="ui.dialog.open" params={{ dialog: 'sort-dialog' }} disabled={disabled} icon="sort" label="Sort Range" onExecute={onExecute} />
              <ToolBtn commandId="ui.panel.open" params={{ panel: 'conditionalFormat' }} disabled={disabled} icon="sparkles" label="Conditional Format" onExecute={onExecute} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'insert' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Tables & Pivots">
              <Button size="sm" variant="ghost" icon="table-pivot" onClick={() => onExecute('ui.panel.open', { panel: 'pivot' })}>
                Pivot Table
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Charts & Visuals">
              <Button size="sm" variant="ghost" icon="chart" onClick={() => onExecute('ui.panel.open', { panel: 'chart' })}>
                Chart Builder
              </Button>
              <Button size="sm" variant="ghost" icon="sparkline" onClick={() => onExecute('ui.panel.open', { panel: 'sparkline' })}>
                Sparkline
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Illustrations">
              <Button size="sm" variant="ghost" icon="shape-square" onClick={() => onExecute('ui.panel.open', { panel: 'shape' })}>
                Shapes & Lines
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Functions">
              <Button size="sm" variant="ghost" icon="function" onClick={() => onExecute('ui.dialog.open', { dialog: 'function-wizard' })}>
                Insert Function (fx)
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Cells">
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.rows.insert', { count: 1 })}>Insert Row</Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.columns.insert', { count: 1 })}>Insert Column</Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.rows.delete')}>Delete Row</Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.columns.delete')}>Delete Column</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'data' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Sort & Filter">
              <Button size="sm" variant="ghost" icon="sort" onClick={() => onExecute('sheet.sort.multi', { ascending: true })}>
                Sort A to Z
              </Button>
              <Button size="sm" variant="ghost" icon="sort" onClick={() => onExecute('sheet.sort.multi', { ascending: false })}>
                Sort Z to A
              </Button>
              <Button size="sm" variant="ghost" icon="sliders" onClick={() => onExecute('ui.dialog.open', { dialog: 'sort-dialog' })}>
                Custom Sort...
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Data Tools">
              <Button size="sm" variant="ghost" icon="table" onClick={() => onExecute('ui.panel.open', { panel: 'data' })}>
                Data Model
              </Button>
              <Button size="sm" variant="ghost" icon="table" onClick={() => onExecute('table.create')}>
                Create Data Table
              </Button>
              <Button size="sm" variant="ghost" icon="check-circle" onClick={() => onExecute('ui.panel.open', { panel: 'dataValidation' })}>
                Data Validation
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.filter.set')}>
                Filter Selection
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.filter.remove')}>
                Clear Filter
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Find & Transform">
              <Button size="sm" variant="ghost" onClick={() => onExecute('ui.dialog.open', { dialog: 'find-replace' })}>
                Find & Replace
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('matrix.transpose')}>
                Transpose
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('matrix.flip', { axis: 'h' })}>Flip H</Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('matrix.flip', { axis: 'v' })}>Flip V</Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.splitColumn', ',')}>Split by Delimiter</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'review' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="History & Audit">
              <Button size="sm" variant="ghost" icon="history" onClick={() => onExecute('ui.panel.open', { panel: 'history' })}>
                Revision Log
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'view' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Freeze Panes">
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onExecute('sheet.freeze.set', { freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 } })}>
                Freeze Top Row
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onExecute('sheet.freeze.set', { freeze: { xSplit: 1, ySplit: 0, startRow: 0, startColumn: 1 } })}>
                Freeze First Column
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onExecute('sheet.freeze.set', { freeze: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 } })}>
                Unfreeze All
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Zoom">
              <Button size="sm" variant="ghost" icon="zoom-in" onClick={() => onExecute('ui.zoom.adjust', { delta: 10 })}>
                Zoom In
              </Button>
              <Button size="sm" variant="ghost" icon="zoom-out" onClick={() => onExecute('ui.zoom.adjust', { delta: -10 })}>
                Zoom Out
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('ui.zoom.set', { value: 100 })}>
                100%
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Print Layout">
              <Button size="sm" variant="ghost" icon="printer" onClick={() => onExecute('ui.dialog.open', { dialog: 'print-preview' })}>
                Print & PDF
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Appearance & Files">
              <Button size="sm" variant="ghost" onClick={() => onExecute('sheet.banded.set')}>
                Banded Rows
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('ui.file.export-xlsx')}>
                Export .xlsx
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onExecute('ui.file.import-xlsx')}>
                Import .xlsx
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}
      </RibbonShell>
    </RibbonLocaleContext.Provider>
  );
}
