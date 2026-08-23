import React, { createContext, useContext } from 'react';
import {
  Box,
  Button,
  ColorPicker,
  Divider,
  DropdownMenu,
  Icon,
  Inline,
  Select,
  Stack,
  Tab,
  TabList,
  Tabs,
  Text,
} from '@react-sheets/ui-system';
import type { RibbonTabId, WorkspacePhase } from '../state/workspace';
import { localizeText, translate, type Locale } from '../i18n';
import type { RibbonAction } from '../domain/ribbon-actions';

export interface RibbonProps {
  activeTab: RibbonTabId;
  locale: Locale;
  onAction: (action: RibbonAction, payload?: unknown) => void;
  onTabChange: (tab: RibbonTabId) => void;
  phase: WorkspacePhase;
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

const ribbonTabs: Array<{ id: RibbonTabId; label: string }> = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'view', label: 'View' },
];

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
  action,
  disabled,
  icon,
  label,
  active,
  onAction,
  className,
}: {
  action: RibbonAction;
  disabled: boolean;
  icon: React.ComponentProps<typeof Icon>['name'];
  label: string;
  active?: boolean;
  onAction: (action: RibbonAction) => void;
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
      onClick={() => onAction(action)}
      size="sm"
      variant={active ? 'primary' : 'ghost'}
      className={className}
    />
  );
}

export function Ribbon({ activeTab, locale, onAction, onTabChange, phase, cellStyle = {} }: RibbonProps) {
  const disabled = phase !== 'ready';

  return (
    <RibbonLocaleContext.Provider value={locale}>
      <Tabs className="border-b border-slate-200 bg-white">
      {/* Top Tab Bar */}
      <Inline gap="lg" className="h-10 overflow-x-auto px-4">
        <TabList label="Workbook ribbon tabs" className="h-full gap-1">
              {ribbonTabs.map((tab) => (
            <Tab
              key={tab.id}
              active={activeTab === tab.id}
              disabled={disabled}
              onClick={() => onTabChange(tab.id)}
              className="h-full border-b-2 border-transparent px-3 text-xs font-semibold data-active:border-blue-600 data-active:text-blue-600"
            >
              {translate(locale, tab.id)}
            </Tab>
          ))}
        </TabList>

        <Inline gap="xs" className="ml-auto shrink-0 border-l border-slate-100 pl-3">
          <Icon name="cloud-check" size="sm" className={disabled ? 'text-slate-300' : 'text-emerald-500'} />
          <Text size="xs" tone="muted">{translate(locale, 'engineConnected')}</Text>
        </Inline>
      </Inline>

      {/* Main Ribbon Toolbar Area */}
      <Box className="overflow-x-auto border-t border-slate-100 bg-slate-50/80 px-4 py-2">
        {activeTab === 'home' ? (
          <Inline gap="md" className="min-w-max items-start">
            {/* Undo / Redo */}
            <RibbonGroup label="History">
              <ToolBtn action="undo" disabled={disabled} icon="undo" label="Undo (Ctrl+Z)" onAction={onAction} />
              <ToolBtn action="redo" disabled={disabled} icon="redo" label="Redo (Ctrl+Y)" onAction={onAction} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            {/* Clipboard */}
            <RibbonGroup label="Clipboard">
              <ToolBtn action="cut" disabled={disabled} icon="scissors" label="Cut (Ctrl+X)" onAction={onAction} />
              <ToolBtn action="copy" disabled={disabled} icon="copy" label="Copy (Ctrl+C)" onAction={onAction} />
              <ToolBtn action="paste" disabled={disabled} icon="clipboard" label="Paste (Ctrl+V)" onAction={onAction} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            {/* Font Formatting */}
            <RibbonGroup label="Font">
              <ToolBtn
                action="bold"
                disabled={disabled}
                icon="bold"
                label="Bold (Ctrl+B)"
                active={cellStyle.bold}
                onAction={onAction}
              />
              <ToolBtn
                action="italic"
                disabled={disabled}
                icon="italic"
                label="Italic (Ctrl+I)"
                active={cellStyle.italic}
                onAction={onAction}
              />
              <ToolBtn
                action="underline"
                disabled={disabled}
                icon="underline"
                label="Underline (Ctrl+U)"
                active={cellStyle.underline}
                onAction={onAction}
              />
              <ToolBtn
                action="strikethrough"
                disabled={disabled}
                icon="strikethrough"
                label="Strikethrough"
                active={cellStyle.strikethrough}
                onAction={onAction}
              />

              {/* Text Color Picker */}
              <DropdownMenu
                trigger={
                  <Button
                    variant="ghost"
                    size="sm"
                    icon="type"
                    iconOnly
                    title="Text Color"
                    disabled={disabled}
                    className="relative"
                  />
                }
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.textColor || '#1e293b'}
                    onChange={(c) => {
                      onAction('textColor', c);
                      close();
                    }}
                  />
                )}
              </DropdownMenu>

              {/* Fill Color Picker */}
              <DropdownMenu
                trigger={
                  <Button
                    variant="ghost"
                    size="sm"
                    icon="paint-bucket"
                    iconOnly
                    title="Fill Background"
                    disabled={disabled}
                  />
                }
              >
                {({ close }) => (
                  <ColorPicker
                    color={cellStyle.background || '#ffffff'}
                    onChange={(c) => {
                      onAction('background', c);
                      close();
                    }}
                  />
                )}
              </DropdownMenu>

              {/* Borders */}
              <ToolBtn action="border-all" disabled={disabled} icon="borders" label="All Borders" onAction={onAction} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            {/* Alignment */}
            <RibbonGroup label="Alignment">
              <ToolBtn
                action="align-left"
                disabled={disabled}
                icon="align-left"
                label="Align Left"
                active={cellStyle.align === 'left'}
                onAction={onAction}
              />
              <ToolBtn
                action="align-center"
                disabled={disabled}
                icon="align-center"
                label="Align Center"
                active={cellStyle.align === 'center'}
                onAction={onAction}
              />
              <ToolBtn
                action="align-right"
                disabled={disabled}
                icon="align-right"
                label="Align Right"
                active={cellStyle.align === 'right'}
                onAction={onAction}
              />
              <ToolBtn
                action="wrap-text"
                disabled={disabled}
                icon="wrap-text"
                label="Wrap Text"
                active={cellStyle.wrapText}
                onAction={onAction}
              />
              <ToolBtn
                action="merge-cells"
                disabled={disabled}
                icon="merge-cells"
                label="Merge & Center"
                onAction={onAction}
              />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            {/* Number Formats */}
            <RibbonGroup label="Number">
              <div className="w-28">
                <Select
                  sizeVariant="sm"
                  disabled={disabled}
                  value={cellStyle.numberFormat || 'general'}
                  onChange={(e) => onAction('numberFormat', e.target.value)}
                >
                  <option value="general">General</option>
                  <option value="$#,##0">Currency ($)</option>
                  <option value="0%">Percent (%)</option>
                  <option value="#,##0">Comma (1,000)</option>
                  <option value="0.00">Decimal (0.00)</option>
                </Select>
              </div>
              <ToolBtn action="format-currency" disabled={disabled} icon="dollar-sign" label="Currency Format" onAction={onAction} />
              <ToolBtn action="format-percent" disabled={disabled} icon="percent" label="Percent Format" onAction={onAction} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            {/* Cells & Editing */}
            <RibbonGroup label="Cells">
              <ToolBtn action="insert-row" disabled={disabled} icon="rows" label="Insert Row" onAction={onAction} />
              <ToolBtn action="insert-column" disabled={disabled} icon="columns" label="Insert Column" onAction={onAction} />
              <ToolBtn action="clear-range" disabled={disabled} icon="trash" label="Clear Content" onAction={onAction} />
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            {/* Analysis & Tools */}
            <RibbonGroup label="Editing">
              <ToolBtn action="autosum" disabled={disabled} icon="calculator" label="AutoSum =SUM()" onAction={onAction} />
              <ToolBtn action="sort-dialog" disabled={disabled} icon="sort" label="Sort Range" onAction={onAction} />
              <ToolBtn action="open-conditional-format" disabled={disabled} icon="sparkles" label="Conditional Format" onAction={onAction} />
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'insert' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Tables & Pivots">
              <Button size="sm" variant="ghost" icon="table-pivot" onClick={() => onAction('open-pivot')}>
                Pivot Table
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Charts & Visuals">
              <Button size="sm" variant="ghost" icon="chart" onClick={() => onAction('open-chart')}>
                Chart Builder
              </Button>
              <Button size="sm" variant="ghost" icon="sparkline" onClick={() => onAction('open-sparkline')}>
                Sparkline
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Illustrations">
              <Button size="sm" variant="ghost" icon="shape-square" onClick={() => onAction('open-shape')}>
                Shapes & Lines
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Functions">
              <Button size="sm" variant="ghost" icon="function" onClick={() => onAction('function-wizard')}>
                Insert Function (fx)
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Cells">
              <Button size="sm" variant="ghost" onClick={() => onAction('insert-row')}>Insert Row</Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('insert-column')}>Insert Column</Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('delete-row')}>Delete Row</Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('delete-column')}>Delete Column</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'data' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Sort & Filter">
              <Button size="sm" variant="ghost" icon="sort" onClick={() => onAction('sort-asc')}>
                Sort A to Z
              </Button>
              <Button size="sm" variant="ghost" icon="sort" onClick={() => onAction('sort-desc')}>
                Sort Z to A
              </Button>
              <Button size="sm" variant="ghost" icon="sliders" onClick={() => onAction('sort-dialog')}>
                Custom Sort...
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Data Tools">
              <Button size="sm" variant="ghost" icon="table" onClick={() => onAction('open-data-table')}>
                Data Model
              </Button>
              <Button size="sm" variant="ghost" icon="table" onClick={() => onAction('create-data-table')}>
                Create Data Table
              </Button>
              <Button size="sm" variant="ghost" icon="check-circle" onClick={() => onAction('open-data-validation')}>
                Data Validation
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('apply-filter-selection')}>
                Filter Selection
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('filter-clear')}>
                Clear Filter
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Find & Transform">
              <Button size="sm" variant="ghost" onClick={() => onAction('find-replace')}>
                Find & Replace
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('transpose')}>
                Transpose
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('flip-h')}>Flip H</Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('flip-v')}>Flip V</Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('split-column')}>Split by Delimiter</Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'review' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="History & Audit">
              <Button size="sm" variant="ghost" icon="history" onClick={() => onAction('open-history')}>
                Revision Log
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}

        {activeTab === 'view' ? (
          <Inline gap="md" className="min-w-max items-start">
            <RibbonGroup label="Freeze Panes">
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onAction('freeze-top-row')}>
                Freeze Top Row
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onAction('freeze-first-col')}>
                Freeze First Column
              </Button>
              <Button size="sm" variant="ghost" icon="freeze" onClick={() => onAction('unfreeze')}>
                Unfreeze All
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Zoom">
              <Button size="sm" variant="ghost" icon="zoom-in" onClick={() => onAction('zoom-in')}>
                Zoom In
              </Button>
              <Button size="sm" variant="ghost" icon="zoom-out" onClick={() => onAction('zoom-out')}>
                Zoom Out
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('zoom-100')}>
                100%
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Print Layout">
              <Button size="sm" variant="ghost" icon="printer" onClick={() => onAction('open-print')}>
                Print & PDF
              </Button>
            </RibbonGroup>
            <Divider orientation="vertical" className="h-10" />

            <RibbonGroup label="Appearance & Files">
              <Button size="sm" variant="ghost" onClick={() => onAction('banded-toggle')}>
                Banded Rows
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('export-xlsx')}>
                Export .xlsx
              </Button>
              <Button size="sm" variant="ghost" onClick={() => onAction('import-xlsx')}>
                Import .xlsx
              </Button>
            </RibbonGroup>
          </Inline>
        ) : null}
      </Box>
      </Tabs>
    </RibbonLocaleContext.Provider>
  );
}
