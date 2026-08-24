import React from 'react';
import { Button, Divider, DropdownMenu, Inline, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { getRibbonGroupDefinition, type RibbonCommandId } from '@react-sheets/spreadsheet-app';
import type { ChartDrawingPayload, FormControlType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';

export interface InsertRibbonProps {
  locale: Locale;
  layout: RibbonLayoutState;
  disabled: boolean;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
  onInsertChart: (type: ChartDrawingPayload['chartType']) => void;
  onInsertSparkline: (type: SparklineModel['type']) => void;
  onInsertShape: (type: ShapeDrawingPayload['type']) => void;
  onInsertFormControl: (type: FormControlType) => void;
}

type InsertGroupId = 'insertSheets' | 'insertTables' | 'insertCharts' | 'insertDataCharts' | 'illustrations' | 'insertLinks' | 'insertControls';

const widths: Record<InsertGroupId, string> = {
  insertSheets: 'w-[194px]', insertTables: 'w-[160px]', insertCharts: 'w-[350px]', insertDataCharts: 'w-[82px]', illustrations: 'w-[282px]', insertLinks: 'w-[70px]', insertControls: 'w-[132px]',
};

function InsertGroup({ children, id, locale }: { children: React.ReactNode; id: InsertGroupId; locale: Locale }) {
  return (
    <Stack gap="none" className={`h-[102px] shrink-0 justify-between ${widths[id]}`}>
      <Inline gap="none" className="min-h-0 flex-1 items-start justify-center pt-2">{children}</Inline>
      <Text size="xs" tone="subtle" className="h-4 shrink-0 text-center text-[10px] font-medium text-[#413c40]">{translateRibbonText(locale, getRibbonGroupDefinition(id).labelKey)}</Text>
    </Stack>
  );
}

function RibbonLarge({ children, icon, onClick, disabled, title }: { children: React.ReactNode; icon: React.ComponentProps<typeof Button>['icon']; onClick?: () => void; disabled?: boolean; title: string }) {
  return <Button aria-label={title} title={title} disabled={disabled} icon={icon} size="sm" variant="ghost" className="!h-[72px] !min-h-0 !w-[64px] flex-col gap-1 rounded-none px-1 text-[12px] leading-4 [&>svg]:!h-7 [&>svg]:!w-7" onClick={onClick}>{children}</Button>;
}

export function InsertRibbon({ locale, layout, disabled, renderCommand, onInsertChart, onInsertSparkline, onInsertShape, onInsertFormControl }: InsertRibbonProps) {
  if (layout.width < 1024) {
    const commandIds: RibbonCommandId[] = ['tableSheet', 'ganttSheet', 'reportSheet', 'worksheetTable', 'pivotTable', 'chartBuilder', 'barcode', 'sparkline', 'dataChart', 'shapesLines', 'camera', 'formControls', 'hyperlink', 'checkbox', 'textbox'];
    return <Inline gap="xs" className="h-[96px] items-center overflow-hidden">{commandIds.map((id) => <React.Fragment key={id}>{renderCommand(id, { iconOnly: true, className: '!h-10 !w-10' })}</React.Fragment>)}</Inline>;
  }
  return (
    <Inline gap="none" className="h-[102px] min-w-max flex-nowrap items-start overflow-hidden" data-testid="insert-ribbon-groups">
      <InsertGroup id="insertSheets" locale={locale}>{renderCommand('tableSheet', { tile: true })}{renderCommand('ganttSheet', { tile: true })}{renderCommand('reportSheet', { tile: true })}</InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertTables" locale={locale}>{renderCommand('worksheetTable', { tile: true })}{renderCommand('pivotTable', { tile: true })}</InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertCharts" locale={locale}>
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="chart-column" title="图表">图表</RibbonLarge>}>
          <Stack gap="none" className="min-w-[15rem] p-1">
            {([['column', 'chart-column', '柱形图'], ['bar', 'chart-bar', '条形图'], ['line', 'chart-line', '折线图'], ['area', 'chart-area', '面积图'], ['pie', 'chart-pie', '饼图'], ['scatter', 'chart-scatter', '散点图'], ['combo', 'data-chart', '组合图']] as const).map(([type, icon, label]) => <Button key={type} icon={icon} size="sm" variant="ghost" className="justify-start" onClick={() => onInsertChart(type)}>{label}</Button>)}
          </Stack>
        </DropdownMenu>
        <Stack gap="none" className="w-[116px] px-1 pt-1">
          <Inline gap="none"><Button aria-label="柱形图" icon="chart-column" iconOnly size="sm" variant="ghost" onClick={() => onInsertChart('column')} /><Button aria-label="条形图" icon="chart-bar" iconOnly size="sm" variant="ghost" onClick={() => onInsertChart('bar')} /><Button aria-label="折线图" icon="chart-line" iconOnly size="sm" variant="ghost" onClick={() => onInsertChart('line')} /></Inline>
          <Inline gap="none"><Button aria-label="面积图" icon="chart-area" iconOnly size="sm" variant="ghost" onClick={() => onInsertChart('area')} /><Button aria-label="饼图" icon="chart-pie" iconOnly size="sm" variant="ghost" onClick={() => onInsertChart('pie')} /><Button aria-label="散点图" icon="chart-scatter" iconOnly size="sm" variant="ghost" onClick={() => onInsertChart('scatter')} /></Inline>
        </Stack>
        {renderCommand('barcode', { tile: true })}
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="sparkline" title="迷你图">迷你图</RibbonLarge>}>
          <Stack gap="none" className="min-w-[10rem] p-1"><Button icon="chart-line" size="sm" variant="ghost" onClick={() => onInsertSparkline('line')}>折线</Button><Button icon="chart-column" size="sm" variant="ghost" onClick={() => onInsertSparkline('column')}>柱形</Button><Button icon="chart-bar" size="sm" variant="ghost" onClick={() => onInsertSparkline('win-loss')}>盈亏</Button></Stack>
        </DropdownMenu>
      </InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertDataCharts" locale={locale}>{renderCommand('dataChart', { tile: true })}</InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="illustrations" locale={locale}>
        {renderCommand('picture', { tile: true })}
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="shape-square" title="形状">形状</RibbonLarge>}><Stack gap="none" className="min-w-[12rem] p-1">{(['rectangle', 'rounded-rectangle', 'ellipse', 'line', 'arrow', 'callout', 'star'] as const).map((type) => <Button key={type} icon={type === 'ellipse' ? 'shape-circle' : 'shape-square'} size="sm" variant="ghost" className="justify-start" onClick={() => onInsertShape(type)}>{type}</Button>)}</Stack></DropdownMenu>
        {renderCommand('camera', { tile: true })}
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="form-control" title="控件">控件</RibbonLarge>}><Stack gap="none" className="min-w-[13rem] p-1">{(['button', 'spin-button', 'list-box', 'combo-box', 'checkbox', 'option-button', 'group-box', 'label', 'scrollbar'] as const).map((type) => <Button key={type} icon={type === 'checkbox' ? 'checkbox' : 'form-control'} size="sm" variant="ghost" className="justify-start" onClick={() => onInsertFormControl(type)}>{type}</Button>)}</Stack></DropdownMenu>
      </InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertLinks" locale={locale}>{renderCommand('hyperlink', { tile: true })}</InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertControls" locale={locale}>{renderCommand('checkbox', { tile: true })}{renderCommand('textbox', { tile: true })}</InsertGroup>
    </Inline>
  );
}
