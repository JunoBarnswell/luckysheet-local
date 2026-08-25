import React from 'react';
import { Button, Divider, DropdownMenu, Inline, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { getRibbonGroupDefinition, type RibbonCommandId } from '@react-sheets/spreadsheet-app';
import type { ChartDrawingPayload, FormControlType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { Locale } from '../i18n';
import { insertText, translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { INSERT_CHART_VARIANTS, INSERT_FORM_CONTROL_VARIANTS, INSERT_SHAPE_VARIANTS, INSERT_SPARKLINE_VARIANTS } from './insert-ribbon-catalog';

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
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="chart-column" title={insertText(locale, 'chart')}>{insertText(locale, 'chart')}</RibbonLarge>}>
          <Stack gap="none" className="min-w-[15rem] p-1">
            {INSERT_CHART_VARIANTS.map((variant) => { const label = insertText(locale, variant.labelKey); return <Button key={variant.id} aria-label={insertText(locale, variant.ariaLabelKey)} title={insertText(locale, variant.tooltipKey)} icon={variant.icon} size="sm" variant="ghost" className="justify-start" onClick={() => onInsertChart(variant.value)}>{label}</Button>; })}
          </Stack>
        </DropdownMenu>
        <Stack gap="none" className="w-[116px] px-1 pt-1">
          <Inline gap="none">{INSERT_CHART_VARIANTS.slice(0, 3).map((variant) => <Button key={variant.id} aria-label={insertText(locale, variant.ariaLabelKey)} title={insertText(locale, variant.tooltipKey)} icon={variant.icon} iconOnly size="sm" variant="ghost" onClick={() => onInsertChart(variant.value)} />)}</Inline>
          <Inline gap="none">{INSERT_CHART_VARIANTS.slice(3, 6).map((variant) => <Button key={variant.id} aria-label={insertText(locale, variant.ariaLabelKey)} title={insertText(locale, variant.tooltipKey)} icon={variant.icon} iconOnly size="sm" variant="ghost" onClick={() => onInsertChart(variant.value)} />)}</Inline>
        </Stack>
        {renderCommand('barcode', { tile: true })}
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="sparkline" title={insertText(locale, 'sparkline')}>{insertText(locale, 'sparkline')}</RibbonLarge>}>
          <Stack gap="none" className="min-w-[10rem] p-1">{INSERT_SPARKLINE_VARIANTS.map((variant) => { const label = insertText(locale, variant.labelKey); return <Button key={variant.id} aria-label={insertText(locale, variant.ariaLabelKey)} title={insertText(locale, variant.tooltipKey)} icon={variant.icon} size="sm" variant="ghost" onClick={() => onInsertSparkline(variant.value)}>{label}</Button>; })}</Stack>
        </DropdownMenu>
      </InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertDataCharts" locale={locale}>{renderCommand('dataChart', { tile: true })}</InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="illustrations" locale={locale}>
        {renderCommand('picture', { tile: true })}
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="shape-square" title={insertText(locale, 'shape')}>{insertText(locale, 'shape')}</RibbonLarge>}><Stack gap="none" className="min-w-[12rem] p-1">{INSERT_SHAPE_VARIANTS.map((variant) => { const label = insertText(locale, variant.labelKey); return <Button key={variant.id} aria-label={insertText(locale, variant.ariaLabelKey)} title={insertText(locale, variant.tooltipKey)} icon={variant.icon} size="sm" variant="ghost" className="justify-start" onClick={() => onInsertShape(variant.value)}>{label}</Button>; })}</Stack></DropdownMenu>
        {renderCommand('camera', { tile: true })}
        <DropdownMenu align="left" trigger={<RibbonLarge disabled={disabled} icon="form-control" title={insertText(locale, 'formControl')}>{insertText(locale, 'formControl')}</RibbonLarge>}><Stack gap="none" className="min-w-[13rem] p-1">{INSERT_FORM_CONTROL_VARIANTS.map((variant) => { const label = insertText(locale, variant.labelKey); return <Button key={variant.id} aria-label={insertText(locale, variant.ariaLabelKey)} title={insertText(locale, variant.tooltipKey)} icon={variant.icon} size="sm" variant="ghost" className="justify-start" onClick={() => onInsertFormControl(variant.value)}>{label}</Button>; })}</Stack></DropdownMenu>
      </InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertLinks" locale={locale}>{renderCommand('hyperlink', { tile: true })}</InsertGroup>
      <Divider orientation="vertical" className="h-[96px]" />
      <InsertGroup id="insertControls" locale={locale}>{renderCommand('checkbox', { tile: true })}{renderCommand('textbox', { tile: true })}</InsertGroup>
    </Inline>
  );
}
