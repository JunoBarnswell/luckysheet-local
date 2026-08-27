import React from 'react';
import { Button, DropdownMenu, Inline, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { type RibbonCommandId, type RibbonSurfaceDefinition } from '@react-sheets/spreadsheet-app';
import type { ChartDrawingPayload, DrawingConnectorType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { Locale } from '../i18n';
import { insertText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { RibbonLayoutRenderer } from './RibbonLayoutRenderer';
import { INSERT_CHART_FAMILIES, INSERT_CONNECTOR_VARIANTS, INSERT_SHAPE_GALLERY, INSERT_SPARKLINE_VARIANTS, type InsertChartFamilyVariant } from './insert-ribbon-catalog';

export interface InsertRibbonProps {
  locale: Locale;
  layout: RibbonLayoutState;
  disabled: boolean;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
  onInsertChart: (type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype']) => void;
  onInsertSparkline: (type: SparklineModel['type']) => void;
  onInsertShape: (type: ShapeDrawingPayload['type']) => void;
  onInsertConnector: (type: DrawingConnectorType) => void;
}

function RibbonLarge({ children, compact = false, icon, disabled, surfaceId, title, onClick }: { children: React.ReactNode; compact?: boolean; icon: React.ComponentProps<typeof Button>['icon']; disabled?: boolean; surfaceId: string; title: string; onClick?: () => void }) {
  return <Button aria-label={title} data-ribbon-surface={surfaceId} title={title} disabled={disabled} icon={icon} onClick={onClick} size="sm" variant="ghost" className={compact ? '!h-6 !min-h-0 !w-6 rounded-none px-0 [&>svg]:!h-3 [&>svg]:!w-3' : '!h-[104px] !min-h-0 min-w-[42px] max-w-[64px] flex-col gap-1 overflow-hidden rounded-none px-1 text-center text-[13px] leading-4 !whitespace-normal break-words [&>svg]:!h-8 [&>svg]:!w-8 [&>svg]:!shrink-0'}>{compact ? null : children}</Button>;
}

function variantButton({ id, icon, label, onSelect, surfaceId, disabled }: { id: string; icon: React.ComponentProps<typeof Button>['icon']; label: string; onSelect: () => void; surfaceId: string; disabled?: boolean }) {
  return <Button key={id} aria-label={label} data-ribbon-surface={surfaceId} data-ribbon-variant={id} title={label} icon={icon} disabled={disabled} size="sm" variant="ghost" className="w-full justify-start" onClick={onSelect}>{label}</Button>;
}

/** Compact icon-only button used inside the chart-type icon grid. */
const CHART_ICON_BTN = '!h-7 !min-h-0 !w-7 rounded-none px-0 [&>svg]:!h-3.5 [&>svg]:!w-3.5';

const chartVariantLabel = (locale: Locale, variant: InsertChartFamilyVariant): string => {
  const labels: Partial<Record<ChartDrawingPayload['subtype'], readonly [string, string]>> = {
    clustered: ['簇状', 'Clustered'], stacked: ['堆积', 'Stacked'], 'percent-stacked': ['百分比堆积', '100% Stacked'], 'three-dimensional': ['三维', '3D'],
    line: ['折线', 'Line'], 'line-markers': ['带数据标记的折线', 'Line with Markers'], area: ['面积', 'Area'], pie: ['饼图', 'Pie'], 'pie-of-pie': ['复合饼图', 'Pie of Pie'], 'bar-of-pie': ['复合条饼图', 'Bar of Pie'], doughnut: ['圆环图', 'Doughnut'],
    'scatter-markers': ['仅带数据标记的散点图', 'Scatter with Markers'], 'scatter-smooth-lines': ['平滑线散点图', 'Scatter with Smooth Lines'], bubble: ['气泡图', 'Bubble'], treemap: ['树状图', 'Treemap'], sunburst: ['旭日图', 'Sunburst'], histogram: ['直方图', 'Histogram'], pareto: ['帕累托图', 'Pareto'], 'box-whisker': ['箱形图', 'Box & Whisker'],
    waterfall: ['瀑布图', 'Waterfall'], funnel: ['漏斗图', 'Funnel'], 'stock-high-low-close': ['盘高-盘低-收盘图', 'High-Low-Close'], 'surface-three-dimensional': ['三维曲面图', '3D Surface'], radar: ['雷达图', 'Radar'], 'custom-combo': ['自定义组合图', 'Custom Combo'], 'filled-map': ['填充地图', 'Filled Map'],
  };
  const label = labels[variant.subtype];
  return label ? label[locale === 'zh-CN' ? 0 : 1] : variant.subtype;
};

export function InsertRibbon({ locale, layout, disabled, renderCommand, onInsertChart, onInsertSparkline, onInsertShape, onInsertConnector }: InsertRibbonProps) {
  const isNarrow = layout.mode === 'narrow';

  /**
   * SpreadJS parity: for chartBuilder in wide mode, render a 2-row × 4-column grid
   * of individual chart-type icon buttons directly in the ribbon (no dropdown).
   * Each button inserts that chart type immediately on click.
   */
  const renderChartIconGrid = (surfaceId: string): React.ReactNode => {
    const row1 = INSERT_CHART_FAMILIES.slice(0, 5);
    const row2 = INSERT_CHART_FAMILIES.slice(5);
    const familyControl = (family: typeof INSERT_CHART_FAMILIES[number]) => {
      const primary = family.variants[0]!;
      const familyLabel = insertText(locale, family.labelKey);
      return <Inline key={family.id} gap="none" className="items-stretch">
        <Button aria-label={familyLabel} data-ribbon-surface={surfaceId} data-ribbon-variant={family.id} title={familyLabel} icon={family.icon} iconOnly disabled={disabled} size="sm" variant="ghost" className={CHART_ICON_BTN} onClick={() => onInsertChart(primary.chartType, primary.subtype)} />
        <DropdownMenu align="left" trigger={<Button aria-label={`${familyLabel} options`} icon="chevron-down" iconOnly disabled={disabled} size="sm" variant="ghost" className="!h-7 !min-h-0 !w-3 rounded-none px-0 [&>svg]:!h-2.5 [&>svg]:!w-2.5" />}>
          <Stack gap="none" className="min-w-[13rem] p-1">{family.variants.map((variant) => variantButton({ id: variant.id, icon: family.icon, label: chartVariantLabel(locale, variant), disabled, onSelect: () => onInsertChart(variant.chartType, variant.subtype), surfaceId }))}</Stack>
        </DropdownMenu>
      </Inline>;
    };
    return (
      <Stack key={surfaceId} gap="none" data-ribbon-surface={surfaceId} className="items-center justify-center">
        <Inline gap="none" className="flex-nowrap">
          {row1.map(familyControl)}
        </Inline>
        <Inline gap="none" className="flex-nowrap">
          {row2.map(familyControl)}
        </Inline>
      </Stack>
    );
  };

  const renderSplitGallery = (surface: RibbonSurfaceDefinition, title: string, icon: React.ComponentProps<typeof Button>['icon'], variants: React.ReactNode[], onSelect: () => void): React.ReactNode => (
    <Inline key={surface.id} gap="none" className="items-stretch">
      <RibbonLarge compact={isNarrow} disabled={disabled} icon={icon} onClick={onSelect} surfaceId={surface.id} title={title}>
        {title}
      </RibbonLarge>
      <DropdownMenu align="left" trigger={<Button aria-label={`${title} options`} data-ribbon-surface={`${surface.id}.menu`} title={`${title} options`} disabled={disabled} icon="chevron-down" iconOnly size="sm" variant="ghost" className={isNarrow ? '!h-7 !min-h-0 !w-5 rounded-none px-0 [&>svg]:!h-3.5 [&>svg]:!w-3.5' : '!h-[104px] !min-h-0 !w-5 rounded-none px-0 [&>svg]:!h-3 [&>svg]:!w-3'} />}>
        <Stack gap="none" className="min-w-[14rem] p-1">{variants}</Stack>
      </DropdownMenu>
    </Inline>
  );

  const galleryItems = (commandId: RibbonCommandId, surfaceId: string): React.ReactNode[] => {
    if (commandId === 'chartBuilder') return INSERT_CHART_FAMILIES.flatMap((family) => family.variants.map((variant) => variantButton({ id: variant.id, icon: family.icon, label: chartVariantLabel(locale, variant), disabled, onSelect: () => onInsertChart(variant.chartType, variant.subtype), surfaceId })));
    if (commandId === 'sparkline') return INSERT_SPARKLINE_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertSparkline(variant.value), surfaceId }));
    if (commandId === 'shapesLines') return INSERT_SHAPE_GALLERY.flatMap((category) => [
      <Text key={`${category.id}.label`} size="xs" weight="semibold" className="px-2 pb-1 pt-2 text-slate-500">{insertText(locale, category.labelKey)}</Text>,
      ...category.variants.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertShape(variant.value), surfaceId })),
      <Text key="connectors.label" size="xs" weight="semibold" className="px-2 pb-1 pt-2 text-slate-500">{insertText(locale, 'connectorCategory')}</Text>,
      ...INSERT_CONNECTOR_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertConnector(variant.value), surfaceId })),
    ]);
    return [];
  };

  const renderSurface = (surface: RibbonSurfaceDefinition, mode: 'wide' | 'menu'): React.ReactNode => {
    if (!surface.commandId) return null;

    // SpreadJS parity: chartBuilder in wide mode → 2×4 icon grid instead of dropdown tile.
    if (surface.commandId === 'chartBuilder' && mode === 'wide' && !isNarrow) {
      return renderChartIconGrid(surface.id);
    }

    const variants = galleryItems(surface.commandId, surface.id);
    if (variants.length > 0) {
      const title = insertText(locale, surface.commandId === 'chartBuilder' ? 'chart' : surface.commandId === 'sparkline' ? 'sparkline' : 'shape');
      if (mode === 'menu') return <React.Fragment key={surface.id}>{variants}</React.Fragment>;
      if (surface.commandId === 'sparkline') {
        const first = INSERT_SPARKLINE_VARIANTS[0];
        return renderSplitGallery(surface, title, 'sparkline', variants, () => onInsertSparkline(first.value));
      }
      const icon = surface.commandId === 'chartBuilder' ? 'chart-column' : 'shape-square';
      return <DropdownMenu key={surface.id} align="left" trigger={<RibbonLarge compact={isNarrow} disabled={disabled} icon={icon} surfaceId={surface.id} title={title}>{title}</RibbonLarge>}><Stack gap="none" className="min-w-[14rem] p-1">{variants}</Stack></DropdownMenu>;
    }

    // SpreadJS parity: 'large', 'tile', 'gallery', and 'split' all render as full-height tiles.
    // Previously 'split' was excluded, causing picture and worksheet-table to render as small buttons.
    const tile = surface.appearance === 'large' || surface.appearance === 'gallery' || surface.appearance === 'tile' || surface.appearance === 'split';
    return <React.Fragment key={surface.id}>{renderCommand(surface.commandId, mode === 'menu' ? { className: 'w-full justify-start', ribbonSurfaceId: surface.id } : { tile: !isNarrow && tile, ribbonSurfaceId: surface.id })}</React.Fragment>;
  };

  return <RibbonLayoutRenderer tab="insert" locale={locale} layout={layout} renderCommand={renderCommand} renderSurface={(surface, context) => renderSurface(surface, context.mode)} />;
}
