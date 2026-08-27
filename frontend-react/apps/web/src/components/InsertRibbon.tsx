import React from 'react';
import { AssetIcon, Button, DropdownMenu, Inline, Stack, Text, type IconName, type RibbonLayoutState } from '@react-sheets/ui-system';
import { type RibbonCommandId, type RibbonSurfaceDefinition } from '@react-sheets/spreadsheet-app';
import type { ChartDrawingPayload, DrawingConnectorType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { Locale } from '../i18n';
import { insertText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { RibbonLayoutRenderer } from './RibbonLayoutRenderer';
import { INSERT_CHART_FAMILIES, INSERT_CONNECTOR_VARIANTS, INSERT_SHAPE_GALLERY, INSERT_SPARKLINE_VARIANTS, type InsertChartFamilyDefinition, type InsertChartFamilyVariant } from './insert-ribbon-catalog';

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

const FLUENT_ASSETS: Partial<Record<IconName, string>> = { table: '/icons/fluent/ic_fluent_table_24_regular.svg', 'form-control': '/icons/fluent/ic_fluent_form_24_regular.svg', picture: '/icons/fluent/ic_fluent_image_24_regular.svg', 'shape-square': '/icons/fluent/ic_fluent_shapes_24_regular.svg', 'shape-circle': '/icons/fluent/ic_fluent_shapes_24_regular.svg', chart: '/icons/fluent/ic_fluent_chart_multiple_24_regular.svg', 'chart-column': '/icons/fluent/ic_fluent_data_bar_vertical_24_regular.svg', 'chart-line': '/icons/fluent/ic_fluent_data_line_24_regular.svg', 'chart-pie': '/icons/fluent/ic_fluent_data_pie_24_regular.svg', 'chart-scatter': '/icons/fluent/ic_fluent_data_scatter_24_regular.svg', filter: '/icons/fluent/ic_fluent_filter_24_regular.svg', link: '/icons/fluent/ic_fluent_link_24_regular.svg', comment: '/icons/fluent/ic_fluent_comment_24_regular.svg', checkbox: '/icons/fluent/ic_fluent_form_24_regular.svg', camera: '/icons/fluent/ic_fluent_screenshot_24_regular.svg' };
const FLUENT_SURFACE_ASSETS: Record<string, string> = { 'illustrations.icons': '/icons/fluent/ic_fluent_icons_24_regular.svg', 'illustrations.models3d': '/icons/fluent/ic_fluent_cube_24_regular.svg', 'illustrations.smartart': '/icons/fluent/ic_fluent_flowchart_24_regular.svg', 'illustrations.screenshot': '/icons/fluent/ic_fluent_screenshot_24_regular.svg', 'tables.forms': '/icons/fluent/ic_fluent_form_24_regular.svg', 'controls.checkbox': '/icons/fluent/ic_fluent_form_24_regular.svg', 'filters.timeline': '/icons/fluent/ic_fluent_timeline_24_regular.svg' };
function fluentIcon(icon: React.ComponentProps<typeof Button>['icon'], size: 'sm' | 'md' | 'lg' | 'xl' = 'lg'): React.ReactNode { const src = typeof icon === 'string' ? FLUENT_ASSETS[icon] : undefined; return src ? <AssetIcon src={src} size={size} aria-hidden="true" /> : undefined; }
function fluentSurfaceIcon(surfaceId: string, size: 'sm' | 'md' | 'lg' | 'xl' = 'lg'): React.ReactNode { const src = FLUENT_SURFACE_ASSETS[surfaceId]; return src ? <AssetIcon src={src} size={size} aria-hidden="true" /> : undefined; }
function RibbonLarge({ children, compact = false, icon, iconNode, disabled, surfaceId, title, onClick, className }: { children: React.ReactNode; compact?: boolean; icon?: React.ComponentProps<typeof Button>['icon']; iconNode?: React.ReactNode; disabled?: boolean; surfaceId: string; title: string; onClick?: () => void; className?: string }) {
  return <Button aria-label={title} data-ribbon-surface={surfaceId} title={title} disabled={disabled} icon={iconNode ? undefined : icon} iconNode={iconNode} onClick={onClick} size="sm" variant="ghost" className={`${compact ? '!h-6 !min-h-0 !w-6 rounded-none px-0 [&>svg]:!h-3 [&>svg]:!w-3' : '!h-[104px] !min-h-0 min-w-[42px] max-w-[64px] flex-col gap-1 overflow-hidden rounded-none px-1 text-center text-[13px] leading-4 !whitespace-normal break-words [&>svg]:!h-8 [&>svg]:!w-8 [&>img]:!h-8 [&>img]:!w-8 [&>img]:!shrink-0'} ${className ?? ''}`}>{compact ? null : children}</Button>;
}

function variantButton({ id, icon, label, onSelect, surfaceId, disabled }: { id: string; icon: React.ComponentProps<typeof Button>['icon']; label: string; onSelect: () => void; surfaceId: string; disabled?: boolean }) {
  const iconNode = fluentIcon(icon, 'md');
  return <Button key={id} aria-label={label} data-ribbon-surface={surfaceId} data-ribbon-variant={id} title={label} icon={iconNode ? undefined : icon} iconNode={iconNode} disabled={disabled} size="sm" variant="ghost" className="w-full justify-start" onClick={onSelect}>{label}</Button>;
}

/** Compact icon-only button used inside the chart-type icon grid. */
const CHART_ICON_BTN = '!h-8 !min-h-0 !w-12 rounded-none px-0 [&>svg]:!h-5 [&>svg]:!w-5 [&>img]:!h-5 [&>img]:!w-5';

const chartVariantLabel = (locale: Locale, variant: InsertChartFamilyVariant): string => {
  const labels: Partial<Record<ChartDrawingPayload['subtype'], readonly [string, string]>> = {
    clustered: ['簇状', 'Clustered'], stacked: ['堆积', 'Stacked'], 'percent-stacked': ['百分比堆积', '100% Stacked'], 'three-dimensional': ['三维', '3D'], 'three-dimensional-stacked': ['三维堆积', '3D Stacked'], 'three-dimensional-percent-stacked': ['三维百分比堆积', '3D 100% Stacked'],
    line: ['折线', 'Line'], 'line-markers': ['带数据标记的折线', 'Line with Markers'], area: ['面积', 'Area'], pie: ['饼图', 'Pie'], 'pie-of-pie': ['复合饼图', 'Pie of Pie'], 'bar-of-pie': ['复合条饼图', 'Bar of Pie'], doughnut: ['圆环图', 'Doughnut'],
    'scatter-markers': ['仅带数据标记的散点图', 'Scatter with Markers'], 'scatter-smooth-lines': ['平滑线散点图', 'Scatter with Smooth Lines'], bubble: ['气泡图', 'Bubble'], treemap: ['树状图', 'Treemap'], sunburst: ['旭日图', 'Sunburst'], histogram: ['直方图', 'Histogram'], pareto: ['帕累托图', 'Pareto'], 'box-whisker': ['箱形图', 'Box & Whisker'],
    waterfall: ['瀑布图', 'Waterfall'], funnel: ['漏斗图', 'Funnel'], 'stock-high-low-close': ['盘高-盘低-收盘图', 'High-Low-Close'], 'surface-three-dimensional': ['三维曲面图', '3D Surface'], radar: ['雷达图', 'Radar'], 'custom-combo': ['自定义组合图', 'Custom Combo'], 'filled-map': ['填充地图', 'Filled Map'],
  };
  const label = labels[variant.subtype];
  return label ? label[locale === 'zh-CN' ? 0 : 1] : variant.subtype;
};

const chartGallerySections = (family: InsertChartFamilyDefinition): readonly { title: [string, string]; variants: readonly InsertChartFamilyVariant[] }[] => {
  const threeD = (v: InsertChartFamilyVariant) => v.subtype.startsWith('three-dimensional');
  if (family.id === 'chart-family.column-bar') return [
    { title: ['二维柱形图', '2-D Column'], variants: family.variants.filter((v) => v.chartType === 'column' && !threeD(v)) },
    { title: ['三维柱形图', '3-D Column'], variants: family.variants.filter((v) => v.chartType === 'column' && threeD(v)) },
    { title: ['二维条形图', '2-D Bar'], variants: family.variants.filter((v) => v.chartType === 'bar' && !threeD(v)) },
    { title: ['三维条形图', '3-D Bar'], variants: family.variants.filter((v) => v.chartType === 'bar' && threeD(v)) },
  ];
  if (family.id === 'chart-family.line-area') return [
    { title: ['二维折线图', '2-D Line'], variants: family.variants.filter((v) => v.chartType === 'line' && !threeD(v)) },
    { title: ['三维折线图', '3-D Line'], variants: family.variants.filter((v) => v.chartType === 'line' && threeD(v)) },
    { title: ['二维面积图', '2-D Area'], variants: family.variants.filter((v) => v.chartType === 'area' && !threeD(v)) },
    { title: ['三维面积图', '3-D Area'], variants: family.variants.filter((v) => v.chartType === 'area' && threeD(v)) },
  ];
  return [{ title: [insertText('zh-CN', family.labelKey), insertText('en-US', family.labelKey)], variants: family.variants }];
};

function chartFamilyMenu(locale: Locale, family: InsertChartFamilyDefinition, disabled: boolean, surfaceId: string, onInsertChart: InsertRibbonProps['onInsertChart']): React.ReactNode {
  return <Stack gap="none" className="w-[284px] max-w-[284px] overflow-hidden rounded-md border border-slate-300 bg-white p-0 shadow-lg">
    {chartGallerySections(family).map((section, index) => section.variants.length > 0 ? <Stack key={`${family.id}.${index}`} gap="none" className={`${index > 0 ? 'border-t border-slate-200' : ''} px-3 pb-3 pt-3`}>
      <Text size="sm" weight="semibold" className="mb-1 text-slate-800">{section.title[locale === 'zh-CN' ? 0 : 1]}</Text>
      <Inline gap="none" className="flex-wrap items-start">
        {section.variants.map((variant) => { const icon = variant.chartType === 'line' ? 'chart-line' : variant.chartType === 'area' ? 'chart-area' : variant.chartType === 'pie' || variant.chartType === 'doughnut' ? 'chart-pie' : variant.chartType === 'scatter' || variant.chartType === 'bubble' ? 'chart-scatter' : family.icon; const node = fluentIcon(icon, 'lg'); return <Button key={variant.id} aria-label={chartVariantLabel(locale, variant)} data-ribbon-surface={surfaceId} data-ribbon-variant={variant.id} title={chartVariantLabel(locale, variant)} disabled={disabled} icon={node ? undefined : icon} iconNode={node} size="sm" variant="ghost" className="!h-[68px] !w-[58px] !min-w-[58px] flex-col gap-0.5 rounded-none px-0 text-[10px] leading-3 [&>img]:!h-9 [&>img]:!w-9 [&>svg]:!h-9 [&>svg]:!w-9" onClick={() => onInsertChart(variant.chartType, variant.subtype)}><Text size="xs" className="max-w-[56px] truncate">{chartVariantLabel(locale, variant)}</Text></Button>; })}
      </Inline>
    </Stack> : null)}
    <Button aria-label={locale === 'zh-CN' ? '更多图表' : 'More Charts'} iconNode={fluentIcon('chart', 'md')} iconOnly={false} disabled={disabled} size="sm" variant="ghost" className="!h-8 !w-full justify-start rounded-none border-t border-slate-200 px-3 text-xs">{locale === 'zh-CN' ? '更多图表(M)...' : 'More Charts...'}</Button>
  </Stack>;
}

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
        <Button aria-label={familyLabel} data-ribbon-surface={surfaceId} data-ribbon-variant={family.id} title={familyLabel} icon={fluentIcon(family.icon, 'md') ? undefined : family.icon} iconNode={fluentIcon(family.icon, 'md')} iconOnly disabled={disabled} size="sm" variant="ghost" className={CHART_ICON_BTN} onClick={() => onInsertChart(primary.chartType, primary.subtype)} />
        <DropdownMenu align="left" trigger={<Button aria-label={`${familyLabel} options`} icon="chevron-down" iconOnly disabled={disabled} size="sm" variant="ghost" className="!h-8 !min-h-0 !w-4 rounded-none px-0 [&>svg]:!h-3 [&>svg]:!w-3" />}>
          {chartFamilyMenu(locale, family, disabled, surfaceId, onInsertChart)}
        </DropdownMenu>
      </Inline>;
    };
    return (
      <Stack key={surfaceId} gap="none" data-ribbon-surface={surfaceId} className="!w-[320px] !min-w-[320px] shrink-0 items-center justify-center">
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

  const renderSurface = (surface: RibbonSurfaceDefinition, mode: RibbonLayoutState['mode'] | 'menu'): React.ReactNode => {
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
        if (mode === 'wide' && !isNarrow) return <Inline key={surface.id} gap="none" className="h-[104px] w-[181px] min-w-[181px] items-stretch justify-center">{INSERT_SPARKLINE_VARIANTS.map((variant) => <RibbonLarge key={variant.id} disabled={disabled} icon={fluentIcon(variant.icon, 'lg') ? undefined : variant.icon} iconNode={fluentIcon(variant.icon, 'lg')} surfaceId={variant.id} title={insertText(locale, variant.labelKey)} className="!w-[58px] !min-w-[58px] !max-w-[58px]" onClick={() => onInsertSparkline(variant.value)}>{insertText(locale, variant.labelKey)}</RibbonLarge>)}</Inline>;
        const first = INSERT_SPARKLINE_VARIANTS[0];
        return renderSplitGallery(surface, title, 'sparkline', variants, () => onInsertSparkline(first.value));
      }
      const icon = surface.commandId === 'chartBuilder' ? 'chart-column' : 'shape-square';
      return <DropdownMenu key={surface.id} align="left" trigger={<RibbonLarge compact={isNarrow} disabled={disabled} icon={fluentIcon(icon, 'lg') ? undefined : icon} iconNode={fluentIcon(icon, 'lg')} surfaceId={surface.id} title={title}>{title}</RibbonLarge>}><Stack gap="none" className="min-w-[14rem] p-1">{variants}</Stack></DropdownMenu>;
    }

    // SpreadJS parity: 'large', 'tile', 'gallery', and 'split' all render as full-height tiles.
    // Previously 'split' was excluded, causing picture and worksheet-table to render as small buttons.
    const tile = surface.appearance === 'large' || surface.appearance === 'gallery' || surface.appearance === 'tile' || surface.appearance === 'split';
    const icon = surface.commandId === 'picture' ? 'picture' : surface.commandId === 'shapesLines' ? 'shape-square' : undefined;
    const iconNode = FLUENT_SURFACE_ASSETS[surface.id] ? fluentSurfaceIcon(surface.id, 'lg') : icon ? fluentIcon(icon, 'lg') : undefined;
    const illustrationWidth = surface.id.startsWith('illustrations.') ? '!w-[50px] !min-w-[50px] !max-w-[50px]' : undefined;
    return <React.Fragment key={surface.id}>{renderCommand(surface.commandId, mode === 'menu' ? { className: 'w-full justify-start', iconNode, ribbonSurfaceId: surface.id } : { tile: !isNarrow && tile, className: illustrationWidth, iconNode, ribbonSurfaceId: surface.id })}</React.Fragment>;
  };

  return <RibbonLayoutRenderer tab="insert" locale={locale} layout={layout} renderCommand={renderCommand} renderSurface={(surface, context) => renderSurface(surface, context.mode)} />;
}
