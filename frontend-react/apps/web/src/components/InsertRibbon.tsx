import React, { useMemo } from 'react';
import { Button, Divider, DropdownMenu, Inline, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { getRibbonGroupDefinition, getRibbonSurfaces, type RibbonCommandId, type RibbonGroupId, type RibbonSurfaceBreakpoint, type RibbonSurfaceDefinition } from '@react-sheets/spreadsheet-app';
import type { BarcodeSymbology, ChartDrawingPayload, DataChartPlotType, FormControlType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { Locale } from '../i18n';
import { insertText, translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { INSERT_BARCODE_VARIANTS, INSERT_CHART_VARIANTS, INSERT_DATA_CHART_VARIANTS, INSERT_FORM_CONTROL_VARIANTS, INSERT_SHAPE_VARIANTS, INSERT_SPARKLINE_VARIANTS } from './insert-ribbon-catalog';

export interface InsertRibbonProps {
  locale: Locale;
  layout: RibbonLayoutState;
  disabled: boolean;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
  onInsertChart: (type: ChartDrawingPayload['chartType']) => void;
  onInsertDataChart: (type: DataChartPlotType) => void;
  onInsertBarcode: (symbology: BarcodeSymbology) => void;
  onInsertSparkline: (type: SparklineModel['type']) => void;
  onInsertShape: (type: ShapeDrawingPayload['type']) => void;
  onInsertFormControl: (type: FormControlType) => void;
}

const INSERT_GROUPS = ['insertSheets', 'insertTables', 'insertCharts', 'insertDataCharts', 'illustrations', 'insertLinks', 'insertControls'] as const satisfies readonly RibbonGroupId[];

function breakpointFor(layout: RibbonLayoutState): RibbonSurfaceBreakpoint {
  if (layout.width >= 1280) return 'wide';
  if (layout.width >= 1024) return 'compact';
  return 'narrow';
}

function RibbonLarge({ children, icon, disabled, title }: { children: React.ReactNode; icon: React.ComponentProps<typeof Button>['icon']; disabled?: boolean; title: string }) {
  return <Button aria-label={title} title={title} disabled={disabled} icon={icon} size="sm" variant="ghost" className="!h-[68px] !min-h-0 !w-[68px] flex-col gap-1 rounded-none px-1 text-[11px] leading-4 [&>svg]:!h-6 [&>svg]:!w-6">{children}</Button>;
}

function variantButton({ id, icon, label, onSelect, disabled }: { id: string; icon: React.ComponentProps<typeof Button>['icon']; label: string; onSelect: () => void; disabled?: boolean }) {
  return <Button key={id} aria-label={label} title={label} icon={icon} disabled={disabled} size="sm" variant="ghost" className="w-full justify-start" onClick={onSelect}>{label}</Button>;
}

export function InsertRibbon({ locale, layout, disabled, renderCommand, onInsertChart, onInsertDataChart, onInsertBarcode, onInsertSparkline, onInsertShape, onInsertFormControl }: InsertRibbonProps) {
  const breakpoint = breakpointFor(layout);
  const surfacesByGroup = useMemo(() => new Map(INSERT_GROUPS.map((group) => [group, getRibbonSurfaces('insert', group, breakpoint)] as const)), [breakpoint]);

  const allSurfaces = (group: RibbonGroupId): readonly RibbonSurfaceDefinition[] => {
    const seen = new Set<string>();
    return (['wide', 'compact', 'narrow'] as const).flatMap((candidate) => getRibbonSurfaces('insert', group, candidate)).filter((surface) => {
      if (seen.has(surface.id)) return false;
      seen.add(surface.id);
      return true;
    });
  };

  const galleryItems = (commandId: RibbonCommandId): React.ReactNode[] => {
    if (commandId === 'chartBuilder') return INSERT_CHART_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertChart(variant.value) }));
    if (commandId === 'sparkline') return INSERT_SPARKLINE_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertSparkline(variant.value) }));
    if (commandId === 'shapesLines') return INSERT_SHAPE_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertShape(variant.value) }));
    if (commandId === 'formControls') return INSERT_FORM_CONTROL_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertFormControl(variant.value) }));
    if (commandId === 'dataChart') return INSERT_DATA_CHART_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertDataChart(variant.value) }));
    if (commandId === 'barcode') return INSERT_BARCODE_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertBarcode(variant.value) }));
    return [];
  };

  const renderSurface = (surface: RibbonSurfaceDefinition, mode: 'wide' | 'menu'): React.ReactNode => {
    if (!surface.commandId) return null;
    const variants = galleryItems(surface.commandId);
    if (variants.length > 0) {
      const title = insertText(locale, surface.commandId === 'chartBuilder' ? 'chart' : surface.commandId === 'sparkline' ? 'sparkline' : surface.commandId === 'shapesLines' ? 'shape' : surface.commandId === 'formControls' ? 'formControl' : surface.commandId === 'dataChart' ? 'dataChart' : 'barcode');
      if (mode === 'menu') return <React.Fragment key={surface.id}>{variants}</React.Fragment>;
      return <DropdownMenu key={surface.id} align="left" trigger={<RibbonLarge disabled={disabled} icon={surface.commandId === 'chartBuilder' ? 'chart-column' : surface.commandId === 'sparkline' ? 'sparkline' : surface.commandId === 'shapesLines' ? 'shape-square' : surface.commandId === 'formControls' ? 'form-control' : surface.commandId === 'dataChart' ? 'data-chart' : 'barcode'} title={title}>{title}</RibbonLarge>}><Stack gap="none" className="min-w-[14rem] p-1">{variants}</Stack></DropdownMenu>;
    }
    return <React.Fragment key={surface.id}>{renderCommand(surface.commandId, mode === 'menu' ? { className: 'w-full justify-start' } : { tile: surface.appearance === 'large' || surface.appearance === 'gallery' })}</React.Fragment>;
  };

  const renderMenuGroup = (group: RibbonGroupId) => {
    const definition = getRibbonGroupDefinition(group);
    const label = translateRibbonText(locale, definition.labelKey);
    return (
      <DropdownMenu key={group} align="left" trigger={<Button aria-label={label} title={label} icon="more-horizontal" size="sm" variant="ghost" className="h-[68px] min-w-0 flex-1 flex-col gap-1 rounded-none px-1 text-[10px] leading-3">{label}</Button>}>
        <Stack gap="none" className="min-w-[14rem] p-1">{allSurfaces(group).flatMap((surface) => renderSurface(surface, 'menu'))}</Stack>
      </DropdownMenu>
    );
  };

  if (breakpoint !== 'wide') {
    return <Inline gap="none" className="h-[102px] w-full min-w-0 flex-nowrap items-start overflow-visible" data-testid="insert-ribbon-groups" data-ribbon-breakpoint={breakpoint}>{INSERT_GROUPS.map((group, index) => <React.Fragment key={group}>{index > 0 ? <Divider orientation="vertical" className="h-[96px]" /> : null}{renderMenuGroup(group)}</React.Fragment>)}</Inline>;
  }

  return <Inline gap="none" className="h-[102px] w-full min-w-0 flex-nowrap items-start overflow-visible" data-testid="insert-ribbon-groups" data-ribbon-breakpoint={breakpoint}>
    {INSERT_GROUPS.map((group, index) => {
      const surfaces = surfacesByGroup.get(group) ?? [];
      return <React.Fragment key={group}>
        {index > 0 ? <Divider orientation="vertical" className="h-[96px]" /> : null}
        <Stack gap="none" className="h-[102px] min-w-0 flex-1 justify-between overflow-visible px-1">
          <Inline gap="none" className="min-h-0 flex-1 items-start justify-center overflow-visible pt-2">{surfaces.map((surface) => renderSurface(surface, 'wide'))}</Inline>
          <Text size="xs" tone="subtle" className="h-4 shrink-0 text-center text-[10px] font-medium text-[#413c40]">{translateRibbonText(locale, getRibbonGroupDefinition(group).labelKey)}</Text>
        </Stack>
      </React.Fragment>;
    })}
  </Inline>;
}
