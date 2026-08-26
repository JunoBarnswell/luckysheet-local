import React from 'react';
import { Button, DropdownMenu, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { type RibbonCommandId, type RibbonSurfaceDefinition } from '@react-sheets/spreadsheet-app';
import type { BarcodeSymbology, ChartDrawingPayload, DataChartPlotType, DrawingConnectorType, FormControlType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { Locale } from '../i18n';
import { insertText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { RibbonLayoutRenderer } from './RibbonLayoutRenderer';
import { INSERT_BARCODE_VARIANTS, INSERT_CHART_VARIANTS, INSERT_CONNECTOR_VARIANTS, INSERT_DATA_CHART_VARIANTS, INSERT_FORM_CONTROL_VARIANTS, INSERT_SHAPE_GALLERY, INSERT_SPARKLINE_VARIANTS } from './insert-ribbon-catalog';

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
  onInsertConnector: (type: DrawingConnectorType) => void;
  onInsertFormControl: (type: FormControlType) => void;
}

function RibbonLarge({ children, icon, disabled, surfaceId, title }: { children: React.ReactNode; icon: React.ComponentProps<typeof Button>['icon']; disabled?: boolean; surfaceId: string; title: string }) {
  return <Button aria-label={title} data-ribbon-surface={surfaceId} title={title} disabled={disabled} icon={icon} size="sm" variant="ghost" className="!h-[68px] !min-h-0 !w-[68px] flex-col gap-1 rounded-none px-1 text-[11px] leading-4 [&>svg]:!h-6 [&>svg]:!w-6">{children}</Button>;
}

function variantButton({ id, icon, label, onSelect, surfaceId, disabled }: { id: string; icon: React.ComponentProps<typeof Button>['icon']; label: string; onSelect: () => void; surfaceId: string; disabled?: boolean }) {
  return <Button key={id} aria-label={label} data-ribbon-surface={surfaceId} data-ribbon-variant={id} title={label} icon={icon} disabled={disabled} size="sm" variant="ghost" className="w-full justify-start" onClick={onSelect}>{label}</Button>;
}

export function InsertRibbon({ locale, layout, disabled, renderCommand, onInsertChart, onInsertDataChart, onInsertBarcode, onInsertSparkline, onInsertShape, onInsertConnector, onInsertFormControl }: InsertRibbonProps) {
  const galleryItems = (commandId: RibbonCommandId, surfaceId: string): React.ReactNode[] => {
    if (commandId === 'chartBuilder') return INSERT_CHART_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertChart(variant.value), surfaceId }));
    if (commandId === 'sparkline') return INSERT_SPARKLINE_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertSparkline(variant.value), surfaceId }));
    if (commandId === 'shapesLines') return INSERT_SHAPE_GALLERY.flatMap((category) => [
      <Text key={`${category.id}.label`} size="xs" weight="semibold" className="px-2 pb-1 pt-2 text-slate-500">{insertText(locale, category.labelKey)}</Text>,
      ...category.variants.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertShape(variant.value), surfaceId })),
      <Text key="connectors.label" size="xs" weight="semibold" className="px-2 pb-1 pt-2 text-slate-500">{insertText(locale, 'connectorCategory')}</Text>,
      ...INSERT_CONNECTOR_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertConnector(variant.value), surfaceId })),
    ]);
    if (commandId === 'formControls') return INSERT_FORM_CONTROL_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertFormControl(variant.value), surfaceId }));
    if (commandId === 'dataChart') return INSERT_DATA_CHART_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertDataChart(variant.value), surfaceId }));
    if (commandId === 'barcode') return INSERT_BARCODE_VARIANTS.map((variant) => variantButton({ id: variant.id, icon: variant.icon, label: insertText(locale, variant.labelKey), disabled, onSelect: () => onInsertBarcode(variant.value), surfaceId }));
    return [];
  };

  const renderSurface = (surface: RibbonSurfaceDefinition, mode: 'wide' | 'menu'): React.ReactNode => {
    if (!surface.commandId) return null;
    const variants = galleryItems(surface.commandId, surface.id);
    if (variants.length > 0) {
      const title = insertText(locale, surface.commandId === 'chartBuilder' ? 'chart' : surface.commandId === 'sparkline' ? 'sparkline' : surface.commandId === 'shapesLines' ? 'shape' : surface.commandId === 'formControls' ? 'formControl' : surface.commandId === 'dataChart' ? 'dataChart' : 'barcode');
      if (mode === 'menu') return <React.Fragment key={surface.id}>{variants}</React.Fragment>;
      return <DropdownMenu key={surface.id} align="left" trigger={<RibbonLarge disabled={disabled} icon={surface.commandId === 'chartBuilder' ? 'chart-column' : surface.commandId === 'sparkline' ? 'sparkline' : surface.commandId === 'shapesLines' ? 'shape-square' : surface.commandId === 'formControls' ? 'form-control' : surface.commandId === 'dataChart' ? 'data-chart' : 'barcode'} surfaceId={surface.id} title={title}>{title}</RibbonLarge>}><Stack gap="none" className="min-w-[14rem] p-1">{variants}</Stack></DropdownMenu>;
    }
    return <React.Fragment key={surface.id}>{renderCommand(surface.commandId, mode === 'menu' ? { className: 'w-full justify-start', ribbonSurfaceId: surface.id } : { tile: surface.appearance === 'large' || surface.appearance === 'gallery', ribbonSurfaceId: surface.id })}</React.Fragment>;
  };

  return <RibbonLayoutRenderer tab="insert" locale={locale} layout={layout} renderCommand={renderCommand} renderSurface={(surface, context) => renderSurface(surface, context.mode)} />;
}
