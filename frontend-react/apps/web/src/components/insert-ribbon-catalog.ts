import type { ChartDrawingPayload, FormControlType, ShapeDrawingPayload, SparklineModel } from '@react-sheets/core-model';
import type { IconName } from '@react-sheets/ui-system';
import type { InsertUiTextKey } from '../i18n';

export interface InsertVariantDefinition<T extends string> {
  readonly id: string;
  readonly value: T;
  readonly icon: IconName;
  readonly labelKey: InsertUiTextKey;
  readonly ariaLabelKey: InsertUiTextKey;
  readonly tooltipKey: InsertUiTextKey;
}

type ChartGalleryType = ChartDrawingPayload['chartType'];

export const INSERT_CHART_VARIANTS = [
  { id: 'chart.column', value: 'column', icon: 'chart-column', labelKey: 'chartColumn', ariaLabelKey: 'chartColumn', tooltipKey: 'chartColumn' },
  { id: 'chart.bar', value: 'bar', icon: 'chart-bar', labelKey: 'chartBar', ariaLabelKey: 'chartBar', tooltipKey: 'chartBar' },
  { id: 'chart.line', value: 'line', icon: 'chart-line', labelKey: 'chartLine', ariaLabelKey: 'chartLine', tooltipKey: 'chartLine' },
  { id: 'chart.area', value: 'area', icon: 'chart-area', labelKey: 'chartArea', ariaLabelKey: 'chartArea', tooltipKey: 'chartArea' },
  { id: 'chart.pie', value: 'pie', icon: 'chart-pie', labelKey: 'chartPie', ariaLabelKey: 'chartPie', tooltipKey: 'chartPie' },
  { id: 'chart.doughnut', value: 'doughnut', icon: 'chart-pie', labelKey: 'chartDoughnut', ariaLabelKey: 'chartDoughnut', tooltipKey: 'chartDoughnut' },
  { id: 'chart.scatter', value: 'scatter', icon: 'chart-scatter', labelKey: 'chartScatter', ariaLabelKey: 'chartScatter', tooltipKey: 'chartScatter' },
  { id: 'chart.combo', value: 'combo', icon: 'data-chart', labelKey: 'chartCombo', ariaLabelKey: 'chartCombo', tooltipKey: 'chartCombo' },
] as const satisfies readonly InsertVariantDefinition<ChartGalleryType>[];

export const INSERT_SPARKLINE_VARIANTS = [
  { id: 'sparkline.line', value: 'line', icon: 'chart-line', labelKey: 'sparklineLine', ariaLabelKey: 'sparklineLine', tooltipKey: 'sparklineLine' },
  { id: 'sparkline.column', value: 'column', icon: 'chart-column', labelKey: 'sparklineColumn', ariaLabelKey: 'sparklineColumn', tooltipKey: 'sparklineColumn' },
  { id: 'sparkline.win-loss', value: 'win-loss', icon: 'chart-bar', labelKey: 'sparklineWinLoss', ariaLabelKey: 'sparklineWinLoss', tooltipKey: 'sparklineWinLoss' },
] as const satisfies readonly InsertVariantDefinition<SparklineModel['type']>[];

export const INSERT_SHAPE_VARIANTS = [
  { id: 'shape.rectangle', value: 'rectangle', icon: 'shape-square', labelKey: 'shapeRectangle', ariaLabelKey: 'shapeRectangle', tooltipKey: 'shapeRectangle' },
  { id: 'shape.rounded-rectangle', value: 'rounded-rectangle', icon: 'shape-square', labelKey: 'shapeRoundedRectangle', ariaLabelKey: 'shapeRoundedRectangle', tooltipKey: 'shapeRoundedRectangle' },
  { id: 'shape.ellipse', value: 'ellipse', icon: 'shape-circle', labelKey: 'shapeEllipse', ariaLabelKey: 'shapeEllipse', tooltipKey: 'shapeEllipse' },
  { id: 'shape.line', value: 'line', icon: 'arrow-right', labelKey: 'shapeLine', ariaLabelKey: 'shapeLine', tooltipKey: 'shapeLine' },
  { id: 'shape.arrow', value: 'arrow', icon: 'arrow-right', labelKey: 'shapeArrow', ariaLabelKey: 'shapeArrow', tooltipKey: 'shapeArrow' },
  { id: 'shape.callout', value: 'callout', icon: 'comment', labelKey: 'shapeCallout', ariaLabelKey: 'shapeCallout', tooltipKey: 'shapeCallout' },
  { id: 'shape.star', value: 'star', icon: 'star', labelKey: 'shapeStar', ariaLabelKey: 'shapeStar', tooltipKey: 'shapeStar' },
] as const satisfies readonly InsertVariantDefinition<ShapeDrawingPayload['type']>[];

export const INSERT_FORM_CONTROL_VARIANTS = [
  { id: 'form-control.button', value: 'button', icon: 'form-control', labelKey: 'formButton', ariaLabelKey: 'formButton', tooltipKey: 'formButton' },
  { id: 'form-control.spin-button', value: 'spin-button', icon: 'form-control', labelKey: 'formSpinButton', ariaLabelKey: 'formSpinButton', tooltipKey: 'formSpinButton' },
  { id: 'form-control.list-box', value: 'list-box', icon: 'form-control', labelKey: 'formListBox', ariaLabelKey: 'formListBox', tooltipKey: 'formListBox' },
  { id: 'form-control.combo-box', value: 'combo-box', icon: 'form-control', labelKey: 'formComboBox', ariaLabelKey: 'formComboBox', tooltipKey: 'formComboBox' },
  { id: 'form-control.checkbox', value: 'checkbox', icon: 'checkbox', labelKey: 'formCheckbox', ariaLabelKey: 'formCheckbox', tooltipKey: 'formCheckbox' },
  { id: 'form-control.option-button', value: 'option-button', icon: 'form-control', labelKey: 'formOptionButton', ariaLabelKey: 'formOptionButton', tooltipKey: 'formOptionButton' },
  { id: 'form-control.group-box', value: 'group-box', icon: 'form-control', labelKey: 'formGroupBox', ariaLabelKey: 'formGroupBox', tooltipKey: 'formGroupBox' },
  { id: 'form-control.label', value: 'label', icon: 'form-control', labelKey: 'formLabel', ariaLabelKey: 'formLabel', tooltipKey: 'formLabel' },
  { id: 'form-control.scrollbar', value: 'scrollbar', icon: 'form-control', labelKey: 'formScrollbar', ariaLabelKey: 'formScrollbar', tooltipKey: 'formScrollbar' },
] as const satisfies readonly InsertVariantDefinition<FormControlType>[];
