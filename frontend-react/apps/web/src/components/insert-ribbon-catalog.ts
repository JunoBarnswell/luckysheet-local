import type { BarcodeSymbology, ChartDrawingPayload, DataChartPlotType, FormControlType, ShapeDrawingCategory, ShapeDrawingType, SparklineModel } from '@react-sheets/core-model';
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

export const INSERT_DATA_CHART_VARIANTS = [
  { id: 'data-chart.column', value: 'column', icon: 'chart-column', labelKey: 'dataChartColumn', ariaLabelKey: 'dataChartColumn', tooltipKey: 'dataChartColumn' },
  { id: 'data-chart.bar', value: 'bar', icon: 'chart-bar', labelKey: 'dataChartBar', ariaLabelKey: 'dataChartBar', tooltipKey: 'dataChartBar' },
  { id: 'data-chart.line', value: 'line', icon: 'chart-line', labelKey: 'dataChartLine', ariaLabelKey: 'dataChartLine', tooltipKey: 'dataChartLine' },
  { id: 'data-chart.area', value: 'area', icon: 'chart-area', labelKey: 'dataChartArea', ariaLabelKey: 'dataChartArea', tooltipKey: 'dataChartArea' },
  { id: 'data-chart.pie', value: 'pie', icon: 'chart-pie', labelKey: 'dataChartPie', ariaLabelKey: 'dataChartPie', tooltipKey: 'dataChartPie' },
  { id: 'data-chart.doughnut', value: 'doughnut', icon: 'chart-pie', labelKey: 'dataChartDoughnut', ariaLabelKey: 'dataChartDoughnut', tooltipKey: 'dataChartDoughnut' },
  { id: 'data-chart.scatter', value: 'scatter', icon: 'chart-scatter', labelKey: 'dataChartScatter', ariaLabelKey: 'dataChartScatter', tooltipKey: 'dataChartScatter' },
] as const satisfies readonly InsertVariantDefinition<DataChartPlotType>[];

export const INSERT_BARCODE_VARIANTS = [
  { id: 'barcode.qr', value: 'qr', icon: 'barcode', labelKey: 'barcodeQr', ariaLabelKey: 'barcodeQr', tooltipKey: 'barcodeQr' },
  { id: 'barcode.code128', value: 'code128', icon: 'barcode', labelKey: 'barcodeCode128', ariaLabelKey: 'barcodeCode128', tooltipKey: 'barcodeCode128' },
  { id: 'barcode.code39', value: 'code39', icon: 'barcode', labelKey: 'barcodeCode39', ariaLabelKey: 'barcodeCode39', tooltipKey: 'barcodeCode39' },
  { id: 'barcode.code93', value: 'code93', icon: 'barcode', labelKey: 'barcodeCode93', ariaLabelKey: 'barcodeCode93', tooltipKey: 'barcodeCode93' },
  { id: 'barcode.code49', value: 'code49', icon: 'barcode', labelKey: 'barcodeCode49', ariaLabelKey: 'barcodeCode49', tooltipKey: 'barcodeCode49' },
  { id: 'barcode.codabar', value: 'codabar', icon: 'barcode', labelKey: 'barcodeCodabar', ariaLabelKey: 'barcodeCodabar', tooltipKey: 'barcodeCodabar' },
  { id: 'barcode.ean13', value: 'ean13', icon: 'barcode', labelKey: 'barcodeEan13', ariaLabelKey: 'barcodeEan13', tooltipKey: 'barcodeEan13' },
  { id: 'barcode.ean8', value: 'ean8', icon: 'barcode', labelKey: 'barcodeEan8', ariaLabelKey: 'barcodeEan8', tooltipKey: 'barcodeEan8' },
  { id: 'barcode.upca', value: 'upca', icon: 'barcode', labelKey: 'barcodeUpca', ariaLabelKey: 'barcodeUpca', tooltipKey: 'barcodeUpca' },
  { id: 'barcode.gs1-128', value: 'gs1-128', icon: 'barcode', labelKey: 'barcodeGs1128', ariaLabelKey: 'barcodeGs1128', tooltipKey: 'barcodeGs1128' },
  { id: 'barcode.pdf417', value: 'pdf417', icon: 'barcode', labelKey: 'barcodePdf417', ariaLabelKey: 'barcodePdf417', tooltipKey: 'barcodePdf417' },
  { id: 'barcode.data-matrix', value: 'data-matrix', icon: 'barcode', labelKey: 'barcodeDataMatrix', ariaLabelKey: 'barcodeDataMatrix', tooltipKey: 'barcodeDataMatrix' },
] as const satisfies readonly InsertVariantDefinition<BarcodeSymbology>[];

export const INSERT_SPARKLINE_VARIANTS = [
  { id: 'sparkline.line', value: 'line', icon: 'chart-line', labelKey: 'sparklineLine', ariaLabelKey: 'sparklineLine', tooltipKey: 'sparklineLine' },
  { id: 'sparkline.column', value: 'column', icon: 'chart-column', labelKey: 'sparklineColumn', ariaLabelKey: 'sparklineColumn', tooltipKey: 'sparklineColumn' },
  { id: 'sparkline.win-loss', value: 'win-loss', icon: 'chart-bar', labelKey: 'sparklineWinLoss', ariaLabelKey: 'sparklineWinLoss', tooltipKey: 'sparklineWinLoss' },
] as const satisfies readonly InsertVariantDefinition<SparklineModel['type']>[];

export interface ShapeGalleryCategory {
  readonly id: ShapeDrawingCategory;
  readonly labelKey: InsertUiTextKey;
  readonly variants: readonly InsertVariantDefinition<ShapeDrawingType>[];
}

const SHAPE_VARIANTS_BASIC = [
  { id: 'shape.rectangle', value: 'rectangle', icon: 'shape-square', labelKey: 'shapeRectangle', ariaLabelKey: 'shapeRectangle', tooltipKey: 'shapeRectangle' },
  { id: 'shape.rounded-rectangle', value: 'rounded-rectangle', icon: 'shape-square', labelKey: 'shapeRoundedRectangle', ariaLabelKey: 'shapeRoundedRectangle', tooltipKey: 'shapeRoundedRectangle' },
  { id: 'shape.ellipse', value: 'ellipse', icon: 'shape-circle', labelKey: 'shapeEllipse', ariaLabelKey: 'shapeEllipse', tooltipKey: 'shapeEllipse' },
] as const satisfies readonly InsertVariantDefinition<ShapeDrawingType>[];

const SHAPE_VARIANTS_LINES = [
  { id: 'shape.line', value: 'line', icon: 'arrow-right', labelKey: 'shapeLine', ariaLabelKey: 'shapeLine', tooltipKey: 'shapeLine' },
  { id: 'shape.arrow', value: 'arrow', icon: 'arrow-right', labelKey: 'shapeArrow', ariaLabelKey: 'shapeArrow', tooltipKey: 'shapeArrow' },
] as const satisfies readonly InsertVariantDefinition<ShapeDrawingType>[];

const SHAPE_VARIANTS_CALLOUTS = [
  { id: 'shape.callout', value: 'callout', icon: 'comment', labelKey: 'shapeCallout', ariaLabelKey: 'shapeCallout', tooltipKey: 'shapeCallout' },
  { id: 'shape.star', value: 'star', icon: 'star', labelKey: 'shapeStar', ariaLabelKey: 'shapeStar', tooltipKey: 'shapeStar' },
] as const satisfies readonly InsertVariantDefinition<ShapeDrawingType>[];

/** The gallery is categorized from the exact supported model/renderer set. */
export const INSERT_SHAPE_GALLERY: readonly ShapeGalleryCategory[] = [
  { id: 'basic-shapes', labelKey: 'shapeCategoryBasic', variants: SHAPE_VARIANTS_BASIC },
  { id: 'lines', labelKey: 'shapeCategoryLines', variants: SHAPE_VARIANTS_LINES },
  { id: 'callouts-and-stars', labelKey: 'shapeCategoryCallouts', variants: SHAPE_VARIANTS_CALLOUTS },
] as const;


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
