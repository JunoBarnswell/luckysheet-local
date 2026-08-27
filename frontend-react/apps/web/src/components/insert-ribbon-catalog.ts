import { SHAPE_DRAWING_PRESETS, type BarcodeSymbology, type ChartDrawingPayload, type ChartSubtype, type DrawingConnectorType, type FormControlType, type ShapeDrawingCategory, type ShapeDrawingType, type SparklineModel } from '@react-sheets/core-model';
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

export interface InsertChartFamilyVariant {
  readonly id: string;
  readonly chartType: ChartGalleryType;
  readonly subtype: ChartSubtype;
}

export interface InsertChartFamilyDefinition {
  readonly id: string;
  readonly icon: IconName;
  readonly labelKey: InsertUiTextKey;
  readonly variants: readonly InsertChartFamilyVariant[];
}

export const INSERT_CHART_VARIANTS = [
  { id: 'chart.column', value: 'column', icon: 'chart-column', labelKey: 'chartColumn', ariaLabelKey: 'chartColumn', tooltipKey: 'chartColumn' },
  { id: 'chart.bar', value: 'bar', icon: 'chart-bar', labelKey: 'chartBar', ariaLabelKey: 'chartBar', tooltipKey: 'chartBar' },
  { id: 'chart.line', value: 'line', icon: 'chart-line', labelKey: 'chartLine', ariaLabelKey: 'chartLine', tooltipKey: 'chartLine' },
  { id: 'chart.area', value: 'area', icon: 'chart-area', labelKey: 'chartArea', ariaLabelKey: 'chartArea', tooltipKey: 'chartArea' },
  { id: 'chart.pie', value: 'pie', icon: 'chart-pie', labelKey: 'chartPie', ariaLabelKey: 'chartPie', tooltipKey: 'chartPie' },
  { id: 'chart.doughnut', value: 'doughnut', icon: 'chart-pie', labelKey: 'chartDoughnut', ariaLabelKey: 'chartDoughnut', tooltipKey: 'chartDoughnut' },
  { id: 'chart.scatter', value: 'scatter', icon: 'chart-scatter', labelKey: 'chartScatter', ariaLabelKey: 'chartScatter', tooltipKey: 'chartScatter' },
  { id: 'chart.hierarchy', value: 'treemap', icon: 'chart-column', labelKey: 'chartHierarchy', ariaLabelKey: 'chartHierarchy', tooltipKey: 'chartHierarchy' },
  { id: 'chart.statistical', value: 'histogram', icon: 'chart-column', labelKey: 'chartStatistical', ariaLabelKey: 'chartStatistical', tooltipKey: 'chartStatistical' },
  { id: 'chart.waterfall', value: 'waterfall', icon: 'chart-column', labelKey: 'chartWaterfall', ariaLabelKey: 'chartWaterfall', tooltipKey: 'chartWaterfall' },
  { id: 'chart.map', value: 'map', icon: 'chart-column', labelKey: 'chartMap', ariaLabelKey: 'chartMap', tooltipKey: 'chartMap' },
  { id: 'chart.combo', value: 'combo', icon: 'chart-column', labelKey: 'chartCombo', ariaLabelKey: 'chartCombo', tooltipKey: 'chartCombo' },
] as const satisfies readonly InsertVariantDefinition<ChartGalleryType>[];

export const INSERT_CHART_FAMILIES: readonly InsertChartFamilyDefinition[] = [
  { id: 'chart-family.column-bar', icon: 'chart-column', labelKey: 'chartColumn', variants: [
    { id: 'chart.column.clustered', chartType: 'column', subtype: 'clustered' }, { id: 'chart.column.stacked', chartType: 'column', subtype: 'stacked' }, { id: 'chart.column.percent', chartType: 'column', subtype: 'percent-stacked' },
    { id: 'chart.bar.clustered', chartType: 'bar', subtype: 'clustered' }, { id: 'chart.bar.stacked', chartType: 'bar', subtype: 'stacked' }, { id: 'chart.bar.percent', chartType: 'bar', subtype: 'percent-stacked' },
  ] },
  { id: 'chart-family.line-area', icon: 'chart-line', labelKey: 'chartLine', variants: [
    { id: 'chart.line.line', chartType: 'line', subtype: 'line' }, { id: 'chart.line.markers', chartType: 'line', subtype: 'line-markers' }, { id: 'chart.line.stacked', chartType: 'line', subtype: 'stacked' },
    { id: 'chart.area.area', chartType: 'area', subtype: 'area' }, { id: 'chart.area.stacked', chartType: 'area', subtype: 'stacked' }, { id: 'chart.area.percent', chartType: 'area', subtype: 'percent-stacked' },
  ] },
  { id: 'chart-family.pie', icon: 'chart-pie', labelKey: 'chartPie', variants: [
    { id: 'chart.pie.pie', chartType: 'pie', subtype: 'pie' }, { id: 'chart.pie.3d', chartType: 'pie', subtype: 'three-dimensional' }, { id: 'chart.pie.pie-of-pie', chartType: 'pie', subtype: 'pie-of-pie' }, { id: 'chart.pie.bar-of-pie', chartType: 'pie', subtype: 'bar-of-pie' }, { id: 'chart.doughnut', chartType: 'doughnut', subtype: 'doughnut' },
  ] },
  { id: 'chart-family.hierarchy', icon: 'chart-column', labelKey: 'chartHierarchy', variants: [
    { id: 'chart.treemap', chartType: 'treemap', subtype: 'treemap' }, { id: 'chart.sunburst', chartType: 'sunburst', subtype: 'sunburst' },
  ] },
  { id: 'chart-family.statistical', icon: 'chart-column', labelKey: 'chartStatistical', variants: [
    { id: 'chart.histogram', chartType: 'histogram', subtype: 'histogram' }, { id: 'chart.pareto', chartType: 'pareto', subtype: 'pareto' }, { id: 'chart.box-whisker', chartType: 'box-whisker', subtype: 'box-whisker' },
  ] },
  { id: 'chart-family.scatter-bubble', icon: 'chart-scatter', labelKey: 'chartScatter', variants: [
    { id: 'chart.scatter.markers', chartType: 'scatter', subtype: 'scatter-markers' }, { id: 'chart.scatter.smooth', chartType: 'scatter', subtype: 'scatter-smooth-lines' }, { id: 'chart.bubble', chartType: 'bubble', subtype: 'bubble' },
  ] },
  { id: 'chart-family.waterfall-more', icon: 'chart-column', labelKey: 'chartWaterfall', variants: [
    { id: 'chart.waterfall', chartType: 'waterfall', subtype: 'waterfall' }, { id: 'chart.funnel', chartType: 'funnel', subtype: 'funnel' }, { id: 'chart.stock', chartType: 'stock', subtype: 'stock-high-low-close' }, { id: 'chart.surface', chartType: 'surface', subtype: 'surface-three-dimensional' }, { id: 'chart.radar', chartType: 'radar', subtype: 'radar' },
  ] },
  { id: 'chart-family.combo', icon: 'chart-column', labelKey: 'chartCombo', variants: [{ id: 'chart.combo', chartType: 'combo', subtype: 'custom-combo' }] },
  { id: 'chart-family.map', icon: 'chart-column', labelKey: 'chartMap', variants: [{ id: 'chart.map', chartType: 'map', subtype: 'filled-map' }] },
] as const;

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

const SHAPE_META: Record<ShapeDrawingType, { icon: IconName; labelKey: InsertUiTextKey }> = {
  rectangle: { icon: 'shape-square', labelKey: 'shapeRectangle' },
  'rounded-rectangle': { icon: 'shape-square', labelKey: 'shapeRoundedRectangle' },
  ellipse: { icon: 'shape-circle', labelKey: 'shapeEllipse' },
  line: { icon: 'arrow-right', labelKey: 'shapeLine' },
  arrow: { icon: 'arrow-right', labelKey: 'shapeArrow' },
  callout: { icon: 'comment', labelKey: 'shapeCallout' },
  star: { icon: 'star', labelKey: 'shapeStar' },
};

const shapeVariants = (category: ShapeDrawingCategory): readonly InsertVariantDefinition<ShapeDrawingType>[] => SHAPE_DRAWING_PRESETS
  .filter((preset) => preset.category === category)
  .map((preset) => ({ id: `shape.${preset.type}`, value: preset.type, icon: SHAPE_META[preset.type].icon, labelKey: SHAPE_META[preset.type].labelKey, ariaLabelKey: SHAPE_META[preset.type].labelKey, tooltipKey: SHAPE_META[preset.type].labelKey }));

/** The gallery is categorized from the exact supported model/renderer set. */
export const INSERT_SHAPE_GALLERY: readonly ShapeGalleryCategory[] = [
  { id: 'basic-shapes', labelKey: 'shapeCategoryBasic', variants: shapeVariants('basic-shapes') },
  { id: 'lines', labelKey: 'shapeCategoryLines', variants: shapeVariants('lines') },
  { id: 'callouts-and-stars', labelKey: 'shapeCategoryCallouts', variants: shapeVariants('callouts-and-stars') },
] as const;

export const INSERT_CONNECTOR_VARIANTS = [
  { id: 'connector.straight', value: 'straight', icon: 'arrow-right', labelKey: 'connectorStraight', ariaLabelKey: 'connectorStraight', tooltipKey: 'connectorStraight' },
  { id: 'connector.elbow', value: 'elbow', icon: 'arrow-right', labelKey: 'connectorElbow', ariaLabelKey: 'connectorElbow', tooltipKey: 'connectorElbow' },
  { id: 'connector.curved', value: 'curved', icon: 'arrow-right', labelKey: 'connectorCurved', ariaLabelKey: 'connectorCurved', tooltipKey: 'connectorCurved' },
] as const satisfies readonly InsertVariantDefinition<DrawingConnectorType>[];


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
