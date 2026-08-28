import { strFromU8 } from 'fflate';
import { child, children, descendants, localName, parseXml } from './xml';
import type { CompatibilityFeatureDetection } from './compatibility-report';
import type { OpcPackageGraph } from './types';

export type NativeCapabilityState = 'full' | 'partial' | 'none';

export interface NativeCapabilityDeclaration {
  feature: string;
  detect: NativeCapabilityState;
  read: NativeCapabilityState;
  write: NativeCapabilityState;
  edit: NativeCapabilityState;
  preserve: NativeCapabilityState;
  approximation?: string;
}

/** Machine-readable source for import/export reporting and strict-mode gating. */
export const NATIVE_DOCUMENT_CAPABILITY_MANIFEST = {
  cells: capability('cells', 'full', 'full', 'full', 'full', 'full'),
  formulas: capability('formulas', 'full', 'partial', 'full', 'partial', 'full'),
  styles: capability('styles', 'full', 'full', 'full', 'full', 'full'),
  'cell-style-template': capability('cell-style-template', 'full', 'full', 'partial', 'partial', 'partial', 'OOXML named cell styles retain template names and styles; editor metadata remains workbook-native.'),
  'rich-text': capability('rich-text', 'full', 'full', 'full', 'partial', 'full', 'Unsupported run properties remain source-package metadata.'),
  merges: capability('merges', 'full', 'full', 'full', 'full', 'full'),
  freeze: capability('freeze', 'full', 'full', 'full', 'full', 'full'),
  split: capability('split', 'full', 'full', 'full', 'partial', 'full'),
  hyperlinks: capability('hyperlinks', 'full', 'full', 'full', 'full', 'full'),
  tables: capability('tables', 'full', 'full', 'full', 'full', 'full'),
  'conditional-format': capability('conditional-format', 'full', 'partial', 'partial', 'partial', 'full'),
  validation: capability('validation', 'full', 'partial', 'partial', 'partial', 'full'),
  filters: capability('filters', 'full', 'partial', 'partial', 'partial', 'full'),
  outline: capability('outline', 'full', 'full', 'full', 'full', 'full'),
  protection: capability('protection', 'full', 'partial', 'partial', 'partial', 'full'),
  'print-setup': capability('print-setup', 'full', 'partial', 'partial', 'partial', 'full'),
  charts: capability('charts', 'full', 'none', 'none', 'none', 'full', 'Generic chart parts are preserved-only; native-authored PivotCharts are reported as pivot-chart.'),
  'pivot-chart': capability('pivot-chart', 'full', 'full', 'partial', 'full', 'full', 'Native PivotChart write supports column, bar, line and area charts linked to the emitted PivotTable/PivotCache.'),
  images: capability('images', 'full', 'partial', 'partial', 'partial', 'full'),
  'table-sheet': capability('table-sheet', 'full', 'full', 'partial', 'full', 'full', 'Exported as a materialized worksheet and Excel table; canonical metadata is retained in custom XML.'),
  'gantt-sheet': capability('gantt-sheet', 'full', 'full', 'partial', 'full', 'full', 'Task data is materialized; the canonical Gantt definition is retained in custom XML.'),
  'report-sheet': capability('report-sheet', 'full', 'full', 'partial', 'full', 'full', 'The generated report grid is exported and the canonical template binding is retained in custom XML.'),
  barcode: capability('barcode', 'full', 'full', 'partial', 'full', 'full', 'Barcode source and symbology are retained and projected for Excel.'),
  camera: capability('camera', 'full', 'full', 'partial', 'full', 'full', 'The live source range is retained in custom XML.'),
  'form-control': capability('form-control', 'full', 'full', 'partial', 'full', 'full', 'Supported controls are projected; canonical cell-link metadata is retained.'),
  icons: capability('icons', 'full', 'full', 'full', 'full', 'full', 'Local Fluent SVG paths are retained in React Sheets custom XML and rendered without a host.'),
  models3d: capability('models3d', 'full', 'full', 'full', 'full', 'full', 'Local OBJ geometry is normalized and retained in custom XML; the renderer is deterministic wireframe.'),
  smartart: capability('smartart', 'full', 'full', 'full', 'full', 'full', 'SmartArt nodes, edges and layout are local canonical data.'),
  wordart: capability('wordart', 'full', 'full', 'full', 'full', 'full', 'WordArt text styling is rendered and retained locally.'),
  'signature-line': capability('signature-line', 'full', 'full', 'full', 'full', 'full', 'Signature metadata and local signed state are workbook-owned; certificate providers are not required.'),
  'embedded-object': capability('embedded-object', 'full', 'full', 'full', 'full', 'full', 'File bytes are content-addressed in the local asset store; activation is intentionally not part of the object contract.'),
  equation: capability('equation', 'full', 'full', 'full', 'full', 'full', 'The supported linear/OMML token subset is rendered and retained locally.'),
  screenshot: capability('screenshot', 'full', 'full', 'full', 'full', 'full', 'Screenshot captures a workbook region locally and never requests OS screen permission.'),
  xmlss: capability('xmlss', 'full', 'full', 'full', 'full', 'full', 'SpreadsheetML 2003 is parsed and written directly without OOXML conversion.'),
  text: capability('text', 'full', 'full', 'full', 'full', 'full', 'Text dialect encoding, BOM, delimiter, quote and row terminators are owned by the text codec.'),
  ods: capability('ods', 'full', 'full', 'full', 'full', 'full', 'ODF package parts are parsed and written directly; unknown parts remain in the package graph.'),
  sjs: capability('sjs', 'full', 'full', 'full', 'full', 'full', 'SpreadJS SJS JSON parts are parsed and written directly.'),
  ssjson: capability('ssjson', 'full', 'full', 'full', 'full', 'full', 'SpreadJS SSJSON is validated as its own JSON document.'),
  sharedStrings: capability('sharedStrings', 'full', 'full', 'partial', 'partial', 'full', 'Shared string tables are read natively and retained; edited cells use direct native string records without rewriting untouched entries.'),
  xlsb: capability('xlsb', 'full', 'partial', 'partial', 'partial', 'full', 'BIFF12 cell records and package parts are read and rewritten natively; formula expressions and unsupported row structures remain preserved-only.'),
  biff: capability('biff', 'full', 'partial', 'partial', 'partial', 'full', 'BIFF/CFB workbook records and basic cell records are read and rewritten natively; formula expressions and unsupported record structures remain preserved-only.'),
  dbf: capability('dbf', 'full', 'none', 'none', 'none', 'full', 'Excel lists DBF as open-only; the local runtime refuses projection without a DBF reader.'),
  works: capability('works', 'full', 'none', 'none', 'none', 'full', 'Works spreadsheet files are detected but remain blocked without a native reader.'),
  web: capability('web', 'full', 'none', 'none', 'none', 'full', 'Office web documents are detected but are not workbook round-trip formats.'),
  presentation: capability('presentation', 'full', 'none', 'none', 'none', 'full', 'PDF/XPS are presentation exports, not workbook round-trip formats.'),
  pivot: capability('pivot', 'full', 'partial', 'partial', 'partial', 'full'),
  slicer: capability('slicer', 'full', 'partial', 'partial', 'partial', 'full'),
  timeline: capability('timeline', 'full', 'partial', 'partial', 'partial', 'full'),
  vba: capability('vba', 'full', 'none', 'none', 'none', 'full'),
  'external-connection': capability('external-connection', 'full', 'none', 'none', 'none', 'full'),
  'unknown-extension': capability('unknown-extension', 'full', 'none', 'none', 'none', 'full'),
  'extended-validation': capability('extended-validation', 'full', 'none', 'none', 'none', 'full'),
  'extended-conditional-format': capability('extended-conditional-format', 'full', 'none', 'none', 'none', 'full'),
  'unknown-worksheet-node': capability('unknown-worksheet-node', 'full', 'none', 'none', 'none', 'none'),
} as const satisfies Record<string, NativeCapabilityDeclaration>;

function capability(
  feature: string,
  detect: NativeCapabilityState,
  read: NativeCapabilityState,
  write: NativeCapabilityState,
  edit: NativeCapabilityState,
  preserve: NativeCapabilityState,
  approximation?: string,
): NativeCapabilityDeclaration {
  return { feature, detect, read, write, edit, preserve, ...(approximation ? { approximation } : {}) };
}

const WORKSHEET_NODES = new Map<string, string>([
  ['sheetData', 'cells'], ['cols', 'styles'], ['mergeCells', 'merges'], ['hyperlinks', 'hyperlinks'],
  ['conditionalFormatting', 'conditional-format'], ['dataValidations', 'validation'], ['autoFilter', 'filters'],
  ['sheetProtection', 'protection'], ['printOptions', 'print-setup'], ['pageMargins', 'print-setup'],
  ['pageSetup', 'print-setup'], ['headerFooter', 'print-setup'], ['rowBreaks', 'print-setup'], ['colBreaks', 'print-setup'],
  ['tableParts', 'tables'], ['drawing', 'images'], ['legacyDrawing', 'images'], ['pivotTableParts', 'pivot'],
]);

const STRUCTURAL_NODES = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'sheetCalcPr', 'phoneticPr', 'extLst']);

export function detectWorksheetCapabilities(files: Record<string, Uint8Array>, pkg: OpcPackageGraph): CompatibilityFeatureDetection[] {
  const detections: CompatibilityFeatureDetection[] = [];
  for (const part of new Set(Object.values(pkg.sheetPartById))) {
    const bytes = files[part];
    if (!bytes) continue;
    const root = descendants(parseXml(strFromU8(bytes)), 'worksheet')[0];
    if (!root) continue;
    const view = child(child(root, 'sheetViews'), 'sheetView');
    const pane = child(view, 'pane');
    if (pane) detections.push({ feature: pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit' ? 'freeze' : 'split', location: part });
    if (children(child(root, 'cols'), 'col').some((node) => Number(node.attrs.outlineLevel ?? 0) > 0)
      || children(child(root, 'sheetData'), 'row').some((node) => Number(node.attrs.outlineLevel ?? 0) > 0)) detections.push({ feature: 'outline', location: part });
    for (const node of root.children) {
      const name = localName(node.name);
      const feature = WORKSHEET_NODES.get(name);
      if (feature) detections.push({ feature, location: part });
      else if (!STRUCTURAL_NODES.has(name)) detections.push({ feature: 'unknown-worksheet-node', location: `${part}#${name}`, reason: `No validated reader/writer contract exists for worksheet node <${name}>` });
      if (name === 'extLst') {
        if (descendants(node, 'dataValidations').length || descendants(node, 'dataValidation').length) detections.push({ feature: 'extended-validation', location: `${part}#extLst`, reason: 'Extended data validation is preserved in the source package but is not editable' });
        if (descendants(node, 'conditionalFormatting').length) detections.push({ feature: 'extended-conditional-format', location: `${part}#extLst`, reason: 'Extended conditional formatting is preserved in the source package but is not editable' });
        for (const extension of children(node, 'ext')) detections.push({ feature: 'unknown-extension', location: `${part}#${extension.attrs.uri ?? 'ext'}`, reason: 'Worksheet extension is retained byte-for-byte from the source package' });
      }
    }
  }
  const sharedStringsRelation = (pkg.relationships[pkg.workbookPart] ?? []).find((relation) => relation.type.replace(/\/+$/, '').endsWith('/sharedStrings'));
  const sharedStringsPart = sharedStringsRelation ? resolvePart(pkg.workbookPart, sharedStringsRelation.target) : 'xl/sharedStrings.xml';
  const sharedStrings = files[sharedStringsPart];
  if (sharedStrings && /<(?:\w+:)?r(?:\s|>)/.test(strFromU8(sharedStrings))) detections.push({ feature: 'rich-text', location: sharedStringsPart });
  return deduplicateDetections(detections);
}

function resolvePart(source: string, target: string): string {
  const pieces = `${source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : ''}${target}`.replace(/\\/g, '/').split('/');
  const result: string[] = [];
  for (const piece of pieces) {
    if (!piece || piece === '.') continue;
    if (piece === '..') result.pop(); else result.push(piece);
  }
  return result.join('/');
}

export function capabilityFor(feature: string): NativeCapabilityDeclaration {
  return NATIVE_DOCUMENT_CAPABILITY_MANIFEST[feature as keyof typeof NATIVE_DOCUMENT_CAPABILITY_MANIFEST]
    ?? capability(feature, 'partial', 'none', 'none', 'none', 'none');
}

function deduplicateDetections(values: CompatibilityFeatureDetection[]): CompatibilityFeatureDetection[] {
  const map = new Map<string, CompatibilityFeatureDetection>();
  for (const value of values) map.set(`${value.feature}\0${value.location ?? ''}`, value);
  return [...map.values()];
}
