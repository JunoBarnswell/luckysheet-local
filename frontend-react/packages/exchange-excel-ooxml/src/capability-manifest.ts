import { strFromU8 } from 'fflate';
import { child, children, descendants, localName, parseXml } from './xml';
import type { CompatibilityFeatureDetection } from './compatibility-report';
import type { OpcPackageGraph } from './types';

export type XlsxCapabilityState = 'full' | 'partial' | 'none';

export interface XlsxCapabilityDeclaration {
  feature: string;
  detect: XlsxCapabilityState;
  read: XlsxCapabilityState;
  write: XlsxCapabilityState;
  edit: XlsxCapabilityState;
  preserve: XlsxCapabilityState;
  approximation?: string;
}

/** Machine-readable source for import/export reporting and strict-mode gating. */
export const XLSX_CAPABILITY_MANIFEST = {
  cells: capability('cells', 'full', 'full', 'full', 'full', 'full'),
  formulas: capability('formulas', 'full', 'partial', 'full', 'partial', 'full'),
  styles: capability('styles', 'full', 'full', 'full', 'full', 'full'),
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
  charts: capability('charts', 'full', 'none', 'none', 'none', 'full'),
  images: capability('images', 'full', 'partial', 'partial', 'partial', 'full'),
  pivot: capability('pivot', 'full', 'partial', 'partial', 'partial', 'full'),
  slicer: capability('slicer', 'full', 'partial', 'partial', 'partial', 'full'),
  timeline: capability('timeline', 'full', 'partial', 'partial', 'partial', 'full'),
  vba: capability('vba', 'full', 'none', 'none', 'none', 'full'),
  'external-connection': capability('external-connection', 'full', 'none', 'none', 'none', 'full'),
  'unknown-extension': capability('unknown-extension', 'full', 'none', 'none', 'none', 'full'),
  'extended-validation': capability('extended-validation', 'full', 'none', 'none', 'none', 'full'),
  'extended-conditional-format': capability('extended-conditional-format', 'full', 'none', 'none', 'none', 'full'),
  'unknown-worksheet-node': capability('unknown-worksheet-node', 'full', 'none', 'none', 'none', 'none'),
} as const satisfies Record<string, XlsxCapabilityDeclaration>;

function capability(
  feature: string,
  detect: XlsxCapabilityState,
  read: XlsxCapabilityState,
  write: XlsxCapabilityState,
  edit: XlsxCapabilityState,
  preserve: XlsxCapabilityState,
  approximation?: string,
): XlsxCapabilityDeclaration {
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

export function capabilityFor(feature: string): XlsxCapabilityDeclaration {
  return XLSX_CAPABILITY_MANIFEST[feature as keyof typeof XLSX_CAPABILITY_MANIFEST]
    ?? capability(feature, 'partial', 'none', 'none', 'none', 'none');
}

function deduplicateDetections(values: CompatibilityFeatureDetection[]): CompatibilityFeatureDetection[] {
  const map = new Map<string, CompatibilityFeatureDetection>();
  for (const value of values) map.set(`${value.feature}\0${value.location ?? ''}`, value);
  return [...map.values()];
}
