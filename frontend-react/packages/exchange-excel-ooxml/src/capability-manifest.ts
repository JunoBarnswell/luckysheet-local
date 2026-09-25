import { strFromU8 } from 'fflate';
import { child, children, descendants, localName, parseXml, type XmlNode } from './xml';
import type { CompatibilityFeatureDetection } from './compatibility-report';
import type { NativePivotControlDefinition, OpcPackageGraph } from './types';

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
  theme: capability('theme', 'full', 'full', 'full', 'full', 'full'),
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
  charts: capability('charts', 'full', 'full', 'partial', 'full', 'full', 'Canonical chart families are editable; unsupported native identities remain preserved-native.'),
  'preserved-native-chart': capability('preserved-native-chart', 'full', 'none', 'none', 'none', 'full', 'The native chart part is retained byte-for-byte because this runtime does not own its complete semantic editor.'),
  'pivot-chart': capability('pivot-chart', 'full', 'full', 'partial', 'full', 'full', 'PivotChart uses the same native writer as worksheet Chart and remains linked to its PivotTable/PivotCache.'),
  sparklines: capability('sparklines', 'full', 'full', 'partial', 'full', 'full', 'Sparkline groups use canonical metadata plus native worksheet extensions for Excel round-trip.'),
  images: capability('images', 'full', 'partial', 'partial', 'partial', 'full'),
  'table-sheet': capability('table-sheet', 'full', 'full', 'partial', 'full', 'full', 'Exported as a materialized worksheet and Excel table; canonical metadata is retained in custom XML.'),
  'gantt-sheet': capability('gantt-sheet', 'full', 'full', 'partial', 'full', 'full', 'Task data is materialized; the canonical Gantt definition is retained in custom XML.'),
  'report-sheet': capability('report-sheet', 'full', 'full', 'partial', 'full', 'full', 'The generated report grid is exported and the canonical template binding is retained in custom XML.'),
  barcode: capability('barcode', 'full', 'full', 'partial', 'full', 'full', 'Barcode source and symbology are retained and projected for Excel.'),
  camera: capability('camera', 'full', 'full', 'partial', 'full', 'full', 'The live source range is retained in custom XML.'),
  'form-control': capability('form-control', 'full', 'full', 'partial', 'full', 'full', 'Legacy Form Controls are preserved/projected only where the canonical control contract is supported; ActiveX is never converted.'),
  forms: capability('forms', 'full', 'none', 'none', 'none', 'full', 'Microsoft Forms is an Office service and is not represented by a local Form Control.'),
  icons: capability('icons', 'full', 'full', 'full', 'full', 'full', 'Local Fluent SVG paths are retained in React Sheets custom XML and rendered without a host.'),
  models3d: capability('models3d', 'full', 'partial', 'none', 'none', 'full', 'Native 3D model activation and full view semantics require an Office host; unknown model parts are preserved.'),
  smartart: capability('smartart', 'full', 'partial', 'none', 'none', 'full', 'The local drawing projection is not an Excel SmartArt editor; native SmartArt remains preserve-only.'),
  wordart: capability('wordart', 'full', 'partial', 'none', 'none', 'full', 'The local text projection is not an Excel WordArt editor; native effects remain preserve-only.'),
  'signature-line': capability('signature-line', 'full', 'partial', 'none', 'none', 'full', 'Office digital signatures require certificate and host execution; the runtime never fabricates a signed state.'),
  'embedded-object': capability('embedded-object', 'full', 'partial', 'none', 'none', 'full', 'OLE activation is host-owned; opaque embedded parts are preserved without local activation.'),
  equation: capability('equation', 'full', 'partial', 'none', 'none', 'full', 'Native OMML editing requires a math object owner; unsupported equations remain preserved.'),
  screenshot: capability('screenshot', 'full', 'full', 'none', 'none', 'full', 'System Screenshot is host-owned; workbook range capture is classified as the Camera extension.'),
  xmlss: capability('xmlss', 'full', 'full', 'full', 'full', 'full', 'SpreadsheetML 2003 is parsed and written directly without OOXML conversion.'),
  text: capability('text', 'full', 'full', 'full', 'full', 'full', 'Text dialect encoding, BOM, delimiter, quote and row terminators are owned by the text codec.'),
  ods: capability('ods', 'full', 'full', 'full', 'full', 'full', 'ODF package parts are parsed and written directly; unknown parts remain in the package graph.'),
  sjs: capability('sjs', 'full', 'partial', 'partial', 'partial', 'partial', 'Value-only single-sheet JSON projection. Object, formula, format and layout export is rejected; choose XLSX explicitly.'),
  ssjson: capability('ssjson', 'full', 'partial', 'partial', 'partial', 'partial', 'Value-only single-sheet JSON projection. Object, formula, format and layout export is rejected; choose XLSX explicitly.'),
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
  ['tableParts', 'tables'], ['pivotTableParts', 'pivot'],
]);

const STRUCTURAL_NODES = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'sheetCalcPr', 'phoneticPr', 'extLst', 'drawing', 'legacyDrawing']);
const SPARKLINE_GROUPS_EXTENSION_URI = '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}';
const SLICER_LIST_EXTENSION_URI = '{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}';
const TIMELINE_REFS_EXTENSION_URI = '{7E03D99C-DC04-49D9-9315-930204A7B6E9}';
const SLICER_CACHE_EXTENSION_URI = '{BBE1A952-AA13-448E-AADC-164F8A28A991}';
const TIMELINE_CACHE_EXTENSION_URI = '{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}';
const SPARKLINE_GROUP_ATTRIBUTES = new Set([
  'type', 'lineWeight', 'dateAxis', 'markers', 'high', 'low', 'first', 'last', 'negative',
  'displayXAxis', 'rightToLeft', 'displayHidden', 'displayEmptyCellsAs', 'manualMin', 'manualMax',
  'colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow',
]);
const SPARKLINE_COLOR_NODES = new Set([
  'colorSeries', 'colorNegative', 'colorAxis', 'colorMarkers', 'colorFirst', 'colorLast', 'colorHigh', 'colorLow',
]);

export function detectWorksheetCapabilities(files: Record<string, Uint8Array>, pkg: OpcPackageGraph): CompatibilityFeatureDetection[] {
  const detections: CompatibilityFeatureDetection[] = [];
  const worksheetNames = readWorksheetNames(files, pkg);
  // Drawing containers also hold charts and comments. Only an image relationship
  // establishes image ownership, including media at non-standard package paths.
  for (const [part, relationships] of Object.entries(pkg.relationships)) {
    for (const relationship of relationships) {
      if (relationship.type.replace(/\/+$/, '').endsWith('/image')) {
        detections.push({ feature: 'images', location: `${part}#${relationship.id}` });
      }
    }
  }
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
        for (const extension of children(node, 'ext')) {
          if (!isCanonicallyOwnedWorksheetExtension(extension, worksheetNames, pkg, part)) detections.push({ feature: 'unknown-extension', location: `${part}#${extension.attrs.uri ?? 'ext'}`, reason: 'Worksheet extension is retained byte-for-byte from the source package' });
        }
      }
    }
  }
  const sharedStringsRelation = (pkg.relationships[pkg.workbookPart] ?? []).find((relation) => relation.type.replace(/\/+$/, '').endsWith('/sharedStrings'));
  const sharedStringsPart = sharedStringsRelation ? resolvePart(pkg.workbookPart, sharedStringsRelation.target) : 'xl/sharedStrings.xml';
  const sharedStrings = files[sharedStringsPart];
  if (sharedStrings && /<(?:\w+:)?r(?:\s|>)/.test(strFromU8(sharedStrings))) detections.push({ feature: 'rich-text', location: sharedStringsPart });
  return deduplicateDetections(detections);
}

export function detectWorkbookCapabilities(files: Record<string, Uint8Array>, pkg: OpcPackageGraph): CompatibilityFeatureDetection[] {
  const bytes = files[pkg.workbookPart];
  if (!bytes) return [];
  const workbook = descendants(parseXml(strFromU8(bytes)), 'workbook')[0];
  if (!workbook) return [];
  const detections: CompatibilityFeatureDetection[] = [];
  for (const extensionList of children(workbook, 'extLst')) {
    for (const extension of children(extensionList, 'ext')) {
      if (!isCanonicallyOwnedWorkbookControlExtension(extension, pkg.nativePivotGraph?.controls ?? [])) {
        detections.push({ feature: 'unknown-extension', location: `${pkg.workbookPart}#${extension.attrs.uri ?? 'ext'}`, reason: 'Workbook extension is retained from the source package but has no canonical structural owner' });
      }
    }
  }
  return deduplicateDetections(detections);
}

export function isCanonicallyOwnedSparklineExtension(extension: XmlNode, worksheetNames: ReadonlySet<string>): boolean {
  if (extension.attrs.uri?.toUpperCase() !== SPARKLINE_GROUPS_EXTENSION_URI) return false;
  if (!hasOnlyAttributes(extension, new Set(['uri'])) || !hasOnlyWhitespace(extension.text) || extension.children.length !== 1) return false;
  const groups = extension.children[0]!;
  if (localName(groups.name) !== 'sparklineGroups' || !hasOnlyAttributes(groups, new Set()) || !hasOnlyWhitespace(groups.text) || groups.children.length === 0) return false;
  let sparklineCount = 0;
  const owned = groups.children.every((group) => {
    if (localName(group.name) !== 'sparklineGroup' || !hasOnlyAttributes(group, SPARKLINE_GROUP_ATTRIBUTES) || !hasOnlyWhitespace(group.text)) return false;
    if (group.attrs.type !== undefined && !['line', 'column', 'win-loss'].includes(group.attrs.type)) return false;
    if (['lineWeight', 'manualMin', 'manualMax'].some((name) => group.attrs[name] !== undefined && !Number.isFinite(Number(group.attrs[name])))) return false;
    if ((group.attrs.manualMin === undefined) !== (group.attrs.manualMax === undefined)) return false;
    if (['dateAxis', 'markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'rightToLeft', 'displayHidden']
      .some((name) => group.attrs[name] !== undefined && !/^(?:0|1|true|false)$/.test(group.attrs[name]!))) return false;
    if ([...SPARKLINE_COLOR_NODES].some((name) => group.attrs[name] !== undefined && !isSparklineColor(group.attrs[name]!))) return false;
    const names = group.children.map((node) => localName(node.name));
    if (names.filter((name) => name === 'sparklines').length !== 1
      || names.some((name) => name !== 'sparklines' && !SPARKLINE_COLOR_NODES.has(name))
      || names.filter((name) => SPARKLINE_COLOR_NODES.has(name)).length !== new Set(names.filter((name) => SPARKLINE_COLOR_NODES.has(name))).size) return false;
    return group.children.every((node) => {
      if (localName(node.name) === 'sparklines') {
        return hasOnlyAttributes(node, new Set()) && hasOnlyWhitespace(node.text) && node.children.every((sparkline) => {
          if (localName(sparkline.name) !== 'sparkline' || !hasOnlyAttributes(sparkline, new Set()) || !hasOnlyWhitespace(sparkline.text)) return false;
          const fields = sparkline.children;
          const valid = fields.length === 2
            && fields.every((field) => (localName(field.name) === 'f' || localName(field.name) === 'sqref')
              && hasOnlyAttributes(field, new Set()) && field.children.length === 0 && field.text.trim().length > 0)
            && fields.filter((field) => localName(field.name) === 'f').length === 1
            && fields.filter((field) => localName(field.name) === 'sqref').length === 1
            && isOwnedSparklineFormula(fields.find((field) => localName(field.name) === 'f')!.text, worksheetNames)
            && /^\$?[A-Z]+\$?\d+$/i.test(fields.find((field) => localName(field.name) === 'sqref')!.text.trim());
          if (valid) sparklineCount += 1;
          return valid;
        });
      }
      return hasOnlyAttributes(node, new Set(['rgb', 'val'])) && node.children.length === 0 && hasOnlyWhitespace(node.text)
        && isSparklineColor(node.attrs.rgb ?? node.attrs.val ?? '');
    });
  });
  return owned && sparklineCount > 0;
}

function isOwnedSparklineFormula(formula: string, worksheetNames: ReadonlySet<string>): boolean {
  const match = /^'?((?:[^']|'')+)'?!\s*(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(formula.trim());
  return Boolean(match && worksheetNames.has(match[1]!.replaceAll("''", "'")));
}

function isCanonicallyOwnedWorksheetExtension(extension: XmlNode, worksheetNames: ReadonlySet<string>, pkg: OpcPackageGraph, sheetPart: string): boolean {
  if (isCanonicallyOwnedSparklineExtension(extension, worksheetNames)) return true;
  return isCanonicallyOwnedNativeControlExtension(extension, pkg.nativePivotGraph?.controls?.filter((entry) => entry.sheetPart === sheetPart) ?? []);
}

export function isCanonicallyOwnedNativeControlExtension(extension: XmlNode, sourceControls: readonly NativePivotControlDefinition[]): boolean {
  const uri = extension.attrs.uri?.toUpperCase();
  const control = uri === SLICER_LIST_EXTENSION_URI ? { kind: 'slicer' as const, container: 'slicerList', item: 'slicer' }
    : uri === TIMELINE_REFS_EXTENSION_URI ? { kind: 'timeline' as const, container: 'timelineRefs', item: 'timelineRef' }
      : undefined;
  return control ? isCanonicallyOwnedControlReferenceExtension(extension, sourceControls, control, 'relationshipId') : false;
}

export function isCanonicallyOwnedWorkbookControlExtension(extension: XmlNode, sourceControls: readonly NativePivotControlDefinition[]): boolean {
  const uri = extension.attrs.uri?.toUpperCase();
  const control = uri === SLICER_CACHE_EXTENSION_URI ? { kind: 'slicer' as const, container: 'slicerCaches', item: 'slicerCache' }
    : uri === TIMELINE_CACHE_EXTENSION_URI ? { kind: 'timeline' as const, container: 'timelineCacheRefs', item: 'timelineCacheRef' }
      : undefined;
  return control ? isCanonicallyOwnedControlReferenceExtension(extension, sourceControls, control, 'cacheRelationshipId') : false;
}

function isCanonicallyOwnedControlReferenceExtension(
  extension: XmlNode,
  sourceControls: readonly NativePivotControlDefinition[],
  control: { kind: NativePivotControlDefinition['kind']; container: string; item: string },
  relationshipKey: 'relationshipId' | 'cacheRelationshipId',
): boolean {
  if (!hasOnlyAttributes(extension, new Set(['uri'])) || !hasOnlyWhitespace(extension.text) || extension.children.length !== 1) return false;
  const container = extension.children[0]!;
  if (localName(container.name) !== control.container || !hasOnlyAttributes(container, new Set()) || !hasOnlyWhitespace(container.text)) return false;
  const references = container.children;
  if (!references.every((reference) => localName(reference.name) === control.item
    && hasOnlyAttributes(reference, new Set(['r:id'])) && Boolean(reference.attrs['r:id'])
    && reference.children.length === 0 && hasOnlyWhitespace(reference.text))) return false;
  const referenceIds = references.map((reference) => reference.attrs['r:id']!);
  if (new Set(referenceIds).size !== referenceIds.length) return false;
  const ownedControls = sourceControls.filter((entry) => entry.kind === control.kind);
  return !ownedControls.some((entry) => !entry.valid)
    && referenceIds.every((id) => ownedControls.some((entry) => entry[relationshipKey] === id));
}

function isSparklineColor(value: string): boolean {
  return /^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value);
}

function readWorksheetNames(files: Record<string, Uint8Array>, pkg: OpcPackageGraph): ReadonlySet<string> {
  const workbookBytes = files[pkg.workbookPart];
  if (!workbookBytes) return new Set();
  const workbook = descendants(parseXml(strFromU8(workbookBytes)), 'workbook')[0];
  return new Set(children(child(workbook, 'sheets'), 'sheet').flatMap((sheet) => sheet.attrs.name ? [sheet.attrs.name] : []));
}

function hasOnlyAttributes(node: XmlNode, allowed: ReadonlySet<string>): boolean {
  return Object.keys(node.attrs).every((name) => name.startsWith('xmlns') || allowed.has(name));
}

function hasOnlyWhitespace(value: string): boolean {
  return value.trim().length === 0;
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
