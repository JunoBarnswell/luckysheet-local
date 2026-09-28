import { strFromU8 } from 'fflate';
import { child, children, descendants, localName, parseXml, type XmlNode } from './xml';
import type { CompatibilityFeatureDetection } from './compatibility-report';
import type { NativePivotControlDefinition, OpcPackageGraph } from './types';
import { NATIVE_DOCUMENT_STRUCTURAL_CAPABILITY_POLICY as structuralPolicy } from './generated-structural-capability-policy';

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
  [structuralPolicy.features.preservedNativeChart]: capability(structuralPolicy.features.preservedNativeChart, 'full', 'none', 'none', 'none', 'full', 'The native chart part is retained byte-for-byte because this runtime does not own its complete semantic editor.'),
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
  [structuralPolicy.features.unknownExtension]: capability(structuralPolicy.features.unknownExtension, 'full', 'none', 'none', 'none', 'full'),
  [structuralPolicy.features.extendedValidation]: capability(structuralPolicy.features.extendedValidation, 'full', 'none', 'none', 'none', 'full'),
  [structuralPolicy.features.extendedConditionalFormat]: capability(structuralPolicy.features.extendedConditionalFormat, 'full', 'none', 'none', 'none', 'full'),
  [structuralPolicy.features.unknownWorksheetNode]: capability(structuralPolicy.features.unknownWorksheetNode, 'full', 'none', 'none', 'none', 'none'),
  [structuralPolicy.features.unknownWorkbookNode]: capability(structuralPolicy.features.unknownWorkbookNode, 'full', 'none', 'none', 'none', 'none'),
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

const WORKSHEET_NODES = new Map<string, string>(Object.entries(structuralPolicy.worksheet.featureNodes));
const STRUCTURAL_NODES = new Set<string>(structuralPolicy.worksheet.structuralNodes);
const SPARKLINE_GROUP_ATTRIBUTES = new Set<string>(structuralPolicy.worksheet.sparkline.groupAttributes);
const SPARKLINE_COLOR_NODES = new Set<string>(structuralPolicy.worksheet.sparkline.colorNodes);
const WORKBOOK_ROOT_NODES = new Set<string>(structuralPolicy.workbook.rootNodes);
const WORKSHEET_ROOT_ATTRIBUTES = new Set<string>(structuralPolicy.worksheet.rootAttributes);
const WORKBOOK_ROOT_ATTRIBUTES = new Set<string>(structuralPolicy.workbook.rootAttributes);
const WORKBOOK_NODE_ATTRIBUTES = new Map<string, ReadonlySet<string>>(
  Object.entries(structuralPolicy.workbook.nodes).map(([name, policy]) => [name, new Set<string>(policy.attributes)]),
);

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
    if (!bytes) {
      detections.push({ feature: structuralPolicy.features.unknownWorksheetNode, location: part, reason: 'The package graph references a worksheet part that is missing from the source package' });
      continue;
    }
    const root = packageRoot(bytes, structuralPolicy.worksheet.rootElement);
    if (!root) {
      detections.push({ feature: structuralPolicy.features.unknownWorksheetNode, location: part, reason: 'The referenced package part does not have one canonical worksheet root' });
      continue;
    }
    if (!hasOnlyAttributes(root, WORKSHEET_ROOT_ATTRIBUTES)) {
      detections.push({ feature: structuralPolicy.features.unknownWorksheetNode, location: `${part}#worksheet@attributes`, reason: 'Worksheet root attributes are not represented by the canonical writer' });
    }
    detectDuplicateRootNodes(root, structuralPolicy.worksheet.singletonNodes, part, structuralPolicy.features.unknownWorksheetNode, detections);
    const view = child(child(root, 'sheetViews'), 'sheetView');
    const pane = child(view, 'pane');
    if (pane) detections.push({ feature: pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit' ? 'freeze' : 'split', location: part });
    if (children(child(root, 'cols'), 'col').some((node) => Number(node.attrs.outlineLevel ?? 0) > 0)
      || children(child(root, 'sheetData'), 'row').some((node) => Number(node.attrs.outlineLevel ?? 0) > 0)) detections.push({ feature: 'outline', location: part });
    for (const node of root.children) {
      const name = localName(node.name);
      const feature = WORKSHEET_NODES.get(name);
      if (feature) detections.push({ feature, location: part });
      else if (!STRUCTURAL_NODES.has(name)) detections.push({ feature: structuralPolicy.features.unknownWorksheetNode, location: `${part}#${name}`, reason: `No validated reader/writer contract exists for worksheet node <${name}>` });
      if (name === 'sheetPr') validateWorksheetPropertiesNode(node, part, detections);
      if (name === 'sheetFormatPr') validateWorksheetFormatNode(node, part, detections);
      if (name === structuralPolicy.worksheet.extensionListNode) {
        const extendedRules = structuralPolicy.worksheet.extendedRuleNodes;
        if (extendedRules.validation.some((ruleNode) => descendants(node, ruleNode).length > 0)) detections.push({ feature: structuralPolicy.features.extendedValidation, location: `${part}#${structuralPolicy.worksheet.extensionListNode}`, reason: 'Extended data validation is preserved in the source package but is not editable' });
        if (extendedRules.conditionalFormat.some((ruleNode) => descendants(node, ruleNode).length > 0)) detections.push({ feature: structuralPolicy.features.extendedConditionalFormat, location: `${part}#${structuralPolicy.worksheet.extensionListNode}`, reason: 'Extended conditional formatting is preserved in the source package but is not editable' });
        for (const extension of children(node, structuralPolicy.worksheet.extensionNode)) {
          if (!isCanonicallyOwnedWorksheetExtension(extension, worksheetNames, pkg, part)) detections.push({ feature: structuralPolicy.features.unknownExtension, location: `${part}#${extension.attrs.uri ?? 'ext'}`, reason: 'Worksheet extension is retained byte-for-byte from the source package' });
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
  if (!bytes) return [{ feature: structuralPolicy.features.unknownWorkbookNode, location: pkg.workbookPart, reason: 'The package graph references a workbook part that is missing from the source package' }];
  const workbook = packageRoot(bytes, structuralPolicy.workbook.rootElement);
  if (!workbook) return [{ feature: structuralPolicy.features.unknownWorkbookNode, location: pkg.workbookPart, reason: 'The referenced package part does not have one canonical workbook root' }];
  const detections: CompatibilityFeatureDetection[] = [];
  detectDuplicateRootNodes(workbook, structuralPolicy.workbook.singletonNodes, pkg.workbookPart, structuralPolicy.features.unknownWorkbookNode, detections);
  const detectUnownedNode = (location: string, reason: string): void => {
    detections.push({ feature: structuralPolicy.features.unknownWorkbookNode, location: `${pkg.workbookPart}#${location}`, reason });
  };
  if (!hasOnlyAttributes(workbook, WORKBOOK_ROOT_ATTRIBUTES)) detectUnownedNode('workbook@attributes', 'Workbook root attributes are not represented by the canonical writer');
  for (const node of workbook.children) {
    const name = localName(node.name);
    if (!WORKBOOK_ROOT_NODES.has(name)) {
      detectUnownedNode(name, `No canonical reader/writer owner exists for workbook node <${name}>`);
      continue;
    }
    if (name === 'workbookPr' && (!hasOnlyAttributes(node, WORKBOOK_NODE_ATTRIBUTES.get(name) ?? new Set()) || node.children.length > 0 || !hasOnlyWhitespace(node.text))) {
      detectUnownedNode('workbookPr', 'Only workbookPr@date1904 is represented by the canonical writer');
    }
    if (name === 'pivotCaches') {
      const policy = structuralPolicy.workbook.nodes.pivotCaches;
      if (!hasOnlyAttributes(node, WORKBOOK_NODE_ATTRIBUTES.get(name) ?? new Set()) || !hasOnlyWhitespace(node.text)) {
        detectUnownedNode('pivotCaches', 'Pivot-cache container metadata is not represented by the canonical writer');
      }
      for (const cache of node.children) {
        if (localName(cache.name) !== policy.child.name
          || !hasOnlyAttributes(cache, new Set(policy.child.attributes))
          || (policy.child.mustBeLeaf && cache.children.length > 0)
          || (policy.child.mustHaveWhitespaceText && !hasOnlyWhitespace(cache.text))) {
          detectUnownedNode('pivotCaches#pivotCache', 'Pivot-cache metadata or child markup is not represented by the canonical writer');
        }
      }
    }
    if (name === 'sheets') {
      const policy = structuralPolicy.workbook.nodes.sheets;
      if (!hasOnlyAttributes(node, WORKBOOK_NODE_ATTRIBUTES.get(name) ?? new Set()) || !hasOnlyWhitespace(node.text)) detectUnownedNode('sheets', 'Workbook sheet-container metadata is not represented by the canonical writer');
      for (const sheet of node.children) {
        if (localName(sheet.name) !== policy.child.name
          || !hasOnlyAttributes(sheet, new Set(policy.child.attributes))
          || (policy.child.mustBeLeaf && sheet.children.length > 0)
          || (policy.child.mustHaveWhitespaceText && !hasOnlyWhitespace(sheet.text))
          || (sheet.attrs.state !== undefined && !(policy.child.stateValues as readonly string[]).includes(sheet.attrs.state))) {
          detectUnownedNode('sheets#sheet', 'Workbook sheet metadata or visibility state cannot be round-tripped canonically');
        }
      }
    }
    if (name === 'definedNames') {
      const policy = structuralPolicy.workbook.nodes.definedNames;
      if (!hasOnlyAttributes(node, WORKBOOK_NODE_ATTRIBUTES.get(name) ?? new Set()) || !hasOnlyWhitespace(node.text)) detectUnownedNode('definedNames', 'Workbook defined-name container metadata is not represented by the canonical writer');
      for (const definedName of node.children) {
        if (localName(definedName.name) !== policy.child.name
          || !hasOnlyAttributes(definedName, new Set(policy.child.attributes))
          || (policy.child.mustBeLeaf && definedName.children.length > 0)) {
          detectUnownedNode('definedNames#definedName', 'Defined-name attributes or child markup are not represented by the canonical writer');
        }
      }
    }
  }
  for (const extensionList of children(workbook, structuralPolicy.workbook.extensionListNode)) {
    for (const extension of children(extensionList, structuralPolicy.workbook.extensionNode)) {
      if (!isCanonicallyOwnedWorkbookControlExtension(extension, pkg.nativePivotGraph?.controls ?? [])) {
        detections.push({ feature: structuralPolicy.features.unknownExtension, location: `${pkg.workbookPart}#${extension.attrs.uri ?? 'ext'}`, reason: 'Workbook extension is retained from the source package but has no canonical structural owner' });
      }
    }
  }
  return deduplicateDetections(detections);
}

export function isCanonicallyOwnedSparklineExtension(extension: XmlNode, worksheetNames: ReadonlySet<string>): boolean {
  const policy = structuralPolicy.worksheet.sparkline;
  if (extension.attrs.uri?.toUpperCase() !== structuralPolicy.worksheet.extensionUris.sparklineGroups.toUpperCase()) return false;
  if (!hasOnlyAttributes(extension, new Set(structuralPolicy.controlExtensionAttributes)) || !hasOnlyWhitespace(extension.text) || extension.children.length !== 1) return false;
  const groups = extension.children[0]!;
  if (localName(groups.name) !== policy.containerNode || !hasOnlyAttributes(groups, new Set()) || !hasOnlyWhitespace(groups.text) || groups.children.length === 0) return false;
  let sparklineCount = 0;
  const owned = groups.children.every((group) => {
    if (localName(group.name) !== policy.groupNode || !hasOnlyAttributes(group, SPARKLINE_GROUP_ATTRIBUTES) || !hasOnlyWhitespace(group.text)) return false;
    if (group.attrs.type !== undefined && !(policy.types as readonly string[]).includes(group.attrs.type)) return false;
    if (policy.numericAttributes.some((name) => group.attrs[name] !== undefined && !Number.isFinite(Number(group.attrs[name])))) return false;
    if ((group.attrs.manualMin === undefined) !== (group.attrs.manualMax === undefined)) return false;
    if (policy.booleanAttributes.some((name) => group.attrs[name] !== undefined && !/^(?:0|1|true|false)$/.test(group.attrs[name]!))) return false;
    if (policy.colorAttributes.some((name) => group.attrs[name] !== undefined && !isSparklineColor(group.attrs[name]!))) return false;
    const names = group.children.map((node) => localName(node.name));
    if (names.filter((name) => name === policy.collectionNode).length !== 1
      || names.some((name) => name !== policy.collectionNode && !SPARKLINE_COLOR_NODES.has(name))
      || names.filter((name) => SPARKLINE_COLOR_NODES.has(name)).length !== new Set(names.filter((name) => SPARKLINE_COLOR_NODES.has(name))).size) return false;
    return group.children.every((node) => {
      if (localName(node.name) === policy.collectionNode) {
        return hasOnlyAttributes(node, new Set()) && hasOnlyWhitespace(node.text) && node.children.every((sparkline) => {
          if (localName(sparkline.name) !== policy.itemNode || !hasOnlyAttributes(sparkline, new Set()) || !hasOnlyWhitespace(sparkline.text)) return false;
          const fields = sparkline.children;
          const valid = fields.length === 2
            && fields.every((field) => (localName(field.name) === policy.formulaNode || localName(field.name) === policy.cellReferenceNode)
              && hasOnlyAttributes(field, new Set()) && field.children.length === 0 && field.text.trim().length > 0)
            && fields.filter((field) => localName(field.name) === policy.formulaNode).length === 1
            && fields.filter((field) => localName(field.name) === policy.cellReferenceNode).length === 1
            && isOwnedSparklineFormula(fields.find((field) => localName(field.name) === policy.formulaNode)!.text, worksheetNames)
            && /^\$?[A-Z]+\$?\d+$/i.test(fields.find((field) => localName(field.name) === policy.cellReferenceNode)!.text.trim());
          if (valid) sparklineCount += 1;
          return valid;
        });
      }
      return hasOnlyAttributes(node, new Set(policy.cellColorAttributes)) && node.children.length === 0 && hasOnlyWhitespace(node.text)
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
  const control = structuralPolicy.worksheet.controlExtensions.find((entry) => uri === structuralPolicy.worksheet.extensionUris[entry.uriKey].toUpperCase());
  return control ? isCanonicallyOwnedControlReferenceExtension(extension, sourceControls, control) : false;
}

export function isCanonicallyOwnedWorkbookControlExtension(extension: XmlNode, sourceControls: readonly NativePivotControlDefinition[]): boolean {
  const uri = extension.attrs.uri?.toUpperCase();
  const control = structuralPolicy.workbook.controlExtensions.find((entry) => uri === structuralPolicy.workbook.extensionUris[entry.uriKey].toUpperCase());
  return control ? isCanonicallyOwnedControlReferenceExtension(extension, sourceControls, control) : false;
}

function isCanonicallyOwnedControlReferenceExtension(
  extension: XmlNode,
  sourceControls: readonly NativePivotControlDefinition[],
  control: { kind: NativePivotControlDefinition['kind']; containerNode: string; itemNode: string; relationshipField: 'relationshipId' | 'cacheRelationshipId' },
): boolean {
  if (!hasOnlyAttributes(extension, new Set(structuralPolicy.controlExtensionAttributes)) || !hasOnlyWhitespace(extension.text) || extension.children.length !== 1) return false;
  const container = extension.children[0]!;
  if (localName(container.name) !== control.containerNode || !hasOnlyAttributes(container, new Set()) || !hasOnlyWhitespace(container.text)) return false;
  const references = container.children;
  if (!references.every((reference) => localName(reference.name) === control.itemNode
    && hasOnlyAttributes(reference, new Set(structuralPolicy.controlReferenceAttributes)) && Boolean(reference.attrs['r:id'])
    && reference.children.length === 0 && hasOnlyWhitespace(reference.text))) return false;
  const referenceIds = references.map((reference) => reference.attrs['r:id']!);
  if (new Set(referenceIds).size !== referenceIds.length) return false;
  const ownedControls = sourceControls.filter((entry) => entry.kind === control.kind);
  return !ownedControls.some((entry) => !entry.valid)
    && referenceIds.every((id) => ownedControls.some((entry) => entry[control.relationshipField] === id));
}

function isSparklineColor(value: string): boolean {
  return /^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value);
}

function readWorksheetNames(files: Record<string, Uint8Array>, pkg: OpcPackageGraph): ReadonlySet<string> {
  const workbookBytes = files[pkg.workbookPart];
  if (!workbookBytes) return new Set();
  const workbook = packageRoot(workbookBytes, structuralPolicy.workbook.rootElement);
  return new Set(children(child(workbook, 'sheets'), structuralPolicy.workbook.nodes.sheets.child.name).flatMap((sheet) => sheet.attrs.name ? [sheet.attrs.name] : []));
}

function hasOnlyAttributes(node: XmlNode, allowed: ReadonlySet<string>): boolean {
  return Object.keys(node.attrs).every((name) => name === 'xmlns' || name.startsWith('xmlns:') || allowed.has(name));
}

function hasOnlyWhitespace(value: string): boolean {
  return value.trim().length === 0;
}

function packageRoot(bytes: Uint8Array, expectedName: string): XmlNode | undefined {
  const document = parseXml(strFromU8(bytes));
  if (document.children.length !== 1 || !hasOnlyWhitespace(document.text)) return undefined;
  const root = document.children[0]!;
  return localName(root.name) === expectedName ? root : undefined;
}

function detectDuplicateRootNodes(
  root: XmlNode,
  singletonNames: readonly string[],
  part: string,
  feature: string,
  detections: CompatibilityFeatureDetection[],
): void {
  for (const name of singletonNames) {
    if (children(root, name).length > 1) detections.push({
      feature,
      location: `${part}#${name}`,
      reason: `Multiple <${name}> root owners cannot be represented by the canonical writer`,
    });
  }
}

function validateWorksheetPropertiesNode(node: XmlNode, part: string, detections: CompatibilityFeatureDetection[]): void {
  const policy = structuralPolicy.worksheet.nodes.sheetPr;
  const childNames = node.children.map((childNode) => localName(childNode.name));
  const invalidChild = new Set(childNames).size !== childNames.length || node.children.some((childNode) => {
    const childPolicy = policy.children[localName(childNode.name) as keyof typeof policy.children];
    if (!childPolicy || !hasOnlyAttributes(childNode, new Set(childPolicy.attributes))
      || (childPolicy.mustBeLeaf && childNode.children.length > 0)
      || (childPolicy.mustHaveWhitespaceText && !hasOnlyWhitespace(childNode.text))) return true;
    if (localName(childNode.name) === 'tabColor' && !/^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(childNode.attrs.rgb ?? '')) return true;
    return localName(childNode.name) === 'outlinePr'
      && Object.entries(childPolicy.supportedValues ?? {}).some(([attribute, values]) => childNode.attrs[attribute] !== undefined
        && !(values as readonly string[]).includes(childNode.attrs[attribute]!));
  });
  if (!hasOnlyAttributes(node, new Set(policy.attributes)) || !hasOnlyWhitespace(node.text) || invalidChild) {
    detections.push({ feature: structuralPolicy.features.unknownWorksheetNode, location: `${part}#sheetPr`, reason: 'Worksheet properties contain values not represented by the canonical writer' });
  }
}

function validateWorksheetFormatNode(node: XmlNode, part: string, detections: CompatibilityFeatureDetection[]): void {
  const policy = structuralPolicy.worksheet.nodes.sheetFormatPr;
  const invalidNumeric = policy.positiveNumericAttributes.some((attribute) => {
    const value = node.attrs[attribute];
    return value !== undefined && (!Number.isFinite(Number(value)) || Number(value) <= 0);
  });
  if (!hasOnlyAttributes(node, new Set(policy.attributes)) || !hasOnlyWhitespace(node.text)
    || (policy.mustBeLeaf && node.children.length > 0) || invalidNumeric) {
    detections.push({ feature: structuralPolicy.features.unknownWorksheetNode, location: `${part}#sheetFormatPr`, reason: 'Worksheet format contains values not represented by the canonical writer' });
  }
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
