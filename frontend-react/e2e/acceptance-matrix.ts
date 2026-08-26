import {
  RIBBON_LAYOUT_SPECS,
  RIBBON_TAB_SURFACES,
  validateRibbonLayoutSpecs,
  type RibbonCommandId,
  type RibbonSurfaceDefinition,
} from '@react-sheets/spreadsheet-app';
import {
  INSERT_BARCODE_VARIANTS,
  INSERT_CHART_VARIANTS,
  INSERT_CONNECTOR_VARIANTS,
  INSERT_DATA_CHART_VARIANTS,
  INSERT_FORM_CONTROL_VARIANTS,
  INSERT_SHAPE_GALLERY,
  INSERT_SPARKLINE_VARIANTS,
} from '../apps/web/src/components/insert-ribbon-catalog';

export type AcceptanceLayer = 'domain' | 'contract' | 'parity' | 'browser' | 'visual' | 'persistence';

export interface AcceptanceCase {
  readonly id: string;
  readonly title: string;
  readonly layers: readonly AcceptanceLayer[];
  readonly evidence: readonly string[];
}

export const ACCEPTANCE_VIEWPORTS = [
  { width: 1280, height: 720 },
  { width: 1920, height: 1080 },
] as const;

export const ACCEPTANCE_LOCALES = ['zh-CN', 'en-US'] as const;
export type AcceptanceLocale = typeof ACCEPTANCE_LOCALES[number];

function caseId(prefix: 'H' | 'I', value: string): string {
  return `${prefix}-${value.toUpperCase().replaceAll(/[^A-Z0-9]+/g, '-')}`;
}

function surfaceCase(surface: RibbonSurfaceDefinition): AcceptanceCase & { readonly surface: RibbonSurfaceDefinition } {
  return {
    id: caseId(surface.tab === 'home' ? 'H' : 'I', `surface-${surface.id}`),
    title: `${surface.tab} surface ${surface.id}`,
    layers: ['contract', 'browser', 'visual'],
    evidence: ['catalog identity', 'semantic command/control identity', 'responsive ribbon placement'],
    surface,
  };
}

export const HOME_SURFACE_CASES = RIBBON_TAB_SURFACES
  .filter((surface) => surface.tab === 'home')
  .map(surfaceCase);

export const INSERT_SURFACE_CASES = RIBBON_TAB_SURFACES
  .filter((surface) => surface.tab === 'insert')
  .map(surfaceCase);

export const HOME_BEHAVIOR_CASES: readonly AcceptanceCase[] = [
  { id: 'H-HISTORY-UNDO-REDO', title: 'History performs one reversible authored transaction', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['one operation', 'Undo once', 'Redo once'] },
  { id: 'H-CLIPBOARD-PASTE-SPECIAL', title: 'Clipboard and Paste Special preserve the selected semantic mode', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['typed sparse payload', 'one paste mutation', 'round-trip'] },
  { id: 'H-FONT-STYLE-TRANSACTION', title: 'Font controls commit through the canonical style mutation', layers: ['domain', 'contract', 'browser'], evidence: ['active/disabled state', 'style snapshot', 'one history entry'] },
  { id: 'H-ALIGNMENT-MERGE-TRANSACTION', title: 'Alignment and Merge use the same selection target', layers: ['domain', 'contract', 'browser'], evidence: ['selection target', 'canonical style/merge result', 'reject invalid range'] },
  { id: 'H-NUMBER-FORMAT-TRANSACTION', title: 'Number format controls preserve exact custom format semantics', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['canonical format value', 'Undo/Redo', 'reload'] },
  { id: 'H-STYLES-GALLERY', title: 'Styles gallery exposes every catalog entry without a second list', layers: ['contract', 'browser', 'visual'], evidence: ['catalog-driven membership', 'semantic command identity', 'menu geometry'] },
  { id: 'H-CELLS-DIMENSIONS', title: 'Cells and Dimensions controls share the selection-aware command path', layers: ['domain', 'contract', 'browser'], evidence: ['row/column extent', 'hide/unhide', 'single mutation'] },
  { id: 'H-EDITING-AUTOSUM-FILL', title: 'Editing, AutoSum, Fill, Sort, Filter, Clear and Find retain authored semantics', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['result snapshot', 'one operation', 'save/reload'] },
];

export const INSERT_VARIANT_GROUPS = [
  { id: 'chart', commandId: 'chartBuilder' as const, rootSurfaceId: 'charts.gallery', variants: INSERT_CHART_VARIANTS },
  { id: 'data-chart', commandId: 'dataChart' as const, rootSurfaceId: 'data-charts.insert', variants: INSERT_DATA_CHART_VARIANTS },
  { id: 'barcode', commandId: 'barcode' as const, rootSurfaceId: 'charts.barcode', variants: INSERT_BARCODE_VARIANTS },
  { id: 'sparkline', commandId: 'sparkline' as const, rootSurfaceId: 'charts.sparkline', variants: INSERT_SPARKLINE_VARIANTS },
  { id: 'shape', commandId: 'shapesLines' as const, rootSurfaceId: 'illustrations.shape', variants: INSERT_SHAPE_GALLERY.flatMap((category) => category.variants) },
  { id: 'connector', commandId: 'shapesLines' as const, rootSurfaceId: 'illustrations.shape', variants: INSERT_CONNECTOR_VARIANTS },
  { id: 'form-control', commandId: 'formControls' as const, rootSurfaceId: 'illustrations.controls', variants: INSERT_FORM_CONTROL_VARIANTS },
] as const;

export const INSERT_VARIANT_CASES = INSERT_VARIANT_GROUPS.flatMap((group) => group.variants.map((variant) => ({
  id: caseId('I', `variant-${variant.id}`),
  title: `Insert variant ${variant.id}`,
  layers: ['contract', 'browser', 'visual'] as const,
  evidence: ['typed gallery catalog', 'variant semantic identity', 'canonical insert command'],
  group,
  variant,
})));

export const ELASTIC_GRID_CASES: readonly AcceptanceCase[] = [
  { id: 'G-SCROLL-HIGH-INDEX', title: 'Scroll keeps canonical geometry across high indexes', layers: ['domain', 'browser', 'visual'], evidence: ['virtual geometry', 'no clipped extent', 'no page overflow'] },
  { id: 'G-KEYBOARD-BOUNDARY', title: 'Keyboard navigation resolves high row and column addresses', layers: ['domain', 'browser'], evidence: ['Name Box address', 'active cell identity'] },
  { id: 'G-NAME-BOX-HIGH-ADDRESS', title: 'Name Box accepts high canonical coordinates', layers: ['contract', 'browser', 'persistence'], evidence: ['exact address', 'resolved cell', 'reload'] },
  { id: 'G-DRAG-AUTO-SCROLL', title: 'Pointer drag auto-scroll preserves the release target', layers: ['domain', 'browser'], evidence: ['selection state', 'release cell'] },
  { id: 'G-PASTE-FILL-SPILL-EXTENT', title: 'Paste, Fill and Spill grow the canonical sheet extent', layers: ['domain', 'contract', 'parity', 'persistence'], evidence: ['extent ensure', 'overlay result', 'round-trip'] },
  { id: 'G-STRUCTURAL-SHIFT', title: 'Structural insert and shift transform dependent coordinates', layers: ['domain', 'contract', 'parity', 'persistence'], evidence: ['participant transform', 'history inverse', 'replay'] },
  { id: 'G-HIGH-INDEX-DIMENSIONS', title: 'High-index resize and hide use exact dimension commands', layers: ['domain', 'contract', 'browser'], evidence: ['dimension state', 'visibility projection'] },
  { id: 'G-MILLION-ROW-VIRTUAL-GEOMETRY', title: 'Million-row geometry is virtualized without truncation', layers: ['domain', 'browser', 'visual'], evidence: ['virtual extent', 'bounded render work', 'no hardcoded row cap'] },
  { id: 'G-COLLABORATION-EXTENT-CONVERGENCE', title: 'Local, reload and collaboration converge on the same extent', layers: ['contract', 'parity', 'persistence'], evidence: ['canonical snapshot', 'remote replay', 'extent equality'] },
  { id: 'G-XLSX-EXTENT-PREFLIGHT', title: 'XLSX extent compatibility is explicit before exchange', layers: ['contract', 'parity', 'persistence'], evidence: ['capability preflight', 'explicit unsupported result'] },
];

export const PARITY_MUTATION_FAMILIES = [
  { id: 'structural', prefixes: ['row.', 'column.', 'sheet.', 'workbook.'] },
  { id: 'fill', prefixes: ['fill.'] },
  { id: 'paste', prefixes: ['range.paste', 'range.clear'] },
  { id: 'style-rules', prefixes: ['sheet.style.', 'conditionalFormat.', 'dataValidation.'] },
  { id: 'dimensions', prefixes: ['row.height', 'column.width', 'row.visibility', 'column.visibility'] },
  { id: 'protection', prefixes: ['sheet.protection.', 'workbook.protection.'] },
  { id: 'drawing-insert', prefixes: ['drawing.', 'chart.', 'dataChart.', 'sparkline.'] },
  { id: 'extent-ensure', prefixes: ['sheet.extent.'] },
] as const;

export const PERMISSION_ROLES = ['owner', 'editor', 'commenter', 'viewer'] as const;
export const PROTECTION_STATES = ['none', 'sheet', 'range-locked', 'range-unlocked'] as const;
export const PERMISSION_ENTRYPOINTS = ['ribbon', 'shortcut', 'context-menu', 'command-palette'] as const;

export const PERMISSION_MATRIX_CASES = PERMISSION_ROLES.flatMap((role) => PROTECTION_STATES.flatMap((protection) => PERMISSION_ENTRYPOINTS.map((entrypoint) => ({
  id: `P-${role.toUpperCase()}-${protection.toUpperCase()}-${entrypoint.toUpperCase().replaceAll('-', '_')}`,
  role,
  protection,
  entrypoint,
  layers: ['contract', 'browser', 'parity'] as const,
  evidence: ['UI enabled/disabled state', 'backend allow/reject', 'same capability policy'],
}))));

export const PERSISTENCE_CASES: readonly AcceptanceCase[] = [
  { id: 'G-LOCAL-SAVE-RELOAD', title: 'Local save and reload preserve the canonical workbook snapshot', layers: ['contract', 'browser', 'persistence'], evidence: ['save checkpoint', 'reload', 'semantic value equality'] },
  { id: 'G-SERVER-COLLABORATION-ROUNDTRIP', title: 'Server collaboration round-trip preserves the committed snapshot', layers: ['contract', 'parity', 'persistence'], evidence: ['server commit', 'second-client replay', 'canonical equality'] },
];

export function allAcceptanceCases(): readonly AcceptanceCase[] {
  return [
    ...HOME_SURFACE_CASES,
    ...INSERT_SURFACE_CASES,
    ...HOME_BEHAVIOR_CASES,
    ...INSERT_VARIANT_CASES,
    ...ELASTIC_GRID_CASES,
    ...PERMISSION_MATRIX_CASES,
    ...PERSISTENCE_CASES,
  ];
}

export function validateAcceptanceMatrix(): readonly string[] {
  const errors: string[] = [...validateRibbonLayoutSpecs()];
  const cases = allAcceptanceCases();
  const ids = new Set<string>();
  for (const entry of cases) {
    if (ids.has(entry.id)) errors.push(`duplicate acceptance case ${entry.id}`);
    ids.add(entry.id);
    if (entry.layers.length === 0) errors.push(`${entry.id} has no acceptance layer`);
    if (entry.evidence.length === 0) errors.push(`${entry.id} has no evidence contract`);
  }

  const homeSurfaceIds = new Set(HOME_SURFACE_CASES.map((entry) => entry.surface.id));
  const insertSurfaceIds = new Set(INSERT_SURFACE_CASES.map((entry) => entry.surface.id));
  for (const surface of RIBBON_TAB_SURFACES.filter((entry) => entry.tab === 'home')) {
    if (!homeSurfaceIds.has(surface.id)) errors.push(`missing Home acceptance case for ${surface.id}`);
  }
  for (const surface of RIBBON_TAB_SURFACES.filter((entry) => entry.tab === 'insert')) {
    if (!insertSurfaceIds.has(surface.id)) errors.push(`missing Insert acceptance case for ${surface.id}`);
  }

  const variantIds = new Set<string>();
  for (const group of INSERT_VARIANT_GROUPS) {
    const root = RIBBON_TAB_SURFACES.find((surface) => surface.id === group.rootSurfaceId);
    if (!root || root.commandId !== group.commandId) errors.push(`${group.id} does not resolve to its catalog root surface`);
    for (const variant of group.variants) {
      if (variantIds.has(variant.id)) errors.push(`duplicate Insert variant ${variant.id}`);
      variantIds.add(variant.id);
    }
  }
  if (INSERT_VARIANT_CASES.length !== variantIds.size) errors.push('Insert variant acceptance cases are not one-to-one with variants');

  if (ACCEPTANCE_VIEWPORTS.length !== 2 || !ACCEPTANCE_VIEWPORTS.some((entry) => entry.width === 1280 && entry.height === 720) || !ACCEPTANCE_VIEWPORTS.some((entry) => entry.width === 1920 && entry.height === 1080)) errors.push('visual acceptance requires 1280x720 and 1920x1080');
  if (!ACCEPTANCE_LOCALES.includes('zh-CN') || !ACCEPTANCE_LOCALES.includes('en-US')) errors.push('visual acceptance requires zh-CN and en-US');
  if (PARITY_MUTATION_FAMILIES.length !== 8) errors.push('parity corpus must cover eight declared mutation families');
  if (PERMISSION_MATRIX_CASES.length !== PERMISSION_ROLES.length * PROTECTION_STATES.length * PERMISSION_ENTRYPOINTS.length) errors.push('permission matrix is incomplete');
  if (Object.keys(RIBBON_LAYOUT_SPECS).length < 5) errors.push('Home/Insert/Page Layout/Formulas/Data layout specs are incomplete');
  return errors;
}

export function commandIdsForAcceptance(): readonly RibbonCommandId[] {
  return [...new Set([...HOME_SURFACE_CASES, ...INSERT_SURFACE_CASES].map((entry) => entry.surface.commandId).filter((id): id is RibbonCommandId => id !== undefined))];
}
