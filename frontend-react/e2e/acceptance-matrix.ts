import {
  RIBBON_LAYOUT_SPECS,
  RIBBON_TAB_SURFACES,
  validateRibbonLayoutSpecs,
  type RibbonCommandId,
  type RibbonLayoutNode,
  type RibbonLayoutSpec,
  type RibbonSurfaceDefinition,
} from '@react-sheets/spreadsheet-app';
import {
  INSERT_CHART_VARIANTS,
  INSERT_CONNECTOR_VARIANTS,
  INSERT_SHAPE_GALLERY,
  INSERT_SPARKLINE_VARIANTS,
} from '../apps/web/src/components/insert-ribbon-catalog';

export type AcceptanceLayer = 'domain' | 'contract' | 'interaction' | 'mutation' | 'undo-redo' | 'parity' | 'collaboration' | 'browser' | 'visual' | 'persistence' | 'xlsx' | 'transient';

export type AcceptanceRibbonTab = 'home' | 'insert' | 'pageLayout' | 'formulas' | 'data';
export type AcceptanceDomain =
  | 'ribbon'
  | 'page-layout'
  | 'formulas'
  | 'data'
  | 'designer-shell'
  | 'selection'
  | 'keyboard'
  | 'editing'
  | 'clipboard'
  | 'formula'
  | 'history'
  | 'worksheet'
  | 'ime'
  | 'pointer-selection'
  | 'dimension'
  | 'context-menu'
  | 'fill'
  | 'persistence'
  | 'drawing'
  | 'transient-ui'
  | 'workbook-hub'
  | 'routing'
  | 'backstage'
  | 'responsive-shell'
  | 'collaboration';
export type AcceptancePersistence = 'not-applicable' | 'page-session' | 'server-roundtrip' | 'xlsx-roundtrip';
export type AcceptanceCollaboration = 'not-applicable' | 'server-snapshot-replay';
export type AcceptanceTransient = 'not-applicable' | 'must-not-persist';

export interface AcceptanceContract {
  readonly issue: 289;
  readonly source: string;
  readonly domain: AcceptanceDomain;
  readonly persistence: AcceptancePersistence;
  readonly collaboration: AcceptanceCollaboration;
  readonly transient: AcceptanceTransient;
}

export interface AcceptanceCase {
  readonly id: string;
  readonly title: string;
  readonly layers: readonly AcceptanceLayer[];
  readonly evidence: readonly string[];
  readonly contract?: AcceptanceContract;
}

export const ACCEPTANCE_RIBBON_TABS = ['home', 'insert', 'pageLayout', 'formulas', 'data'] as const;

export const ACCEPTANCE_VIEWPORTS = [
  { width: 1280, height: 720 },
  { width: 1366, height: 768 },
  { width: 1440, height: 900 },
  { width: 1920, height: 1080 },
] as const;

export const ACCEPTANCE_LOCALES = ['zh-CN', 'en-US'] as const;
export type AcceptanceLocale = typeof ACCEPTANCE_LOCALES[number];

function caseId(prefix: string, value: string): string {
  return `${prefix}-${value.toUpperCase().replaceAll(/[^A-Z0-9]+/g, '-')}`;
}

function contract(source: string, domain: AcceptanceDomain, persistence: AcceptancePersistence = 'not-applicable', collaboration: AcceptanceCollaboration = 'not-applicable', transient: AcceptanceTransient = 'not-applicable'): AcceptanceContract {
  return { issue: 289, source, domain, persistence, collaboration, transient };
}

function ribbonTabPrefix(tab: AcceptanceRibbonTab): string {
  return ({ home: 'H', insert: 'I', pageLayout: 'L', formulas: 'F', data: 'D' } as const)[tab];
}

function surfaceCase(surface: RibbonSurfaceDefinition): AcceptanceCase & { readonly surface: RibbonSurfaceDefinition } {
  const tab = surface.tab as AcceptanceRibbonTab;
  return {
    id: caseId(ribbonTabPrefix(tab), `surface-${surface.id}`),
    title: `${surface.tab} surface ${surface.id}`,
    layers: ['contract', 'interaction', 'browser', 'visual'],
    evidence: ['catalog identity', 'semantic command/control identity', 'responsive ribbon placement'],
    contract: contract('RIBBON_TAB_SURFACES', 'ribbon'),
    surface,
  };
}

export const HOME_SURFACE_CASES = RIBBON_TAB_SURFACES
  .filter((surface) => surface.tab === 'home')
  .map(surfaceCase);

export const INSERT_SURFACE_CASES = RIBBON_TAB_SURFACES
  .filter((surface) => surface.tab === 'insert')
  .map(surfaceCase);

export interface RibbonLayoutAcceptanceCase extends AcceptanceCase {
  readonly tab: AcceptanceRibbonTab;
  readonly groupId: string;
  readonly nodeId: string;
  readonly nodeKind: RibbonLayoutNode['kind'];
  readonly commandIds: readonly RibbonCommandId[];
  readonly surfaceId?: string;
}

function layoutNodeCommandIds(node: RibbonLayoutNode): readonly RibbonCommandId[] {
  switch (node.kind) {
    case 'command':
    case 'checkbox':
    case 'spinner':
    case 'combo':
    case 'launcher':
      return [node.commandId];
    case 'split':
      return [node.primary, ...node.items.map((item) => item.commandId)];
    case 'dropdown':
      return [node.trigger, ...node.items.map((item) => item.commandId)];
    case 'surface': {
      const surface = RIBBON_TAB_SURFACES.find((candidate) => candidate.id === node.surfaceId);
      return surface?.commandId ? [surface.commandId] : [];
    }
    case 'column':
    case 'row':
    case 'stack':
      return node.children.flatMap(layoutNodeCommandIds);
    case 'separator':
      return [];
  }
}

interface RibbonLayoutTarget {
  readonly groupId: string;
  readonly node: RibbonLayoutNode;
  readonly commandIds: readonly RibbonCommandId[];
  readonly surfaceId?: string;
}

function layoutTargets(spec: RibbonLayoutSpec): readonly RibbonLayoutTarget[] {
  const targets: RibbonLayoutTarget[] = [];
  const visit = (groupId: string, node: RibbonLayoutNode): void => {
    if (node.kind === 'column' || node.kind === 'row' || node.kind === 'stack') {
      node.children.forEach((child) => visit(groupId, child));
      return;
    }
    if (node.kind === 'separator') return;
    const commandIds = layoutNodeCommandIds(node);
    if (commandIds.length > 0) targets.push({ groupId, node, commandIds, ...(node.kind === 'surface' ? { surfaceId: node.surfaceId } : {}) });
  };
  for (const group of spec.groups) for (const node of group.children) visit(group.id, node);
  return targets;
}

function layoutCase(tab: AcceptanceRibbonTab, target: RibbonLayoutTarget): RibbonLayoutAcceptanceCase {
  return {
    id: caseId(ribbonTabPrefix(tab), `layout-${target.node.id}`),
    title: `${tab} layout node ${target.node.id}`,
    layers: ['contract', 'interaction', 'browser', 'visual'],
    evidence: ['layout-tree ownership', 'semantic command identity', 'responsive placement'],
    contract: contract('RIBBON_LAYOUT_SPECS', tab === 'pageLayout' ? 'page-layout' : tab === 'formulas' ? 'formulas' : tab === 'data' ? 'data' : 'ribbon'),
    tab,
    groupId: target.groupId,
    nodeId: target.node.id,
    nodeKind: target.node.kind,
    commandIds: target.commandIds,
    ...(target.surfaceId ? { surfaceId: target.surfaceId } : {}),
  };
}

export const RIBBON_LAYOUT_CASES: readonly RibbonLayoutAcceptanceCase[] = ACCEPTANCE_RIBBON_TABS.flatMap((tab) => layoutTargets(RIBBON_LAYOUT_SPECS[tab]).map((target) => layoutCase(tab, target)));
export const PAGE_LAYOUT_CASES = RIBBON_LAYOUT_CASES.filter((entry) => entry.tab === 'pageLayout');
export const FORMULAS_CASES = RIBBON_LAYOUT_CASES.filter((entry) => entry.tab === 'formulas');
export const DATA_CASES = RIBBON_LAYOUT_CASES.filter((entry) => entry.tab === 'data');

export interface CoreInteractionAcceptanceCase extends AcceptanceCase {
  readonly sourceSpec: string;
  readonly testTitle: string;
}

const coreInteraction = (id: string, title: string, domain: AcceptanceDomain, sourceSpec: string, layers: readonly AcceptanceLayer[], evidence: readonly string[], persistence: AcceptancePersistence = 'not-applicable', transient: AcceptanceTransient = 'not-applicable'): CoreInteractionAcceptanceCase => ({
  id,
  title,
  layers,
  evidence,
  contract: contract(sourceSpec, domain, persistence, 'not-applicable', transient),
  sourceSpec,
  testTitle: title,
});

/** The 29 concrete cases already present in the core E2E suites remain traceable
 * in the product matrix. The matrix records their contract; it does not invent
 * a second runtime or command implementation. */
export const CORE_INTERACTION_CASES: readonly CoreInteractionAcceptanceCase[] = [
  coreInteraction('CORE-001', 'Designer Demo shell exposes the fixed 1280x720 geometry and real palette entry', 'designer-shell', 'e2e/spreadsheet.spec.ts', ['interaction', 'browser', 'visual'], ['shell geometry', 'command palette entry']),
  coreInteraction('CORE-002', 'selection updates the name box', 'selection', 'e2e/spreadsheet.spec.ts', ['interaction', 'browser'], ['active cell', 'Name Box address']),
  coreInteraction('CORE-003', 'keyboard navigation moves the active cell', 'keyboard', 'e2e/spreadsheet.spec.ts', ['interaction', 'browser'], ['arrow navigation', 'active cell identity']),
  coreInteraction('CORE-004', 'direct typing commits cell editing', 'editing', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser'], ['editor commit', 'canonical cell value']),
  coreInteraction('CORE-005', 'clicking another cell commits the old editor before accepting the new value', 'editing', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser'], ['old editor commit', 'new target identity']),
  coreInteraction('CORE-006', 'clipboard copy and paste round-trip', 'clipboard', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser', 'persistence'], ['clipboard payload', 'paste result'], 'page-session'),
  coreInteraction('CORE-007', 'formula bar cancel restores the draft', 'editing', 'e2e/spreadsheet.spec.ts', ['interaction', 'browser', 'transient'], ['draft state', 'cancel restores value'], 'not-applicable', 'must-not-persist'),
  coreInteraction('CORE-008', 'undo reverts the last edit', 'history', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'undo-redo', 'browser'], ['one history entry', 'inverse mutation']),
  coreInteraction('CORE-009', 'adds and activates a new editable worksheet from the sheet tab control', 'worksheet', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser', 'persistence'], ['sheet creation', 'active sheet', 'editable value'], 'page-session'),
  coreInteraction('CORE-010', 'F2 editing keeps Chinese IME text intact before one committed operation', 'ime', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'undo-redo', 'browser'], ['composition lifecycle', 'single commit']),
  coreInteraction('CORE-011', 'a dragged A2:C9 selection edits and places the editor at active cell C9', 'pointer-selection', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser'], ['drag range', 'release cell', 'editor anchor']),
  coreInteraction('CORE-012', 'a dragged selection accepts direct typing at the release cell', 'pointer-selection', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser'], ['release target', 'typed value']),
  coreInteraction('CORE-013', 'reverse drag keeps the release cell active for F2 editing', 'pointer-selection', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser'], ['reverse selection', 'F2 target']),
  coreInteraction('CORE-014', 'row and column header drags commit the final header target', 'dimension', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser'], ['row header', 'column header', 'final target']),
  coreInteraction('CORE-015', 'Excel-style column width paths share one multi-column transaction surface', 'dimension', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'undo-redo', 'browser', 'persistence'], ['drag width', 'menu width', 'single transaction'], 'page-session'),
  coreInteraction('CORE-016', 'right click changes the command target before the context menu opens', 'context-menu', 'e2e/spreadsheet.spec.ts', ['interaction', 'contract', 'browser'], ['context target', 'menu identity']),
  coreInteraction('CORE-017', 'fill handle follows the primary range bottom-right', 'fill', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser', 'persistence'], ['fill handle', 'primary range', 'filled values'], 'page-session'),
  coreInteraction('CORE-018', 'local formula calculation works in the page memory session without an API request', 'formula', 'e2e/spreadsheet.spec.ts', ['domain', 'mutation', 'browser', 'persistence'], ['formula result', 'session reset after reload', 'no API request'], 'page-session'),
  coreInteraction('CORE-019', 'Home ribbon opens shared format, sort, find, and paste dialogs without rendering Add-ins', 'ribbon', 'e2e/spreadsheet.spec.ts', ['contract', 'interaction', 'browser'], ['shared entrypoint', 'dialog identity', 'no Add-ins surface']),
  coreInteraction('CORE-020', 'Selection Pane selects, renames, and toggles a drawing through host callbacks', 'drawing', 'e2e/spreadsheet.spec.ts', ['interaction', 'mutation', 'browser', 'persistence'], ['selection pane', 'host callback', 'drawing state'], 'page-session'),
  coreInteraction('CORE-021', 'Format Painter enters a transient Home state and completes after one target selection', 'transient-ui', 'e2e/spreadsheet.spec.ts', ['contract', 'interaction', 'mutation', 'browser', 'transient'], ['active transient state', 'one target selection', 'style mutation only'], 'not-applicable', 'must-not-persist'),
  coreInteraction('CORE-022', 'renders the file-center shell and never exposes inactive workbook commands', 'workbook-hub', 'e2e/workbook-hub.spec.ts', ['contract', 'interaction', 'browser'], ['file-center shell', 'inactive command absence']),
  coreInteraction('CORE-023', 'creates a local workbook then preserves its session through Backstage', 'backstage', 'e2e/workbook-hub.spec.ts', ['interaction', 'mutation', 'browser', 'persistence'], ['create workbook', 'Backstage session'], 'page-session'),
  coreInteraction('CORE-024', 'does not turn an unknown unauthenticated route into a blank local workbook', 'routing', 'e2e/workbook-hub.spec.ts', ['contract', 'interaction', 'browser'], ['unknown route', 'no blank workbook']),
  coreInteraction('CORE-025', 'keeps the supplied 1672×941 shell geometry', 'responsive-shell', 'e2e/workbook-hub.spec.ts', ['interaction', 'browser', 'visual'], ['fixed shell geometry', 'no overflow']),
  coreInteraction('CORE-026', 'Designer Shell visual contract 1280x720', 'designer-shell', 'e2e/designer-visual.spec.ts', ['contract', 'browser', 'visual'], ['shell geometry', 'Name Box', 'palette state']),
  coreInteraction('CORE-027', 'Designer Shell visual contract 1366x768', 'designer-shell', 'e2e/designer-visual.spec.ts', ['contract', 'browser', 'visual'], ['shell geometry', 'Name Box', 'palette state']),
  coreInteraction('CORE-028', 'Designer Shell visual contract 1440x900', 'designer-shell', 'e2e/designer-visual.spec.ts', ['contract', 'browser', 'visual'], ['shell geometry', 'Name Box', 'palette state']),
  coreInteraction('CORE-029', 'Designer Shell visual contract 1920x1080', 'designer-shell', 'e2e/designer-visual.spec.ts', ['contract', 'browser', 'visual'], ['shell geometry', 'Name Box', 'palette state']),
];

export interface VisualGoldenCase {
  readonly id: string;
  readonly tab: AcceptanceRibbonTab;
  readonly locale: AcceptanceLocale;
  readonly viewport: (typeof ACCEPTANCE_VIEWPORTS)[number];
  readonly shellScreenshot: string;
  readonly screenshot: string;
  readonly states: readonly ['shell', 'ribbon', 'locale', 'compact-collapse', 'menu-clipping', 'dialog-clipping'];
}

export const RIBBON_VISUAL_GOLDEN_CASES: readonly VisualGoldenCase[] = ACCEPTANCE_LOCALES.flatMap((locale) => ACCEPTANCE_VIEWPORTS.flatMap((viewport) => ACCEPTANCE_RIBBON_TABS.map((tab) => ({
  id: `VISUAL-${tab}-${locale}-${viewport.width}X${viewport.height}`.toUpperCase(),
  tab,
  locale,
  viewport,
  shellScreenshot: `shell/${locale}/${viewport.width}x${viewport.height}.png`,
  screenshot: `ribbon/${tab}/${locale}/${viewport.width}x${viewport.height}.png`,
  states: ['shell', 'ribbon', 'locale', 'compact-collapse', 'menu-clipping', 'dialog-clipping'] as const,
}))));

export const CONNECTED_PARITY_CASES: readonly AcceptanceCase[] = [
  {
    id: 'C-SERVER-JAVA-SNAPSHOT-REPLAY',
    title: 'Client A commit is persisted by the Java server and replayed identically by Client B',
    layers: ['contract', 'mutation', 'parity', 'collaboration', 'persistence', 'transient'],
    evidence: ['Client A canonical mutation journal', 'Java committed snapshot', 'Client B replay', 'canonical equality', 'transient state absent'],
    contract: contract('connected-server parity fixture', 'collaboration', 'server-roundtrip', 'server-snapshot-replay', 'must-not-persist'),
  },
];

export const TRANSIENT_NEGATIVE_CASES: readonly AcceptanceCase[] = [
  {
    id: 'T-FORMAT-PAINTER-NO-SNAPSHOT',
    title: 'Format Painter active state is never written to workbook snapshot or collaboration journal',
    layers: ['contract', 'interaction', 'persistence', 'collaboration', 'transient'],
    evidence: ['active UI state', 'snapshot unchanged before target', 'journal excludes transient state', 'rejected target retains active state'],
    contract: contract('format-painter transient contract', 'transient-ui', 'not-applicable', 'server-snapshot-replay', 'must-not-persist'),
  },
];

export const XLSX_CONTRACT_CASES: readonly AcceptanceCase[] = [
  {
    id: 'XLSX-CORE-MATRIX-PREFLIGHT',
    title: 'Core matrix declares XLSX exchange as an explicit capability boundary',
    layers: ['contract', 'parity', 'persistence', 'xlsx'],
    evidence: ['OOXML preflight', 'supported capability result', 'UNSUPPORTED_FEATURE for unsupported behavior'],
    contract: contract('XLSX exchange preflight', 'persistence', 'xlsx-roundtrip'),
  },
];

export const HOME_BEHAVIOR_CASES: readonly AcceptanceCase[] = [
  { id: 'H-HISTORY-UNDO-REDO', title: 'History performs one reversible authored transaction', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['one operation', 'Undo once', 'Redo once'] },
  { id: 'H-CLIPBOARD-PASTE-SPECIAL', title: 'Clipboard and Paste Special preserve the selected semantic mode', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['typed sparse payload', 'one paste mutation', 'round-trip'] },
  { id: 'H-FONT-STYLE-TRANSACTION', title: 'Font controls commit through the canonical style mutation', layers: ['domain', 'contract', 'browser'], evidence: ['active/disabled state', 'style snapshot', 'one history entry'] },
  { id: 'H-ALIGNMENT-MERGE-TRANSACTION', title: 'Alignment and Merge use the same selection target', layers: ['domain', 'contract', 'browser'], evidence: ['selection target', 'canonical style/merge result', 'reject invalid range'] },
  { id: 'H-NUMBER-FORMAT-TRANSACTION', title: 'Number format controls preserve exact custom format semantics', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['canonical format value', 'Undo/Redo', 'page-session readback'] },
  { id: 'H-STYLES-GALLERY', title: 'Styles gallery exposes every catalog entry without a second list', layers: ['contract', 'browser', 'visual'], evidence: ['catalog-driven membership', 'semantic command identity', 'menu geometry'] },
  { id: 'H-CELLS-DIMENSIONS', title: 'Cells and Dimensions controls share the selection-aware command path', layers: ['domain', 'contract', 'browser'], evidence: ['row/column extent', 'hide/unhide', 'single mutation'] },
  { id: 'H-EDITING-AUTOSUM-FILL', title: 'Editing, AutoSum, Fill, Sort, Filter, Clear and Find retain authored semantics', layers: ['domain', 'contract', 'browser', 'persistence'], evidence: ['result snapshot', 'one operation', 'same-session checkpoint'] },
];

export const INSERT_VARIANT_GROUPS = [
  { id: 'chart', commandId: 'chartBuilder' as const, rootSurfaceId: 'charts.gallery', variants: INSERT_CHART_VARIANTS },
  { id: 'sparkline', commandId: 'sparkline' as const, rootSurfaceId: 'sparklines.gallery', variants: INSERT_SPARKLINE_VARIANTS },
  { id: 'shape', commandId: 'shapesLines' as const, rootSurfaceId: 'illustrations.shape', variants: INSERT_SHAPE_GALLERY.flatMap((category) => category.variants) },
  { id: 'connector', commandId: 'shapesLines' as const, rootSurfaceId: 'illustrations.shape', variants: INSERT_CONNECTOR_VARIANTS },
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
  { id: 'G-NAME-BOX-HIGH-ADDRESS', title: 'Name Box accepts high canonical coordinates', layers: ['contract', 'browser', 'persistence'], evidence: ['exact address', 'resolved cell', 'same-session readback'] },
  { id: 'G-DRAG-AUTO-SCROLL', title: 'Pointer drag auto-scroll preserves the release target', layers: ['domain', 'browser'], evidence: ['selection state', 'release cell'] },
  { id: 'G-PASTE-FILL-SPILL-EXTENT', title: 'Paste, Fill and Spill grow the canonical sheet extent', layers: ['domain', 'contract', 'parity', 'persistence'], evidence: ['extent ensure', 'overlay result', 'round-trip'] },
  { id: 'G-STRUCTURAL-SHIFT', title: 'Structural insert and shift transform dependent coordinates', layers: ['domain', 'contract', 'parity', 'persistence'], evidence: ['participant transform', 'history inverse', 'replay'] },
  { id: 'G-HIGH-INDEX-DIMENSIONS', title: 'High-index resize and hide use exact dimension commands', layers: ['domain', 'contract', 'browser'], evidence: ['dimension state', 'visibility projection'] },
  { id: 'G-MILLION-ROW-VIRTUAL-GEOMETRY', title: 'Million-row geometry is virtualized without truncation', layers: ['domain', 'browser', 'visual'], evidence: ['virtual extent', 'bounded render work', 'no hardcoded row cap'] },
  { id: 'G-COLLABORATION-EXTENT-CONVERGENCE', title: 'Local, page session and collaboration converge on the same extent', layers: ['contract', 'parity', 'persistence'], evidence: ['canonical snapshot', 'remote replay', 'extent equality'] },
  { id: 'G-XLSX-EXTENT-PREFLIGHT', title: 'XLSX extent compatibility is explicit before exchange', layers: ['contract', 'parity', 'persistence'], evidence: ['capability preflight', 'explicit unsupported result'] },
];

export const PARITY_MUTATION_FAMILIES = [
  { id: 'structural', prefixes: ['row.', 'column.', 'sheet.', 'workbook.'] },
  { id: 'fill', prefixes: ['fill.'] },
  { id: 'paste', prefixes: ['range.paste', 'range.clear'] },
  { id: 'style-rules', prefixes: ['sheet.style.', 'conditionalFormat.', 'dataValidation.'] },
  { id: 'dimensions', prefixes: ['row.height', 'column.width', 'row.visibility', 'column.visibility'] },
  { id: 'protection', prefixes: ['sheet.protection.', 'workbook.protection.'] },
  { id: 'drawing-insert', prefixes: ['drawing.', 'chart.', 'sparkline.'] },
  { id: 'extent-ensure', prefixes: ['sheet.extent.'] },
] as const;

export const PERMISSION_ROLES = ['owner', 'editor', 'commenter', 'viewer'] as const;
export const PROTECTION_STATES = ['none', 'sheet', 'range-locked', 'range-unlocked'] as const;
export const PERMISSION_ENTRYPOINTS = ['ribbon', 'shortcut', 'context-menu', 'command-palette'] as const;

export const PERMISSION_MATRIX_CASES = PERMISSION_ROLES.flatMap((role) => PROTECTION_STATES.flatMap((protection) => PERMISSION_ENTRYPOINTS.map((entrypoint) => ({
  id: `P-${role.toUpperCase()}-${protection.toUpperCase()}-${entrypoint.toUpperCase().replaceAll('-', '_')}`,
  title: `${role} with ${protection} protection through ${entrypoint}`,
  role,
  protection,
  entrypoint,
  layers: ['contract', 'browser', 'parity'] as const,
  evidence: ['UI enabled/disabled state', 'backend allow/reject', 'same capability policy'],
}))));

export const PERSISTENCE_CASES: readonly AcceptanceCase[] = [
  { id: 'G-LOCAL-SAVE-RELOAD', title: 'Local save uses one page-session canonical workbook checkpoint', layers: ['contract', 'browser', 'persistence'], evidence: ['save checkpoint', 'same-session readback', 'reset after reload'] },
  { id: 'G-SERVER-COLLABORATION-ROUNDTRIP', title: 'Server collaboration round-trip preserves the committed snapshot', layers: ['contract', 'parity', 'persistence'], evidence: ['server commit', 'second-client replay', 'canonical equality'] },
];

export function allAcceptanceCases(): readonly AcceptanceCase[] {
  return [
    ...HOME_SURFACE_CASES,
    ...INSERT_SURFACE_CASES,
    ...RIBBON_LAYOUT_CASES,
    ...HOME_BEHAVIOR_CASES,
    ...INSERT_VARIANT_CASES,
    ...ELASTIC_GRID_CASES,
    ...PERMISSION_MATRIX_CASES,
    ...PERSISTENCE_CASES,
    ...CORE_INTERACTION_CASES,
    ...CONNECTED_PARITY_CASES,
    ...TRANSIENT_NEGATIVE_CASES,
    ...XLSX_CONTRACT_CASES,
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
    if (entry.contract && entry.contract.issue !== 289) errors.push(`${entry.id} is not traceable to issue #289`);
  }

  const homeSurfaceIds = new Set(HOME_SURFACE_CASES.map((entry) => entry.surface.id));
  const insertSurfaceIds = new Set(INSERT_SURFACE_CASES.map((entry) => entry.surface.id));
  for (const surface of RIBBON_TAB_SURFACES.filter((entry) => entry.tab === 'home')) {
    if (!homeSurfaceIds.has(surface.id)) errors.push(`missing Home acceptance case for ${surface.id}`);
  }
  for (const surface of RIBBON_TAB_SURFACES.filter((entry) => entry.tab === 'insert')) {
    if (!insertSurfaceIds.has(surface.id)) errors.push(`missing Insert acceptance case for ${surface.id}`);
  }

  for (const tab of ACCEPTANCE_RIBBON_TABS) {
    const tabCases = RIBBON_LAYOUT_CASES.filter((entry) => entry.tab === tab);
    if (tabCases.length === 0) errors.push(`missing layout acceptance cases for ${tab}`);
    const expectedTargets = layoutTargets(RIBBON_LAYOUT_SPECS[tab]);
    if (tabCases.length !== expectedTargets.length) errors.push(`${tab} layout acceptance cases are not one-to-one with layout targets`);
    for (const entry of tabCases) {
      if (entry.commandIds.length === 0) errors.push(`${entry.id} has no semantic command identity`);
      if (entry.contract?.source !== 'RIBBON_LAYOUT_SPECS') errors.push(`${entry.id} is not sourced from RIBBON_LAYOUT_SPECS`);
    }
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

  const viewportKeys = ACCEPTANCE_VIEWPORTS.map((entry) => `${entry.width}x${entry.height}`);
  if (ACCEPTANCE_VIEWPORTS.length !== 4 || new Set(viewportKeys).size !== 4 || !['1280x720', '1366x768', '1440x900', '1920x1080'].every((key) => viewportKeys.includes(key))) errors.push('visual acceptance requires all four fixed viewports');
  if (!ACCEPTANCE_LOCALES.includes('zh-CN') || !ACCEPTANCE_LOCALES.includes('en-US')) errors.push('visual acceptance requires zh-CN and en-US');
  if (PARITY_MUTATION_FAMILIES.length !== 8) errors.push('parity corpus must cover eight declared mutation families');
  if (PERMISSION_MATRIX_CASES.length !== PERMISSION_ROLES.length * PROTECTION_STATES.length * PERMISSION_ENTRYPOINTS.length) errors.push('permission matrix is incomplete');
  if (Object.keys(RIBBON_LAYOUT_SPECS).length < 5) errors.push('Home/Insert/Page Layout/Formulas/Data layout specs are incomplete');
  if (CORE_INTERACTION_CASES.length !== 29) errors.push('core interaction matrix must cover the 29 concrete E2E cases');
  if (new Set(CORE_INTERACTION_CASES.map((entry) => entry.id)).size !== CORE_INTERACTION_CASES.length) errors.push('core interaction matrix contains duplicate IDs');
  if (RIBBON_VISUAL_GOLDEN_CASES.length !== ACCEPTANCE_RIBBON_TABS.length * ACCEPTANCE_LOCALES.length * ACCEPTANCE_VIEWPORTS.length) errors.push('visual golden matrix is incomplete for tabs, locales, or viewports');
  if (RIBBON_VISUAL_GOLDEN_CASES.some((entry) => entry.screenshot.trim() === '' || entry.states.length !== 6)) errors.push('visual golden matrix has an incomplete screenshot/state contract');
  if (CONNECTED_PARITY_CASES.length === 0 || TRANSIENT_NEGATIVE_CASES.length === 0 || XLSX_CONTRACT_CASES.length === 0) errors.push('connected, transient-negative, and XLSX acceptance contracts are required');
  return errors;
}

export function commandIdsForAcceptance(): readonly RibbonCommandId[] {
  return [...new Set([
    ...HOME_SURFACE_CASES,
    ...INSERT_SURFACE_CASES,
  ].map((entry) => entry.surface.commandId).filter((id): id is RibbonCommandId => id !== undefined).concat(RIBBON_LAYOUT_CASES.flatMap((entry) => entry.commandIds)))];
}
