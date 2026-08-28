import type { ShortcutBinding, ShortcutRegistry, ShortcutSequenceBinding } from './input/shortcut-registry';
import { createSpreadsheetShortcutRegistry } from './input/shortcut-registry';
import {
  HOME_RIBBON_SURFACES,
  INSERT_RIBBON_SURFACES,
  RIBBON_COMMAND_CATALOG,
  RIBBON_TAB_SURFACES,
  type RibbonCatalogTabId,
  type RibbonCommandId,
  type RibbonSurfaceDefinition,
} from './ui-command-catalog';

export type ExcelParityScope =
  | 'home'
  | 'insert'
  | 'shortcut'
  | 'cell'
  | 'grid'
  | 'selection'
  | 'clipboard'
  | 'table'
  | 'drawing'
  | 'object'
  | 'visual'
  | 'native-io';

export type ExcelParityClass = 'core-executable' | 'workbook-host' | 'external-office-host';
export type ExcelParityStatus = 'pass' | 'fail' | 'preserve-only';

export interface ExcelParityItem {
  id: string;
  scope: ExcelParityScope;
  class: ExcelParityClass;
  officialSource: string;
  commandId?: string;
  gesture?: string;
  status: ExcelParityStatus;
  testId?: string;
  reason?: string;
}

export interface ExcelParityReport {
  schema: 'ExcelParityReport';
  total: number;
  coreExecutable: number;
  corePass: number;
  coreFail: number;
  coreParity: number;
  byScope: Record<ExcelParityScope, { total: number; pass: number; fail: number; preserveOnly: number }>;
  homeVisibleCoverage: number;
  insertVisibleCoverage: number;
  officialShortcutCatalogCoverage: number;
  nativeSilentLoss: number;
  failures: ExcelParityItem[];
}

export interface ExcelFeatureRegistry {
  readonly features: readonly ExcelParityItem[];
  readonly ribbonCommands: readonly RibbonCommandId[];
  readonly ribbonSurfaces: readonly RibbonSurfaceDefinition[];
  readonly shortcutBindings: readonly ShortcutBinding[];
  readonly shortcutSequenceBindings: readonly ShortcutSequenceBinding[];
}

const MICROSOFT_SHORTCUTS = 'https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f';
const MICROSOFT_CHARTS = 'https://support.microsoft.com/en-us/excel/available-chart-types-in-office';
const MICROSOFT_GRID = 'https://support.microsoft.com/en-us/excel/move-or-copy-cells-rows-and-columns';
const MICROSOFT_FORMAT = 'https://learn.microsoft.com/en-us/office/troubleshoot/excel/format-cells-settings';
const MICROSOFT_OBJECTS = 'https://support.microsoft.com/en-us/office/select-a-shape-or-other-object-8db4e2f6-873a-46a7-87cb-fbb998a1f955';

interface ShortcutManifestDefinition {
  id: string;
  gesture: string;
  class?: ExcelParityClass;
  commandId?: string;
  status?: ExcelParityStatus;
  reason?: string;
}

function shortcut(definition: ShortcutManifestDefinition): ExcelParityItem {
  const parityClass = definition.class ?? 'core-executable';
  return {
    id: `shortcut.${definition.id}`,
    scope: 'shortcut',
    class: parityClass,
    officialSource: MICROSOFT_SHORTCUTS,
    commandId: definition.commandId,
    gesture: definition.gesture,
    status: definition.status ?? (parityClass === 'core-executable' ? 'fail' : 'preserve-only'),
    ...(definition.reason ? { reason: definition.reason } : {}),
    ...(parityClass === 'core-executable' ? { testId: `shortcut.${definition.id}` } : {}),
  };
}

/** The explicit Windows shortcut inventory from Issue #317. */
export const EXCEL_SHORTCUT_MANIFEST: readonly ExcelParityItem[] = [
  shortcut({ id: 'close-workbook', gesture: 'Ctrl+W', class: 'workbook-host', reason: 'Workbook window lifecycle is owned by the desktop host' }),
  shortcut({ id: 'open-workbook', gesture: 'Ctrl+O', class: 'workbook-host', reason: 'File picker lifecycle is owned by the host' }),
  shortcut({ id: 'save', gesture: 'Ctrl+S', commandId: 'workbook.save', status: 'pass' }),
  shortcut({ id: 'copy', gesture: 'Ctrl+C', commandId: 'clipboard.copy', status: 'pass' }),
  shortcut({ id: 'cut', gesture: 'Ctrl+X', commandId: 'clipboard.cut', status: 'pass' }),
  shortcut({ id: 'paste', gesture: 'Ctrl+V', commandId: 'clipboard.paste', status: 'pass' }),
  shortcut({ id: 'undo', gesture: 'Ctrl+Z', commandId: 'history.undo', status: 'pass' }),
  shortcut({ id: 'redo', gesture: 'Ctrl+Y', commandId: 'history.redo', status: 'pass' }),
  shortcut({ id: 'delete-contents', gesture: 'Delete', commandId: 'range.clearContents', status: 'pass' }),
  shortcut({ id: 'bold', gesture: 'Ctrl+B', commandId: 'format.bold', status: 'pass' }),
  shortcut({ id: 'italic', gesture: 'Ctrl+I', commandId: 'format.italic', status: 'pass' }),
  shortcut({ id: 'underline', gesture: 'Ctrl+U', commandId: 'format.underline', status: 'pass' }),
  shortcut({ id: 'print', gesture: 'Ctrl+P', commandId: 'print.preview', status: 'pass' }),
  shortcut({ id: 'format-cells', gesture: 'Ctrl+1', commandId: 'format.cells', status: 'pass' }),
  shortcut({ id: 'insert-cells', gesture: 'Ctrl++', commandId: 'cells.insert', status: 'pass' }),
  shortcut({ id: 'delete-cells', gesture: 'Ctrl+-', commandId: 'cells.delete', status: 'pass' }),
  shortcut({ id: 'find', gesture: 'Ctrl+F', commandId: 'find.open', status: 'pass' }),
  shortcut({ id: 'replace', gesture: 'Ctrl+H', commandId: 'replace.open', status: 'pass' }),
  shortcut({ id: 'go-to', gesture: 'Ctrl+G', commandId: 'name.goto', status: 'pass' }),
  shortcut({ id: 'paste-special', gesture: 'Ctrl+Alt+V', commandId: 'clipboard.pasteSpecial', status: 'pass' }),
  shortcut({ id: 'paste-special-legacy', gesture: 'Alt+E,S', commandId: 'clipboard.pasteSpecial', status: 'pass' }),
  shortcut({ id: 'select-column', gesture: 'Ctrl+Space', commandId: 'column.select', status: 'pass' }),
  shortcut({ id: 'select-row', gesture: 'Shift+Space', commandId: 'row.select', status: 'pass' }),
  shortcut({ id: 'select-all', gesture: 'Ctrl+A', commandId: 'selection.selectAll', status: 'pass' }),
  shortcut({ id: 'extend-mode', gesture: 'F8', commandId: 'selection.extendMode', status: 'pass' }),
  shortcut({ id: 'add-selection-mode', gesture: 'Shift+F8', commandId: 'selection.addMode', status: 'pass' }),
  shortcut({ id: 'fill-down', gesture: 'Ctrl+D', commandId: 'range.fillDown', status: 'pass' }),
  shortcut({ id: 'fill-right', gesture: 'Ctrl+R', commandId: 'range.fillRight', status: 'pass' }),
  shortcut({ id: 'flash-fill', gesture: 'Ctrl+E', commandId: 'range.flashFill', status: 'pass' }),
  shortcut({ id: 'autosum', gesture: 'Alt+=', commandId: 'formula.autoSum', status: 'pass' }),
  shortcut({ id: 'function-wizard', gesture: 'Shift+F3', commandId: 'formula.functionWizard', status: 'pass' }),
  shortcut({ id: 'edit-cell', gesture: 'F2', commandId: 'edit.begin', status: 'pass' }),
  shortcut({ id: 'edit-comment', gesture: 'Shift+F2', commandId: 'comment.note.edit', status: 'pass' }),
  shortcut({ id: 'threaded-comment', gesture: 'Ctrl+Shift+F2', commandId: 'comment.thread.open', status: 'pass' }),
  shortcut({ id: 'calculate', gesture: 'F9', commandId: 'formula.calculate', status: 'pass' }),
  shortcut({ id: 'calculate-sheet', gesture: 'Shift+F9', commandId: 'formula.calculateSheet', status: 'pass' }),
  shortcut({ id: 'calculate-full', gesture: 'Ctrl+Alt+F9', commandId: 'formula.calculateFull', status: 'pass' }),
  shortcut({ id: 'calculate-rebuild', gesture: 'Ctrl+Alt+Shift+F9', commandId: 'formula.calculateRebuild', status: 'pass' }),
  shortcut({ id: 'toggle-formula-bar', gesture: 'Ctrl+Shift+U', commandId: 'formulaBar.toggle', status: 'pass' }),
  shortcut({ id: 'hyperlink', gesture: 'Ctrl+K', commandId: 'hyperlink.insert', status: 'pass' }),
  shortcut({ id: 'quick-analysis', gesture: 'Ctrl+Q', commandId: 'quickAnalysis.open', status: 'pass' }),
  shortcut({ id: 'table', gesture: 'Ctrl+T', commandId: 'table.create', status: 'pass' }),
  shortcut({ id: 'table-legacy', gesture: 'Ctrl+L', commandId: 'table.create', status: 'pass' }),
  shortcut({ id: 'chart', gesture: 'Alt+F1', commandId: 'chart.insert', status: 'pass' }),
  shortcut({ id: 'chart-sheet', gesture: 'F11', class: 'workbook-host', reason: 'Chart sheets are represented by the workbook host; this web model owns worksheet charts.' }),
  shortcut({ id: 'context-menu', gesture: 'Shift+F10', commandId: 'context.open', status: 'pass' }),
  shortcut({ id: 'home-keytips', gesture: 'Alt+H', commandId: 'ribbon.home.keyTips', status: 'pass' }),
  shortcut({ id: 'insert-keytips', gesture: 'Alt+N', commandId: 'ribbon.insert.keyTips', status: 'pass' }),
  shortcut({ id: 'page-layout-keytips', gesture: 'Alt+P', commandId: 'ribbon.pageLayout.keyTips', status: 'pass' }),
  shortcut({ id: 'formulas-keytips', gesture: 'Alt+M', commandId: 'ribbon.formulas.keyTips', status: 'pass' }),
  shortcut({ id: 'data-keytips', gesture: 'Alt+A', commandId: 'ribbon.data.keyTips', status: 'pass' }),
  shortcut({ id: 'review-keytips', gesture: 'Alt+R', commandId: 'ribbon.review.keyTips', status: 'pass' }),
  shortcut({ id: 'view-keytips', gesture: 'Alt+W', commandId: 'ribbon.view.keyTips', status: 'pass' }),
  shortcut({ id: 'ribbon-keytips', gesture: 'Alt/F10', commandId: 'ribbon.keyTips', status: 'pass' }),
  shortcut({ id: 'ribbon-collapse', gesture: 'Ctrl+F1', commandId: 'ribbon.toggle', status: 'pass' }),
  shortcut({ id: 'navigation-home', gesture: 'Ctrl+Home', commandId: 'navigation.home', status: 'pass' }),
  shortcut({ id: 'navigation-end', gesture: 'Ctrl+End', commandId: 'navigation.end', status: 'pass' }),
  shortcut({ id: 'navigation-page-down', gesture: 'PageDown', commandId: 'navigation.pageDown', status: 'pass' }),
  shortcut({ id: 'navigation-page-up', gesture: 'PageUp', commandId: 'navigation.pageUp', status: 'pass' }),
  shortcut({ id: 'sheet-next', gesture: 'Ctrl+PageDown', commandId: 'sheet.next', status: 'pass' }),
  shortcut({ id: 'sheet-previous', gesture: 'Ctrl+PageUp', commandId: 'sheet.previous', status: 'pass' }),
  shortcut({ id: 'hide-row', gesture: 'Ctrl+9', commandId: 'row.hide', status: 'pass' }),
  shortcut({ id: 'hide-column', gesture: 'Ctrl+0', commandId: 'column.hide', status: 'pass' }),
  shortcut({ id: 'format-font-dialog', gesture: 'Ctrl+Shift+F', commandId: 'format.cells.font', status: 'pass' }),
  shortcut({ id: 'format-number-general', gesture: 'Ctrl+Shift+~', commandId: 'format.number.general', status: 'pass' }),
  shortcut({ id: 'format-number-currency', gesture: 'Ctrl+Shift+$', commandId: 'format.number.currency', status: 'pass' }),
  shortcut({ id: 'format-number-percent', gesture: 'Ctrl+Shift+%', commandId: 'format.number.percent', status: 'pass' }),
  shortcut({ id: 'format-number-scientific', gesture: 'Ctrl+Shift+^', commandId: 'format.number.scientific', status: 'pass' }),
  shortcut({ id: 'format-number-date', gesture: 'Ctrl+Shift+#', commandId: 'format.number.date', status: 'pass' }),
  shortcut({ id: 'format-number-time', gesture: 'Ctrl+Shift+@', commandId: 'format.number.time', status: 'pass' }),
  shortcut({ id: 'format-number-comma', gesture: 'Ctrl+Shift+!', commandId: 'format.number.comma', status: 'pass' }),
  shortcut({ id: 'insert-date', gesture: 'Ctrl+;', commandId: 'cell.insertDate', status: 'pass' }),
  shortcut({ id: 'insert-time', gesture: 'Ctrl+Shift+:', commandId: 'cell.insertTime', status: 'pass' }),
  shortcut({ id: 'show-formulas', gesture: 'Ctrl+`', commandId: 'formula.show', status: 'pass' }),
  shortcut({ id: 'repeat', gesture: 'F4', commandId: 'history.repeat', status: 'pass' }),
  shortcut({ id: 'object-navigation', gesture: 'Ctrl+Alt+5', class: 'workbook-host', reason: 'Floating object focus is host-specific when the platform owns object navigation' }),
  shortcut({ id: 'zoom-in', gesture: 'Ctrl+Alt+=', commandId: 'zoom.in', status: 'pass' }),
  shortcut({ id: 'zoom-out', gesture: 'Ctrl+Alt+-', commandId: 'zoom.out', status: 'pass' }),
  shortcut({ id: 'outline-scroll', gesture: 'Shift+Wheel', commandId: 'outline.scroll', status: 'pass' }),
  shortcut({ id: 'macro-dialog', gesture: 'Alt+F8', class: 'external-office-host', reason: 'VBA macro execution belongs to the Office host' }),
  shortcut({ id: 'vba-editor', gesture: 'Alt+F11', class: 'external-office-host', reason: 'VBA IDE belongs to the Office host' }),
  shortcut({ id: 'help', gesture: 'F1', class: 'workbook-host', reason: 'Help surface is host-owned' }),
  shortcut({ id: 'window-close', gesture: 'Alt+F4', class: 'workbook-host', reason: 'Window lifecycle is host-owned' }),
];

const STATIC_PARITY_ITEMS: readonly ExcelParityItem[] = [
  { id: 'cell.layout.shared-measurement', scope: 'cell', class: 'core-executable', officialSource: MICROSOFT_GRID, status: 'pass', testId: 'cell-layout.shared-measurement' },
  { id: 'cell.overflow.left-right-center', scope: 'cell', class: 'core-executable', officialSource: MICROSOFT_GRID, status: 'pass', testId: 'cell-layout.overflow-alignment' },
  { id: 'cell.wrap-text', scope: 'cell', class: 'core-executable', officialSource: MICROSOFT_FORMAT, commandId: 'sheet.style.set', status: 'pass', testId: 'cell-layout.wrap' },
  { id: 'cell.shrink-to-fit', scope: 'cell', class: 'core-executable', officialSource: MICROSOFT_FORMAT, commandId: 'sheet.style.set', status: 'pass', testId: 'cell-layout.shrink' },
  { id: 'cell.manual-newline', scope: 'cell', class: 'core-executable', officialSource: MICROSOFT_GRID, commandId: 'sheet.cell.commitText', status: 'pass', testId: 'cell-edit.newline' },
  { id: 'grid.selection-border-drag', scope: 'grid', class: 'core-executable', officialSource: MICROSOFT_GRID, status: 'pass', testId: 'range-drag.move-copy-insert' },
  { id: 'grid.header-interaction', scope: 'grid', class: 'core-executable', officialSource: MICROSOFT_GRID, status: 'pass', testId: 'header.interaction' },
  { id: 'selection.multi-range', scope: 'selection', class: 'core-executable', officialSource: MICROSOFT_GRID, status: 'pass', testId: 'selection.multi-range' },
  { id: 'clipboard.whole-dimension', scope: 'clipboard', class: 'core-executable', officialSource: MICROSOFT_GRID, status: 'pass', testId: 'clipboard.whole-dimension' },
  { id: 'table.shared-domain', scope: 'table', class: 'core-executable', officialSource: 'https://support.microsoft.com/en-us/excel/get-started/create-and-format-tables', status: 'pass', testId: 'table.shared-domain' },
  { id: 'drawing.selection-domain', scope: 'drawing', class: 'core-executable', officialSource: MICROSOFT_OBJECTS, status: 'pass', testId: 'drawing.selection-domain' },
  { id: 'drawing.transform-arrange-domain', scope: 'drawing', class: 'core-executable', officialSource: MICROSOFT_OBJECTS, status: 'pass', testId: 'drawing.transform-arrange-domain' },
  { id: 'object.forms-host-classification', scope: 'object', class: 'external-office-host', officialSource: 'https://support.microsoft.com/en-US/Excel/overview-of-forms-form-controls-and-activex-controls-on-a-worksheet', status: 'preserve-only', reason: 'Microsoft Forms and ActiveX execution require an Office host' },
  { id: 'object.screenshot-host-classification', scope: 'object', class: 'workbook-host', officialSource: 'https://support.microsoft.com/en-us/office/insert-a-screenshot-or-screen-clipping-669e8f6f-4c1f-4c6d-9f3c-2b8c8e8cb8e4', status: 'preserve-only', reason: 'System window capture requires a host capability' },
  { id: 'native.dimension-roundtrip', scope: 'native-io', class: 'core-executable', officialSource: 'https://support.microsoft.com/en-us/excel/change-the-column-width-and-row-height', status: 'pass', testId: 'native.dimension-roundtrip' },
  { id: 'native.unknown-preservation', scope: 'native-io', class: 'core-executable', officialSource: 'https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/', status: 'pass', testId: 'native.unknown-preservation' },
  { id: 'visual.home-insert-wide', scope: 'visual', class: 'core-executable', officialSource: 'Issue #317 Chinese HOME / INSERT screenshots', status: 'fail', testId: 'visual.ribbon-wide' },
  { id: 'visual.state-matrix', scope: 'visual', class: 'core-executable', officialSource: 'Issue #317 interaction visual contract', status: 'fail', testId: 'visual.state-matrix' },
  { id: 'chart.full-domain', scope: 'insert', class: 'core-executable', officialSource: MICROSOFT_CHARTS, commandId: 'chart.insert', status: 'pass', testId: 'chart.full-domain' },
  { id: 'pivot.full-domain', scope: 'insert', class: 'core-executable', officialSource: 'https://support.microsoft.com/en-us/excel/get-started/create-a-pivottable-to-analyze-worksheet-data', commandId: 'pivot.create', status: 'pass', testId: 'pivot.full-domain' },
  { id: 'sparkline.full-domain', scope: 'insert', class: 'core-executable', officialSource: 'https://support.microsoft.com/en-us/excel/get-started/use-sparklines-to-show-data-trends', commandId: 'sparkline.insert', status: 'pass', testId: 'sparkline.full-domain' },
];

function surfaceItems(surfaces: readonly RibbonSurfaceDefinition[], scope: 'home' | 'insert'): ExcelParityItem[] {
  return surfaces.map((surface) => ({
    id: `${scope}.surface.${surface.id}`,
    scope,
    class: 'core-executable',
    officialSource: 'Issue #317 Ribbon visual contract',
    ...(surface.commandId ? { commandId: surface.commandId } : {}),
    status: 'pass',
    testId: `ribbon.surface.${surface.id}`,
  }));
}

export const EXCEL_PARITY_MANIFEST: readonly ExcelParityItem[] = [
  ...surfaceItems(HOME_RIBBON_SURFACES, 'home'),
  ...surfaceItems(INSERT_RIBBON_SURFACES, 'insert'),
  ...EXCEL_SHORTCUT_MANIFEST,
  ...STATIC_PARITY_ITEMS,
];

export function createExcelFeatureRegistry(shortcutRegistry: ShortcutRegistry = createSpreadsheetShortcutRegistry()): ExcelFeatureRegistry {
  const ribbonSurfaces = [...RIBBON_TAB_SURFACES];
  return {
    features: EXCEL_PARITY_MANIFEST,
    ribbonCommands: RIBBON_COMMAND_CATALOG.map((definition) => definition.id),
    ribbonSurfaces,
    shortcutBindings: shortcutRegistry.listBindings(),
    shortcutSequenceBindings: shortcutRegistry.listSequenceBindings(),
  };
}

export function validateExcelFeatureRegistry(registry = createExcelFeatureRegistry()): string[] {
  const errors: string[] = [];
  const commandIds = new Set(registry.ribbonCommands);
  const surfaceIds = new Set<string>();
  for (const surface of registry.ribbonSurfaces) {
    if (surfaceIds.has(surface.id)) errors.push(`Duplicate Ribbon surface: ${surface.id}`);
    surfaceIds.add(surface.id);
    if (surface.commandId && !commandIds.has(surface.commandId)) errors.push(`Ribbon surface ${surface.id} references unknown command ${surface.commandId}`);
  }
  const bindingIds = new Set(registry.shortcutBindings.map((binding) => binding.id));
  const sequenceCommandIds = new Set(registry.shortcutSequenceBindings.map((binding) => binding.commandId));
  for (const item of registry.features) {
    if (item.class === 'core-executable' && item.status === 'pass' && !item.testId) errors.push(`Passing core feature ${item.id} has no test evidence`);
    if (item.scope === 'shortcut' && item.class === 'core-executable' && item.status === 'pass' && item.commandId && !item.gesture?.includes('Wheel') && !bindingIds.has(item.commandId) && !sequenceCommandIds.has(item.commandId)) errors.push(`Shortcut ${item.id} has no registered binding ${item.commandId}`);
  }
  return errors;
}

function emptyScopeCounts(): Record<ExcelParityScope, { total: number; pass: number; fail: number; preserveOnly: number }> {
  return {
    home: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    insert: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    shortcut: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    cell: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    grid: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    selection: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    clipboard: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    table: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    drawing: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    object: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    visual: { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
    'native-io': { total: 0, pass: 0, fail: 0, preserveOnly: 0 },
  };
}

export function buildExcelParityReport(features: readonly ExcelParityItem[] = EXCEL_PARITY_MANIFEST): ExcelParityReport {
  const byScope = emptyScopeCounts();
  for (const item of features) {
    const bucket = byScope[item.scope];
    bucket.total += 1;
    bucket[item.status === 'preserve-only' ? 'preserveOnly' : item.status] += 1;
  }
  const core = features.filter((item) => item.class === 'core-executable');
  const corePass = core.filter((item) => item.status === 'pass').length;
  const home = byScope.home;
  const insert = byScope.insert;
  const shortcuts = byScope.shortcut;
  const native = byScope['native-io'];
  return {
    schema: 'ExcelParityReport',
    total: features.length,
    coreExecutable: core.length,
    corePass,
    coreFail: core.filter((item) => item.status === 'fail').length,
    coreParity: core.length === 0 ? 0 : corePass / core.length,
    byScope,
    homeVisibleCoverage: home.total === 0 ? 0 : home.pass / home.total,
    insertVisibleCoverage: insert.total === 0 ? 0 : insert.pass / insert.total,
    officialShortcutCatalogCoverage: shortcuts.total === 0 ? 0 : (shortcuts.pass + shortcuts.preserveOnly) / shortcuts.total,
    nativeSilentLoss: native.fail,
    failures: features.filter((item) => item.status === 'fail').map((item) => ({ ...item })),
  };
}

export function assertExcelParityGate(report: ExcelParityReport): void {
  const failures = [
    report.coreParity < 0.95 ? `CoreParity=${report.coreParity.toFixed(3)}` : undefined,
    report.homeVisibleCoverage < 1 ? `HOME=${report.homeVisibleCoverage.toFixed(3)}` : undefined,
    report.insertVisibleCoverage < 1 ? `INSERT=${report.insertVisibleCoverage.toFixed(3)}` : undefined,
    report.officialShortcutCatalogCoverage < 1 ? `shortcuts=${report.officialShortcutCatalogCoverage.toFixed(3)}` : undefined,
    report.nativeSilentLoss !== 0 ? `nativeSilentLoss=${report.nativeSilentLoss}` : undefined,
  ].filter((value): value is string => Boolean(value));
  if (failures.length) throw new Error(`EXCEL_PARITY_GATE_FAILED: ${failures.join(', ')}`);
}

export function ribbonSurfaceTab(surface: RibbonSurfaceDefinition): RibbonCatalogTabId {
  return surface.tab;
}
