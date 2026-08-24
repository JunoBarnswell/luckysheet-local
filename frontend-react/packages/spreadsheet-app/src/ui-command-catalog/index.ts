import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { AppPhase, SidebarPanelId, UiSessionIntent } from '../types';

export type RibbonCatalogTabId =
  | 'file'
  | 'home'
  | 'insert'
  | 'pageLayout'
  | 'formulas'
  | 'data'
  | 'review'
  | 'view'
  | 'settings'
  | 'automate'
  | 'pivotAnalyze'
  | 'pivotDesign';

export type RibbonGroupId =
  | 'workbook'
  | 'scripts'
  | 'calculation'
  | 'formulaAudit'
  | 'definedNames'
  | 'pageSetup'
  | 'scaleToFit'
  | 'sheetOptions'
  | 'history'
  | 'clipboard'
  | 'font'
  | 'alignment'
  | 'number'
  | 'styles'
  | 'cells'
  | 'insertCells'
  | 'editing'
  | 'tablesPivots'
  | 'chartsVisuals'
  | 'illustrations'
  | 'functions'
  | 'sortFilter'
  | 'dataTools'
  | 'outline'
  | 'findTransform'
  | 'comments'
  | 'notesLinks'
  | 'protection'
  | 'historyAudit'
  | 'freezePanes'
  | 'zoom'
  | 'printLayout'
  | 'appearanceFiles'
  | 'settings'
  | 'pivotAnalyze'
  | 'pivotDesign';

export type RibbonCommandId =
  | 'save'
  | 'exportXlsx'
  | 'importXlsx'
  | 'exportXlsxView'
  | 'importXlsxView'
  | 'openAutomate'
  | 'runSampleScript'
  | 'startRecording'
  | 'stopRecording'
  | 'calculateNow'
  | 'goalSeek'
  | 'tracePrecedents'
  | 'traceDependents'
  | 'removeArrows'
  | 'showFormulas'
  | 'errorChecking'
  | 'evaluateFormula'
  | 'calculationAutomatic'
  | 'calculationManual'
  | 'definedNames'
  | 'pageSetup'
  | 'setPrintArea'
  | 'clearPrintArea'
  | 'printTitleRows'
  | 'printTitleColumns'
  | 'setScale100'
  | 'viewGridlines'
  | 'printGridlines'
  | 'viewHeadings'
  | 'printHeadings'
  | 'undo'
  | 'redo'
  | 'cut'
  | 'copy'
  | 'paste'
  | 'pasteSpecial'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strikethrough'
  | 'allBorders'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'alignTop'
  | 'alignMiddle'
  | 'alignBottom'
  | 'indentIncrease'
  | 'indentDecrease'
  | 'wrapText'
  | 'mergeCells'
  | 'formatCells'
  | 'numberFormatGeneral'
  | 'numberFormatCurrency'
  | 'numberFormatPercent'
  | 'numberFormatComma'
  | 'numberFormatDecimal'
  | 'numberFormatDecimalIncrease'
  | 'numberFormatDecimalDecrease'
  | 'insertRow'
  | 'insertColumn'
  | 'insertRowHome'
  | 'insertColumnHome'
  | 'shiftCells'
  | 'clearContents'
  | 'clearFormats'
  | 'clearAll'
  | 'autoSum'
  | 'fillDown'
  | 'fillRight'
  | 'sortRange'
  | 'conditionalFormat'
  | 'cellTemplate'
  | 'cellEditor'
  | 'pivotTable'
  | 'quickPivot'
  | 'chartBuilder'
  | 'columnChart'
  | 'sparkline'
  | 'quickSparkline'
  | 'shapesLines'
  | 'rectangle'
  | 'bringDrawingForward'
  | 'sendDrawingBackward'
  | 'removeDrawing'
  | 'insertFunction'
  | 'deleteRow'
  | 'deleteColumn'
  | 'sortAscending'
  | 'sortDescending'
  | 'customSort'
  | 'dataModel'
  | 'createDataTable'
  | 'formatAsTable'
  | 'totalRow'
  | 'dataValidation'
  | 'filterSelection'
  | 'clearFilter'
  | 'groupRows'
  | 'ungroupRows'
  | 'groupColumns'
  | 'ungroupColumns'
  | 'showLevel1'
  | 'showLevel2'
  | 'showLevel3'
  | 'subtotal'
  | 'removeDuplicates'
  | 'textToColumns'
  | 'findReplace'
  | 'goTo'
  | 'transpose'
  | 'flipHorizontal'
  | 'flipVertical'
  | 'splitByDelimiter'
  | 'newComment'
  | 'resolveComment'
  | 'showComments'
  | 'newNote'
  | 'insertLink'
  | 'protectSelection'
  | 'unprotect'
  | 'revisionLog'
  | 'freezeTopRow'
  | 'freezeFirstColumn'
  | 'freezeAtSelection'
  | 'unfreezeAll'
  | 'zoomIn'
  | 'zoomOut'
  | 'zoomReset'
  | 'printPdf'
  | 'bandedRows'
  | 'settings'
  | 'commandPalette'
  | 'pivotRefresh'
  | 'pivotFieldList';

export type RibbonTextKey = `groups.${RibbonGroupId}` | `commands.${RibbonCommandId}`;

export type RibbonDisplay = 'large' | 'medium' | 'small';

export type RibbonIconName =
  | 'align-center'
  | 'align-left'
  | 'align-right'
  | 'borders'
  | 'bold'
  | 'calculator'
  | 'chart'
  | 'check-circle'
  | 'clipboard'
  | 'comma'
  | 'columns'
  | 'comment'
  | 'copy'
  | 'dollar-sign'
  | 'decimal-decrease'
  | 'decimal-increase'
  | 'filter'
  | 'fill-down'
  | 'fill-right'
  | 'freeze'
  | 'function'
  | 'history'
  | 'italic'
  | 'indent-decrease'
  | 'indent-increase'
  | 'layout'
  | 'lock'
  | 'merge-cells'
  | 'minimize'
  | 'more-horizontal'
  | 'percent'
  | 'plus'
  | 'printer'
  | 'redo'
  | 'rows'
  | 'scissors'
  | 'search'
  | 'shape-square'
  | 'share'
  | 'sliders'
  | 'sort'
  | 'sparkles'
  | 'sparkline'
  | 'star'
  | 'strikethrough'
  | 'table'
  | 'table-pivot'
  | 'trash'
  | 'type'
  | 'underline'
  | 'undo'
  | 'wrap-text'
  | 'x'
  | 'zoom-in'
  | 'zoom-out';

export interface RibbonCellStyleContext {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  align?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  indent?: number;
  wrapText?: boolean;
  numberFormat?: string;
}

export interface RibbonCommandActions {
  onCopy: () => void;
  onCut: () => void;
  onPaste: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onSave: () => void;
  onExportXlsx: () => void;
  onImportXlsx: () => void;
  onRecalculate: () => void;
  onTracePrecedents: () => void;
  onTraceDependents: () => void;
  onRemoveArrows: () => void;
  onToggleShowFormulas: () => void;
  onScanFormulaErrors: () => void;
  onEvaluateFormula: () => void;
  onOpenPrintLayout: () => void;
  onSetPrintArea: () => void;
  onClearPrintArea: () => void;
  onSetPrintTitleRows: () => void;
  onSetPrintTitleColumns: () => void;
  onSetPrintScale: (scale: number) => void;
  onToggleViewGridlines: () => void;
  onTogglePrintGridlines: () => void;
  onToggleViewHeadings: () => void;
  onTogglePrintHeadings: () => void;
  onAutoSum: () => void;
  onFill: (direction: 'down' | 'right') => void;
  onFreezeAtPrimary: () => void;
  onCreatePivot: () => CommandDescriptor | undefined;
  onCreateChart: () => CommandDescriptor | undefined;
  onCreateSparkline: () => CommandDescriptor | undefined;
  onCreateShape: () => CommandDescriptor | undefined;
  onBringDrawingForward: () => CommandDescriptor | undefined;
  onSendDrawingBackward: () => CommandDescriptor | undefined;
  onRemoveDrawing: () => CommandDescriptor | undefined;
  onCreateSheetTable: () => void;
  onCreateDataTable: () => void;
  onToggleSheetTableTotalRow: () => CommandDescriptor | undefined;
  onApplyFilterSelection: () => CommandDescriptor | undefined;
  onClearFilter: () => CommandDescriptor | undefined;
  onGroupRows: () => CommandDescriptor | undefined;
  onUngroupRows: () => CommandDescriptor | undefined;
  onGroupColumns: () => CommandDescriptor | undefined;
  onUngroupColumns: () => CommandDescriptor | undefined;
  onSubtotal: () => CommandDescriptor | undefined;
  onRemoveDuplicates: () => CommandDescriptor | undefined;
  onTextToColumns: () => CommandDescriptor | undefined;
  onResolveComment: () => void;
  onProtectSelection: () => void;
  onUnprotectSelection: () => void;
  onShowOutlineLevel: (level: 1 | 2 | 3) => void;
  onTransposeSelection: () => void;
  onFlipSelection: (axis: 'h' | 'v') => void;
  onSplitByDelimiter: () => void;
  onToggleBandedRows: () => void;
  onSetRecalculationMode: (mode: 'automatic' | 'manual') => void;
  onOpenDefinedNames: () => void;
}

export interface RibbonCommandContext {
  phase: AppPhase;
  disabled: boolean;
  cellStyle: RibbonCellStyleContext;
  canExecute?: (commandId: string, params?: unknown) => boolean;
  /** Host-owned builder; it must resolve the active selection/current region. */
  buildSortDescriptor?: (ascending: boolean) => CommandDescriptor | undefined;
  /** Host-owned Create PivotTable dialog entry point. */
  openCreatePivotDialog?: () => void;
  activePivot?: { sheetId: string; pivotId: string };
  actions: RibbonCommandActions;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
  sampleAutomationScript: string;
}

export type RibbonCommandResult =
  | { type: 'command'; descriptor: CommandDescriptor }
  | { type: 'intent'; intent: UiSessionIntent }
  | { type: 'callback'; invoke: () => void };

export interface CommandDefinition {
  readonly id: RibbonCommandId;
  readonly tab: RibbonCatalogTabId;
  readonly group: RibbonGroupId;
  readonly labelKey: RibbonTextKey;
  readonly tooltipKey?: RibbonTextKey;
  readonly icon?: RibbonIconName;
  readonly shortcut?: string;
  readonly priority: number;
  readonly display: RibbonDisplay;
  readonly collapseInto?: RibbonGroupId;
  readonly commandId?: string;
  readonly when?: (context: RibbonCommandContext) => boolean;
  readonly enabled?: (context: RibbonCommandContext) => boolean;
  readonly active?: (context: RibbonCommandContext) => boolean;
  readonly build: (context: RibbonCommandContext) => RibbonCommandResult | undefined;
}

export interface RibbonGroupDefinition {
  readonly id: RibbonGroupId;
  readonly tab: RibbonCatalogTabId;
  readonly labelKey: `groups.${RibbonGroupId}`;
  /** Lower priority groups collapse first when the Ribbon becomes narrow. */
  readonly priority: number;
}

/** A command's visual placement. The same command can appear on several
 * surfaces without duplicating its build, permission or history contract. */
export type RibbonSurfaceAppearance = 'large' | 'tile' | 'small' | 'split' | 'menu';
export type RibbonSurfaceBreakpoint = 'wide' | 'compact' | 'narrow';

export interface RibbonSurfaceDefinition {
  readonly id: string;
  readonly tab: RibbonCatalogTabId;
  readonly group: RibbonGroupId;
  readonly commandId?: RibbonCommandId;
  readonly controlId?: RibbonControlId;
  readonly order: number;
  readonly appearance: RibbonSurfaceAppearance;
  readonly breakpoints: readonly RibbonSurfaceBreakpoint[];
  readonly overflowTarget?: RibbonGroupId;
  readonly ariaLabel?: string;
}

/** Stateful controls use this same catalog but emit canonical commands. */
export type RibbonControlId =
  | 'format-painter'
  | 'font-family'
  | 'font-size'
  | 'font-increase'
  | 'font-decrease'
  | 'font-color'
  | 'fill-color'
  | 'borders'
  | 'vertical-alignment'
  | 'number-format';

export interface RibbonControlDefinition {
  readonly id: RibbonControlId;
  readonly tab: RibbonCatalogTabId;
  readonly group: RibbonGroupId;
  readonly order: number;
  readonly breakpoints: readonly RibbonSurfaceBreakpoint[];
}

export const RIBBON_TEXT = {
  groups: {
    workbook: 'groups.workbook',
    scripts: 'groups.scripts',
    calculation: 'groups.calculation',
    formulaAudit: 'groups.formulaAudit',
    definedNames: 'groups.definedNames',
    pageSetup: 'groups.pageSetup',
    scaleToFit: 'groups.scaleToFit',
    sheetOptions: 'groups.sheetOptions',
    history: 'groups.history',
    clipboard: 'groups.clipboard',
    font: 'groups.font',
    alignment: 'groups.alignment',
    number: 'groups.number',
    styles: 'groups.styles',
    cells: 'groups.cells',
    insertCells: 'groups.insertCells',
    editing: 'groups.editing',
    tablesPivots: 'groups.tablesPivots',
    chartsVisuals: 'groups.chartsVisuals',
    illustrations: 'groups.illustrations',
    functions: 'groups.functions',
    sortFilter: 'groups.sortFilter',
    dataTools: 'groups.dataTools',
    outline: 'groups.outline',
    findTransform: 'groups.findTransform',
    comments: 'groups.comments',
    notesLinks: 'groups.notesLinks',
    protection: 'groups.protection',
    historyAudit: 'groups.historyAudit',
    freezePanes: 'groups.freezePanes',
    zoom: 'groups.zoom',
    printLayout: 'groups.printLayout',
    appearanceFiles: 'groups.appearanceFiles',
    settings: 'groups.settings',
    pivotAnalyze: 'groups.pivotAnalyze',
    pivotDesign: 'groups.pivotDesign',
  },
  commands: {
    save: 'commands.save',
    exportXlsx: 'commands.exportXlsx',
    importXlsx: 'commands.importXlsx',
    exportXlsxView: 'commands.exportXlsxView',
    importXlsxView: 'commands.importXlsxView',
    openAutomate: 'commands.openAutomate',
    runSampleScript: 'commands.runSampleScript',
    startRecording: 'commands.startRecording',
    stopRecording: 'commands.stopRecording',
    calculateNow: 'commands.calculateNow',
    goalSeek: 'commands.goalSeek',
    tracePrecedents: 'commands.tracePrecedents',
    traceDependents: 'commands.traceDependents',
    removeArrows: 'commands.removeArrows',
    showFormulas: 'commands.showFormulas',
    errorChecking: 'commands.errorChecking',
    evaluateFormula: 'commands.evaluateFormula',
    calculationAutomatic: 'commands.calculationAutomatic',
    calculationManual: 'commands.calculationManual',
    definedNames: 'commands.definedNames',
    pageSetup: 'commands.pageSetup',
    setPrintArea: 'commands.setPrintArea',
    clearPrintArea: 'commands.clearPrintArea',
    printTitleRows: 'commands.printTitleRows',
    printTitleColumns: 'commands.printTitleColumns',
    setScale100: 'commands.setScale100',
    viewGridlines: 'commands.viewGridlines',
    printGridlines: 'commands.printGridlines',
    viewHeadings: 'commands.viewHeadings',
    printHeadings: 'commands.printHeadings',
    undo: 'commands.undo',
    redo: 'commands.redo',
    cut: 'commands.cut',
    copy: 'commands.copy',
    paste: 'commands.paste',
    pasteSpecial: 'commands.pasteSpecial',
    bold: 'commands.bold',
    italic: 'commands.italic',
    underline: 'commands.underline',
    strikethrough: 'commands.strikethrough',
    allBorders: 'commands.allBorders',
    alignLeft: 'commands.alignLeft',
    alignCenter: 'commands.alignCenter',
    alignRight: 'commands.alignRight',
    alignTop: 'commands.alignTop',
    alignMiddle: 'commands.alignMiddle',
    alignBottom: 'commands.alignBottom',
    indentIncrease: 'commands.indentIncrease',
    indentDecrease: 'commands.indentDecrease',
    wrapText: 'commands.wrapText',
    mergeCells: 'commands.mergeCells',
    formatCells: 'commands.formatCells',
    numberFormatGeneral: 'commands.numberFormatGeneral',
    numberFormatCurrency: 'commands.numberFormatCurrency',
    numberFormatPercent: 'commands.numberFormatPercent',
    numberFormatComma: 'commands.numberFormatComma',
    numberFormatDecimal: 'commands.numberFormatDecimal',
    numberFormatDecimalIncrease: 'commands.numberFormatDecimalIncrease',
    numberFormatDecimalDecrease: 'commands.numberFormatDecimalDecrease',
    insertRow: 'commands.insertRow',
    insertColumn: 'commands.insertColumn',
    insertRowHome: 'commands.insertRowHome',
    insertColumnHome: 'commands.insertColumnHome',
    shiftCells: 'commands.shiftCells',
    clearContents: 'commands.clearContents',
    clearFormats: 'commands.clearFormats',
    clearAll: 'commands.clearAll',
    autoSum: 'commands.autoSum',
    fillDown: 'commands.fillDown',
    fillRight: 'commands.fillRight',
    sortRange: 'commands.sortRange',
    conditionalFormat: 'commands.conditionalFormat',
    cellTemplate: 'commands.cellTemplate',
    cellEditor: 'commands.cellEditor',
    pivotTable: 'commands.pivotTable',
    quickPivot: 'commands.quickPivot',
    pivotRefresh: 'commands.pivotRefresh',
    pivotFieldList: 'commands.pivotFieldList',
    chartBuilder: 'commands.chartBuilder',
    columnChart: 'commands.columnChart',
    sparkline: 'commands.sparkline',
    quickSparkline: 'commands.quickSparkline',
    shapesLines: 'commands.shapesLines',
    rectangle: 'commands.rectangle',
    bringDrawingForward: 'commands.bringDrawingForward',
    sendDrawingBackward: 'commands.sendDrawingBackward',
    removeDrawing: 'commands.removeDrawing',
    insertFunction: 'commands.insertFunction',
    deleteRow: 'commands.deleteRow',
    deleteColumn: 'commands.deleteColumn',
    sortAscending: 'commands.sortAscending',
    sortDescending: 'commands.sortDescending',
    customSort: 'commands.customSort',
    dataModel: 'commands.dataModel',
    createDataTable: 'commands.createDataTable',
    formatAsTable: 'commands.formatAsTable',
    totalRow: 'commands.totalRow',
    dataValidation: 'commands.dataValidation',
    filterSelection: 'commands.filterSelection',
    clearFilter: 'commands.clearFilter',
    groupRows: 'commands.groupRows',
    ungroupRows: 'commands.ungroupRows',
    groupColumns: 'commands.groupColumns',
    ungroupColumns: 'commands.ungroupColumns',
    showLevel1: 'commands.showLevel1',
    showLevel2: 'commands.showLevel2',
    showLevel3: 'commands.showLevel3',
    subtotal: 'commands.subtotal',
    removeDuplicates: 'commands.removeDuplicates',
    textToColumns: 'commands.textToColumns',
    findReplace: 'commands.findReplace',
    goTo: 'commands.goTo',
    transpose: 'commands.transpose',
    flipHorizontal: 'commands.flipHorizontal',
    flipVertical: 'commands.flipVertical',
    splitByDelimiter: 'commands.splitByDelimiter',
    newComment: 'commands.newComment',
    resolveComment: 'commands.resolveComment',
    showComments: 'commands.showComments',
    newNote: 'commands.newNote',
    insertLink: 'commands.insertLink',
    protectSelection: 'commands.protectSelection',
    unprotect: 'commands.unprotect',
    revisionLog: 'commands.revisionLog',
    freezeTopRow: 'commands.freezeTopRow',
    freezeFirstColumn: 'commands.freezeFirstColumn',
    freezeAtSelection: 'commands.freezeAtSelection',
    unfreezeAll: 'commands.unfreezeAll',
    zoomIn: 'commands.zoomIn',
    zoomOut: 'commands.zoomOut',
    zoomReset: 'commands.zoomReset',
    printPdf: 'commands.printPdf',
    bandedRows: 'commands.bandedRows',
    settings: 'commands.settings',
    commandPalette: 'commands.commandPalette',
  },
} as const satisfies {
  groups: Record<RibbonGroupId, `groups.${RibbonGroupId}`>;
  commands: Record<RibbonCommandId, `commands.${RibbonCommandId}`>;
};

const group = (
  id: RibbonGroupId,
  tab: RibbonCatalogTabId,
  priority: number,
): RibbonGroupDefinition => ({ id, tab, labelKey: RIBBON_TEXT.groups[id], priority });

export const RIBBON_GROUP_CATALOG: readonly RibbonGroupDefinition[] = [
  group('workbook', 'file', 10),
  group('scripts', 'automate', 70),
  group('calculation', 'formulas', 20),
  group('formulaAudit', 'formulas', 50),
  group('definedNames', 'formulas', 60),
  group('pageSetup', 'pageLayout', 10),
  group('scaleToFit', 'pageLayout', 30),
  group('sheetOptions', 'pageLayout', 50),
  group('history', 'home', 10),
  group('clipboard', 'home', 10),
  group('font', 'home', 20),
  group('alignment', 'home', 30),
  group('number', 'home', 30),
  group('styles', 'home', 35),
  group('cells', 'home', 40),
  group('editing', 'home', 50),
  group('tablesPivots', 'insert', 10),
  group('chartsVisuals', 'insert', 20),
  group('illustrations', 'insert', 40),
  group('functions', 'insert', 60),
  group('insertCells', 'insert', 70),
  group('sortFilter', 'data', 10),
  group('dataTools', 'data', 20),
  group('outline', 'data', 40),
  group('findTransform', 'data', 60),
  group('comments', 'review', 10),
  group('notesLinks', 'review', 20),
  group('protection', 'review', 40),
  group('historyAudit', 'review', 60),
  group('freezePanes', 'view', 10),
  group('zoom', 'view', 20),
  group('printLayout', 'view', 40),
  group('appearanceFiles', 'view', 60),
  group('settings', 'settings', 10),
  group('pivotAnalyze', 'pivotAnalyze', 10),
  group('pivotDesign', 'pivotDesign', 10),
] as const;

const homeSurface = (
  id: string,
  group: RibbonGroupId,
  order: number,
  appearance: RibbonSurfaceAppearance,
  commandId: RibbonCommandId | undefined,
  breakpoints: readonly RibbonSurfaceBreakpoint[] = ['wide', 'compact', 'narrow'],
  overflowTarget?: RibbonGroupId,
): RibbonSurfaceDefinition => ({ id, tab: 'home', group, order, appearance, commandId, breakpoints, overflowTarget });

const homeControl = (
  id: RibbonControlId,
  group: RibbonGroupId,
  order: number,
  breakpoints: readonly RibbonSurfaceBreakpoint[] = ['wide', 'compact', 'narrow'],
): RibbonSurfaceDefinition => ({ id: `control.${id}`, tab: 'home', group, controlId: id, order, appearance: 'small', breakpoints });

/** Single render catalogue for the Home tab. Components must not invent
 * command placements independently from this declaration. */
export const HOME_RIBBON_SURFACES: readonly RibbonSurfaceDefinition[] = [
  homeSurface('history.undo', 'history', 10, 'large', 'undo', ['wide', 'compact']),
  homeSurface('history.redo', 'history', 20, 'large', 'redo', ['wide', 'compact']),
  homeSurface('clipboard.paste', 'clipboard', 10, 'large', 'paste', ['wide', 'compact']),
  homeSurface('clipboard.cut', 'clipboard', 20, 'small', 'cut'),
  homeSurface('clipboard.copy', 'clipboard', 30, 'small', 'copy'),
  homeControl('format-painter', 'clipboard', 40),
  homeSurface('clipboard.paste-special', 'clipboard', 50, 'small', 'pasteSpecial'),
  homeControl('font-family', 'font', 10),
  homeControl('font-size', 'font', 20),
  homeControl('font-increase', 'font', 30),
  homeControl('font-decrease', 'font', 40),
  homeSurface('font.bold', 'font', 50, 'small', 'bold'),
  homeSurface('font.italic', 'font', 60, 'small', 'italic'),
  homeSurface('font.underline', 'font', 70, 'small', 'underline'),
  homeSurface('font.strikethrough', 'font', 80, 'small', 'strikethrough'),
  homeSurface('font.borders', 'font', 85, 'small', 'allBorders'),
  homeControl('borders', 'font', 90),
  homeControl('fill-color', 'font', 100),
  homeControl('vertical-alignment', 'alignment', 10),
  homeSurface('alignment.left', 'alignment', 20, 'small', 'alignLeft'),
  homeSurface('alignment.center', 'alignment', 30, 'small', 'alignCenter'),
  homeSurface('alignment.right', 'alignment', 40, 'small', 'alignRight'),
  homeSurface('alignment.top', 'alignment', 45, 'small', 'alignTop'),
  homeSurface('alignment.middle', 'alignment', 46, 'small', 'alignMiddle'),
  homeSurface('alignment.bottom', 'alignment', 47, 'small', 'alignBottom'),
  homeSurface('alignment.indent-increase', 'alignment', 48, 'small', 'indentIncrease'),
  homeSurface('alignment.indent-decrease', 'alignment', 49, 'small', 'indentDecrease'),
  homeSurface('alignment.wrap', 'alignment', 50, 'small', 'wrapText'),
  homeSurface('alignment.merge', 'alignment', 60, 'split', 'mergeCells'),
  homeControl('number-format', 'number', 10),
  homeSurface('number.currency', 'number', 20, 'small', 'numberFormatCurrency'),
  homeSurface('number.percent', 'number', 30, 'small', 'numberFormatPercent'),
  homeSurface('number.comma', 'number', 40, 'small', 'numberFormatComma'),
  homeSurface('number.decimal', 'number', 50, 'small', 'numberFormatDecimal'),
  homeSurface('number.decimal-increase', 'number', 60, 'small', 'numberFormatDecimalIncrease'),
  homeSurface('number.decimal-decrease', 'number', 70, 'small', 'numberFormatDecimalDecrease'),
  homeSurface('styles.conditional-format', 'styles', 10, 'tile', 'conditionalFormat', ['wide', 'compact']),
  homeSurface('styles.table', 'styles', 20, 'tile', 'formatAsTable', ['wide', 'compact']),
  homeSurface('styles.format-cells', 'styles', 30, 'tile', 'formatCells', ['wide', 'compact']),
  homeSurface('styles.validation', 'styles', 40, 'tile', 'dataValidation', ['wide', 'compact']),
  homeSurface('styles.template', 'styles', 50, 'tile', 'cellTemplate', ['wide', 'compact']),
  homeSurface('styles.editor', 'styles', 60, 'tile', 'cellEditor', ['wide', 'compact']),
  homeSurface('cells.insert', 'cells', 10, 'split', 'insertRowHome'),
  homeSurface('cells.delete', 'cells', 20, 'split', 'deleteRow'),
  homeSurface('cells.format', 'cells', 30, 'split', 'shiftCells'),
  homeSurface('editing.autosum', 'editing', 60, 'small', 'autoSum'),
  homeSurface('editing.fill-down', 'editing', 65, 'small', 'fillDown'),
  homeSurface('editing.fill-right', 'editing', 66, 'small', 'fillRight'),
  homeSurface('editing.sort', 'editing', 70, 'small', 'sortRange'),
  homeSurface('editing.filter', 'editing', 80, 'small', 'filterSelection'),
  homeSurface('editing.clear', 'editing', 90, 'small', 'clearContents'),
  homeSurface('editing.find', 'editing', 100, 'small', 'findReplace'),
] as const;

export function getRibbonSurfaces(
  tab: RibbonCatalogTabId,
  groupId: RibbonGroupId,
  breakpoint: RibbonSurfaceBreakpoint,
): readonly RibbonSurfaceDefinition[] {
  return HOME_RIBBON_SURFACES
    .filter((surface) => surface.tab === tab && surface.group === groupId && surface.breakpoints.includes(breakpoint))
    .sort((left, right) => left.order - right.order);
}

const command = (
  id: RibbonCommandId,
  tab: RibbonCatalogTabId,
  groupId: RibbonGroupId,
  commandId: string,
  labelKey: RibbonTextKey,
  icon: RibbonIconName | undefined,
  params?: unknown,
): CommandDefinition => ({
  id,
  tab,
  group: groupId,
  commandId,
  labelKey,
  icon,
  priority: 30,
  display: icon ? 'small' : 'medium',
  enabled: (context) => !context.canExecute || context.canExecute(commandId, params),
  build: () => ({
    type: 'command',
    descriptor: { commandId, ...(params === undefined ? {} : { params }) },
  }),
});

const callback = (
  id: RibbonCommandId,
  tab: RibbonCatalogTabId,
  groupId: RibbonGroupId,
  labelKey: RibbonTextKey,
  invoke: (context: RibbonCommandContext) => void,
  icon?: RibbonIconName,
): CommandDefinition => ({
  id,
  tab,
  group: groupId,
  labelKey,
  icon,
  priority: 30,
  display: icon ? 'small' : 'medium',
  build: (context) => ({ type: 'callback', invoke: () => invoke(context) }),
});

const intent = (
  id: RibbonCommandId,
  tab: RibbonCatalogTabId,
  groupId: RibbonGroupId,
  labelKey: RibbonTextKey,
  buildIntent: (context: RibbonCommandContext) => UiSessionIntent,
  icon?: RibbonIconName,
): CommandDefinition => ({
  id,
  tab,
  group: groupId,
  labelKey,
  icon,
  priority: 30,
  display: icon ? 'small' : 'medium',
  build: (context) => ({ type: 'intent', intent: buildIntent(context) }),
});

const dynamicCommand = (
  id: RibbonCommandId,
  tab: RibbonCatalogTabId,
  groupId: RibbonGroupId,
  labelKey: RibbonTextKey,
  buildDescriptor: (context: RibbonCommandContext) => CommandDescriptor | undefined,
  icon?: RibbonIconName,
): CommandDefinition => ({
  id,
  tab,
  group: groupId,
  labelKey,
  icon,
  priority: 30,
  display: icon ? 'small' : 'medium',
  build: (context) => {
    const descriptor = buildDescriptor(context);
    return descriptor ? { type: 'command', descriptor } : undefined;
  },
});

const styleCommand = (
  id: RibbonCommandId,
  labelKey: RibbonTextKey,
  icon: RibbonIconName,
  style: (context: RibbonCommandContext) => Record<string, unknown>,
  active?: (context: RibbonCommandContext) => boolean,
  groupId: RibbonGroupId = 'font',
): CommandDefinition => ({
  ...command(id, 'home', groupId, 'sheet.style.set', labelKey, icon),
  build: (context) => ({
    type: 'command',
    descriptor: { commandId: 'sheet.style.set', params: { style: style(context) } },
  }),
  enabled: (context) => (!context.canExecute || context.canExecute('sheet.style.set', { style: style(context) })) && !context.disabled,
  active,
});

/** Structured number-format adjustment used by the Ribbon decimal controls. */
export function adjustRibbonDecimalPlaces(current: string | undefined, delta: number): string {
  const source = current && current !== 'general' ? current : '0';
  const percent = source.includes('%');
  const currency = source.includes('$');
  const comma = source.includes(',');
  const match = source.match(/\.([0#?]+)/);
  const decimals = Math.max(0, Math.min(30, (match?.[1]?.length ?? 0) + delta));
  const integer = currency ? '$#,##0' : comma ? '#,##0' : '0';
  return `${integer}${decimals > 0 ? `.${'0'.repeat(decimals)}` : ''}${percent ? '%' : ''}`;
}

export const RIBBON_COMMAND_CATALOG: readonly CommandDefinition[] = [
  callback('save', 'file', 'workbook', RIBBON_TEXT.commands.save, (context) => context.actions.onSave()),
  callback('exportXlsx', 'file', 'workbook', RIBBON_TEXT.commands.exportXlsx, (context) => context.actions.onExportXlsx()),
  callback('importXlsx', 'file', 'workbook', RIBBON_TEXT.commands.importXlsx, (context) => context.actions.onImportXlsx()),

  intent('openAutomate', 'automate', 'scripts', RIBBON_TEXT.commands.openAutomate, () => ({ type: 'panel.open', panel: 'automate' })),
  dynamicCommand('runSampleScript', 'automate', 'scripts', RIBBON_TEXT.commands.runSampleScript, (context) => ({ commandId: 'automation.run', params: { source: context.sampleAutomationScript } })),
  command('startRecording', 'automate', 'scripts', 'automation.record.start', RIBBON_TEXT.commands.startRecording, undefined, {}),
  command('stopRecording', 'automate', 'scripts', 'automation.record.stop', RIBBON_TEXT.commands.stopRecording, undefined, {}),

  callback('calculateNow', 'formulas', 'calculation', RIBBON_TEXT.commands.calculateNow, (context) => context.actions.onRecalculate(), 'calculator'),
  intent('goalSeek', 'formulas', 'calculation', RIBBON_TEXT.commands.goalSeek, () => ({ type: 'panel.open', panel: 'extended' })),
  callback('tracePrecedents', 'formulas', 'formulaAudit', RIBBON_TEXT.commands.tracePrecedents, (context) => context.actions.onTracePrecedents(), 'search'),
  callback('traceDependents', 'formulas', 'formulaAudit', RIBBON_TEXT.commands.traceDependents, (context) => context.actions.onTraceDependents(), 'share'),
  callback('removeArrows', 'formulas', 'formulaAudit', RIBBON_TEXT.commands.removeArrows, (context) => context.actions.onRemoveArrows(), 'x'),
  callback('showFormulas', 'formulas', 'formulaAudit', RIBBON_TEXT.commands.showFormulas, (context) => context.actions.onToggleShowFormulas(), 'function'),
  callback('errorChecking', 'formulas', 'formulaAudit', RIBBON_TEXT.commands.errorChecking, (context) => context.actions.onScanFormulaErrors(), 'check-circle'),
  callback('evaluateFormula', 'formulas', 'formulaAudit', RIBBON_TEXT.commands.evaluateFormula, (context) => context.actions.onEvaluateFormula(), 'calculator'),
  callback('calculationAutomatic', 'formulas', 'calculation', RIBBON_TEXT.commands.calculationAutomatic, (context) => context.actions.onSetRecalculationMode('automatic'), 'calculator'),
  callback('calculationManual', 'formulas', 'calculation', RIBBON_TEXT.commands.calculationManual, (context) => context.actions.onSetRecalculationMode('manual'), 'calculator'),
  callback('definedNames', 'formulas', 'definedNames', RIBBON_TEXT.commands.definedNames, (context) => context.actions.onOpenDefinedNames(), 'function'),
  callback('pageSetup', 'pageLayout', 'pageSetup', RIBBON_TEXT.commands.pageSetup, (context) => context.actions.onOpenPrintLayout(), 'printer'),
  callback('setPrintArea', 'pageLayout', 'pageSetup', RIBBON_TEXT.commands.setPrintArea, (context) => context.actions.onSetPrintArea(), 'layout'),
  callback('clearPrintArea', 'pageLayout', 'pageSetup', RIBBON_TEXT.commands.clearPrintArea, (context) => context.actions.onClearPrintArea(), 'x'),
  callback('printTitleRows', 'pageLayout', 'pageSetup', RIBBON_TEXT.commands.printTitleRows, (context) => context.actions.onSetPrintTitleRows(), 'rows'),
  callback('printTitleColumns', 'pageLayout', 'pageSetup', RIBBON_TEXT.commands.printTitleColumns, (context) => context.actions.onSetPrintTitleColumns(), 'columns'),
  callback('setScale100', 'pageLayout', 'scaleToFit', RIBBON_TEXT.commands.setScale100, (context) => context.actions.onSetPrintScale(100), 'layout'),
  callback('viewGridlines', 'pageLayout', 'sheetOptions', RIBBON_TEXT.commands.viewGridlines, (context) => context.actions.onToggleViewGridlines(), 'layout'),
  callback('printGridlines', 'pageLayout', 'sheetOptions', RIBBON_TEXT.commands.printGridlines, (context) => context.actions.onTogglePrintGridlines(), 'printer'),
  callback('viewHeadings', 'pageLayout', 'sheetOptions', RIBBON_TEXT.commands.viewHeadings, (context) => context.actions.onToggleViewHeadings(), 'rows'),
  callback('printHeadings', 'pageLayout', 'sheetOptions', RIBBON_TEXT.commands.printHeadings, (context) => context.actions.onTogglePrintHeadings(), 'printer'),

  callback('undo', 'home', 'history', RIBBON_TEXT.commands.undo, (context) => context.actions.onUndo(), 'undo'),
  callback('redo', 'home', 'history', RIBBON_TEXT.commands.redo, (context) => context.actions.onRedo(), 'redo'),
  callback('cut', 'home', 'clipboard', RIBBON_TEXT.commands.cut, (context) => context.actions.onCut(), 'scissors'),
  callback('copy', 'home', 'clipboard', RIBBON_TEXT.commands.copy, (context) => context.actions.onCopy(), 'copy'),
  callback('paste', 'home', 'clipboard', RIBBON_TEXT.commands.paste, (context) => context.actions.onPaste(), 'clipboard'),
  intent('pasteSpecial', 'home', 'clipboard', RIBBON_TEXT.commands.pasteSpecial, () => ({ type: 'dialog.open', dialog: 'paste-special' })),
  styleCommand('bold', RIBBON_TEXT.commands.bold, 'bold', (context) => ({ bold: !context.cellStyle.bold }), (context) => Boolean(context.cellStyle.bold)),
  styleCommand('italic', RIBBON_TEXT.commands.italic, 'italic', (context) => ({ italic: !context.cellStyle.italic }), (context) => Boolean(context.cellStyle.italic)),
  styleCommand('underline', RIBBON_TEXT.commands.underline, 'underline', (context) => ({ underline: !context.cellStyle.underline }), (context) => Boolean(context.cellStyle.underline)),
  styleCommand('strikethrough', RIBBON_TEXT.commands.strikethrough, 'strikethrough', (context) => ({ strikethrough: !context.cellStyle.strikethrough }), (context) => Boolean(context.cellStyle.strikethrough)),
  command('allBorders', 'home', 'font', 'sheet.style.set', RIBBON_TEXT.commands.allBorders, 'borders', { style: { borders: { top: { style: 'thin', color: '#334155' }, right: { style: 'thin', color: '#334155' }, bottom: { style: 'thin', color: '#334155' }, left: { style: 'thin', color: '#334155' } } } }),
  command('alignLeft', 'home', 'alignment', 'sheet.style.set', RIBBON_TEXT.commands.alignLeft, 'align-left', { style: { horizontalAlignment: 'left' } }),
  command('alignCenter', 'home', 'alignment', 'sheet.style.set', RIBBON_TEXT.commands.alignCenter, 'align-center', { style: { horizontalAlignment: 'center' } }),
  command('alignRight', 'home', 'alignment', 'sheet.style.set', RIBBON_TEXT.commands.alignRight, 'align-right', { style: { horizontalAlignment: 'right' } }),
  styleCommand('alignTop', RIBBON_TEXT.commands.alignTop, 'align-left', () => ({ verticalAlignment: 'top' }), (context) => context.cellStyle.verticalAlignment === 'top', 'alignment'),
  styleCommand('alignMiddle', RIBBON_TEXT.commands.alignMiddle, 'align-center', () => ({ verticalAlignment: 'middle' }), (context) => context.cellStyle.verticalAlignment === 'middle', 'alignment'),
  styleCommand('alignBottom', RIBBON_TEXT.commands.alignBottom, 'align-right', () => ({ verticalAlignment: 'bottom' }), (context) => context.cellStyle.verticalAlignment === 'bottom', 'alignment'),
  styleCommand('indentIncrease', RIBBON_TEXT.commands.indentIncrease, 'indent-increase', (context) => ({ indent: Math.min(250, Number(context.cellStyle.indent ?? 0) + 1) }), undefined, 'alignment'),
  styleCommand('indentDecrease', RIBBON_TEXT.commands.indentDecrease, 'indent-decrease', (context) => ({ indent: Math.max(0, Number(context.cellStyle.indent ?? 0) - 1) }), undefined, 'alignment'),
  styleCommand('wrapText', RIBBON_TEXT.commands.wrapText, 'wrap-text', (context) => ({ wrapText: !context.cellStyle.wrapText }), (context) => Boolean(context.cellStyle.wrapText)),
  command('mergeCells', 'home', 'alignment', 'sheet.merge.set', RIBBON_TEXT.commands.mergeCells, 'merge-cells'),
  intent('formatCells', 'home', 'number', RIBBON_TEXT.commands.formatCells, () => ({ type: 'dialog.open', dialog: 'format-cells' }), 'table'),
  command('numberFormatGeneral', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatGeneral, undefined, { style: { numberFormat: 'general' } }),
  command('numberFormatCurrency', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatCurrency, 'dollar-sign', { style: { numberFormat: '$#,##0' } }),
  command('numberFormatPercent', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatPercent, 'percent', { style: { numberFormat: '0%' } }),
  command('numberFormatComma', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatComma, 'comma', { style: { numberFormat: '#,##0' } }),
  command('numberFormatDecimal', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatDecimal, 'decimal-increase', { style: { numberFormat: '0.00' } }),
  styleCommand('numberFormatDecimalIncrease', RIBBON_TEXT.commands.numberFormatDecimalIncrease, 'decimal-increase', (context) => ({ numberFormat: adjustRibbonDecimalPlaces(context.cellStyle.numberFormat, 1) }), undefined, 'number'),
  styleCommand('numberFormatDecimalDecrease', RIBBON_TEXT.commands.numberFormatDecimalDecrease, 'decimal-decrease', (context) => ({ numberFormat: adjustRibbonDecimalPlaces(context.cellStyle.numberFormat, -1) }), undefined, 'number'),
  command('insertRowHome', 'home', 'cells', 'sheet.rows.insert', RIBBON_TEXT.commands.insertRowHome, 'rows', { count: 1 }),
  command('insertColumnHome', 'home', 'cells', 'sheet.columns.insert', RIBBON_TEXT.commands.insertColumnHome, 'columns', { count: 1 }),
  intent('shiftCells', 'home', 'cells', RIBBON_TEXT.commands.shiftCells, () => ({ type: 'dialog.open', dialog: 'shift-cells' })),
  // Delete/Clear Contents must never fall through to the range command's
  // historical default (`all`). The Home entry declares the semantic mode so
  // keyboard, Ribbon and context-menu builders share the same contract.
  command('clearContents', 'home', 'cells', 'sheet.range.clear', RIBBON_TEXT.commands.clearContents, 'trash', { mode: 'contents' }),
  command('clearFormats', 'home', 'cells', 'sheet.range.clear', RIBBON_TEXT.commands.clearFormats, undefined, { mode: 'formats' }),
  command('clearAll', 'home', 'cells', 'sheet.range.clear', RIBBON_TEXT.commands.clearAll, undefined, { mode: 'all' }),
  callback('autoSum', 'home', 'editing', RIBBON_TEXT.commands.autoSum, (context) => context.actions.onAutoSum(), 'calculator'),
  callback('fillDown', 'home', 'editing', RIBBON_TEXT.commands.fillDown, (context) => context.actions.onFill('down'), 'fill-down'),
  callback('fillRight', 'home', 'editing', RIBBON_TEXT.commands.fillRight, (context) => context.actions.onFill('right'), 'fill-right'),
  intent('sortRange', 'home', 'editing', RIBBON_TEXT.commands.sortRange, () => ({ type: 'dialog.open', dialog: 'sort-dialog' }), 'sort'),
  intent('conditionalFormat', 'home', 'editing', RIBBON_TEXT.commands.conditionalFormat, () => ({ type: 'panel.open', panel: 'conditionalFormat' }), 'sparkles'),
  intent('cellTemplate', 'home', 'styles', RIBBON_TEXT.commands.cellTemplate, () => ({ type: 'dialog.open', dialog: 'cell-template' }), 'star'),
  intent('cellEditor', 'home', 'styles', RIBBON_TEXT.commands.cellEditor, () => ({ type: 'dialog.open', dialog: 'cell-editor' }), 'sliders'),

  {
    ...callback('pivotTable', 'insert', 'tablesPivots', RIBBON_TEXT.commands.pivotTable, (context) => context.openCreatePivotDialog?.(), 'table-pivot'),
    enabled: (context) => Boolean(context.openCreatePivotDialog),
  },
  dynamicCommand('quickPivot', 'insert', 'tablesPivots', RIBBON_TEXT.commands.quickPivot, (context) => context.actions.onCreatePivot()),
  {
    id: 'pivotRefresh',
    tab: 'pivotAnalyze',
    group: 'pivotAnalyze',
    labelKey: RIBBON_TEXT.commands.pivotRefresh,
    icon: 'table-pivot',
    priority: 10,
    display: 'medium',
    enabled: (context) => Boolean(context.activePivot)
      && (!context.canExecute || context.canExecute('pivot.refresh', context.activePivot)),
    build: (context) => context.activePivot
      ? { type: 'command', descriptor: { commandId: 'pivot.refresh', params: { sheetId: context.activePivot.sheetId, pivotId: context.activePivot.pivotId } } }
      : undefined,
  },
  {
    ...intent('pivotFieldList', 'pivotDesign', 'pivotDesign', RIBBON_TEXT.commands.pivotFieldList, () => ({ type: 'panel.open', panel: 'pivot' }), 'table-pivot'),
    enabled: (context) => Boolean(context.activePivot),
  },
  intent('chartBuilder', 'insert', 'chartsVisuals', RIBBON_TEXT.commands.chartBuilder, () => ({ type: 'panel.open', panel: 'chart' }), 'chart'),
  dynamicCommand('columnChart', 'insert', 'chartsVisuals', RIBBON_TEXT.commands.columnChart, (context) => context.actions.onCreateChart()),
  intent('sparkline', 'insert', 'chartsVisuals', RIBBON_TEXT.commands.sparkline, () => ({ type: 'panel.open', panel: 'sparkline' }), 'sparkline'),
  dynamicCommand('quickSparkline', 'insert', 'chartsVisuals', RIBBON_TEXT.commands.quickSparkline, (context) => context.actions.onCreateSparkline()),
  intent('shapesLines', 'insert', 'illustrations', RIBBON_TEXT.commands.shapesLines, () => ({ type: 'panel.open', panel: 'shape' }), 'shape-square'),
  dynamicCommand('rectangle', 'insert', 'illustrations', RIBBON_TEXT.commands.rectangle, (context) => context.actions.onCreateShape()),
  dynamicCommand('bringDrawingForward', 'insert', 'illustrations', RIBBON_TEXT.commands.bringDrawingForward, (context) => context.actions.onBringDrawingForward()),
  dynamicCommand('sendDrawingBackward', 'insert', 'illustrations', RIBBON_TEXT.commands.sendDrawingBackward, (context) => context.actions.onSendDrawingBackward()),
  dynamicCommand('removeDrawing', 'insert', 'illustrations', RIBBON_TEXT.commands.removeDrawing, (context) => context.actions.onRemoveDrawing(), 'trash'),
  intent('insertFunction', 'insert', 'functions', RIBBON_TEXT.commands.insertFunction, () => ({ type: 'dialog.open', dialog: 'function-wizard' }), 'function'),
  command('insertRow', 'insert', 'insertCells', 'sheet.rows.insert', RIBBON_TEXT.commands.insertRow, 'rows', { count: 1 }),
  command('insertColumn', 'insert', 'insertCells', 'sheet.columns.insert', RIBBON_TEXT.commands.insertColumn, 'columns', { count: 1 }),
  command('deleteRow', 'insert', 'insertCells', 'sheet.rows.delete', RIBBON_TEXT.commands.deleteRow, 'rows'),
  command('deleteColumn', 'insert', 'insertCells', 'sheet.columns.delete', RIBBON_TEXT.commands.deleteColumn, 'columns'),

  {
    ...dynamicCommand('sortAscending', 'data', 'sortFilter', RIBBON_TEXT.commands.sortAscending, (context) => context.buildSortDescriptor?.(true), 'sort'),
    enabled: (context) => Boolean(context.buildSortDescriptor),
  },
  {
    ...dynamicCommand('sortDescending', 'data', 'sortFilter', RIBBON_TEXT.commands.sortDescending, (context) => context.buildSortDescriptor?.(false), 'sort'),
    enabled: (context) => Boolean(context.buildSortDescriptor),
  },
  intent('customSort', 'data', 'sortFilter', RIBBON_TEXT.commands.customSort, () => ({ type: 'dialog.open', dialog: 'sort-dialog' }), 'sliders'),
  intent('dataModel', 'data', 'dataTools', RIBBON_TEXT.commands.dataModel, () => ({ type: 'panel.open', panel: 'data' }), 'table'),
  callback('createDataTable', 'data', 'dataTools', RIBBON_TEXT.commands.createDataTable, (context) => context.actions.onCreateDataTable(), 'table'),
  callback('formatAsTable', 'data', 'dataTools', RIBBON_TEXT.commands.formatAsTable, (context) => context.actions.onCreateSheetTable(), 'table'),
  dynamicCommand('totalRow', 'data', 'dataTools', RIBBON_TEXT.commands.totalRow, (context) => context.actions.onToggleSheetTableTotalRow(), 'table'),
  intent('dataValidation', 'data', 'dataTools', RIBBON_TEXT.commands.dataValidation, () => ({ type: 'panel.open', panel: 'dataValidation' }), 'check-circle'),
  dynamicCommand('filterSelection', 'data', 'dataTools', RIBBON_TEXT.commands.filterSelection, (context) => context.actions.onApplyFilterSelection(), 'filter'),
  dynamicCommand('clearFilter', 'data', 'dataTools', RIBBON_TEXT.commands.clearFilter, (context) => context.actions.onClearFilter(), 'x'),
  dynamicCommand('groupRows', 'data', 'outline', RIBBON_TEXT.commands.groupRows, (context) => context.actions.onGroupRows()),
  dynamicCommand('ungroupRows', 'data', 'outline', RIBBON_TEXT.commands.ungroupRows, (context) => context.actions.onUngroupRows()),
  dynamicCommand('groupColumns', 'data', 'outline', RIBBON_TEXT.commands.groupColumns, (context) => context.actions.onGroupColumns()),
  dynamicCommand('ungroupColumns', 'data', 'outline', RIBBON_TEXT.commands.ungroupColumns, (context) => context.actions.onUngroupColumns()),
  callback('showLevel1', 'data', 'outline', RIBBON_TEXT.commands.showLevel1, (context) => context.actions.onShowOutlineLevel(1)),
  callback('showLevel2', 'data', 'outline', RIBBON_TEXT.commands.showLevel2, (context) => context.actions.onShowOutlineLevel(2)),
  callback('showLevel3', 'data', 'outline', RIBBON_TEXT.commands.showLevel3, (context) => context.actions.onShowOutlineLevel(3)),
  dynamicCommand('subtotal', 'data', 'outline', RIBBON_TEXT.commands.subtotal, (context) => context.actions.onSubtotal()),
  dynamicCommand('removeDuplicates', 'data', 'outline', RIBBON_TEXT.commands.removeDuplicates, (context) => context.actions.onRemoveDuplicates()),
  dynamicCommand('textToColumns', 'data', 'outline', RIBBON_TEXT.commands.textToColumns, (context) => context.actions.onTextToColumns()),
  intent('findReplace', 'data', 'findTransform', RIBBON_TEXT.commands.findReplace, () => ({ type: 'dialog.open', dialog: 'find-replace' }), 'search'),
  intent('goTo', 'data', 'findTransform', RIBBON_TEXT.commands.goTo, () => ({ type: 'dialog.open', dialog: 'goto' })),
  callback('transpose', 'data', 'findTransform', RIBBON_TEXT.commands.transpose, (context) => context.actions.onTransposeSelection(), 'layout'),
  callback('flipHorizontal', 'data', 'findTransform', RIBBON_TEXT.commands.flipHorizontal, (context) => context.actions.onFlipSelection('h')),
  callback('flipVertical', 'data', 'findTransform', RIBBON_TEXT.commands.flipVertical, (context) => context.actions.onFlipSelection('v')),
  callback('splitByDelimiter', 'data', 'findTransform', RIBBON_TEXT.commands.splitByDelimiter, (context) => context.actions.onSplitByDelimiter()),

  intent('newComment', 'review', 'comments', RIBBON_TEXT.commands.newComment, () => ({ type: 'panel.open', panel: 'inspector', notice: 'Add a comment in the Inspector panel.' }), 'comment'),
  callback('resolveComment', 'review', 'comments', RIBBON_TEXT.commands.resolveComment, (context) => context.actions.onResolveComment(), 'comment'),
  intent('showComments', 'review', 'comments', RIBBON_TEXT.commands.showComments, () => ({ type: 'panel.open', panel: 'inspector' }), 'comment'),
  intent('newNote', 'review', 'notesLinks', RIBBON_TEXT.commands.newNote, () => ({ type: 'panel.open', panel: 'inspector', notice: 'Add a cell note in the Inspector panel.' }), 'comment'),
  intent('insertLink', 'review', 'notesLinks', RIBBON_TEXT.commands.insertLink, () => ({ type: 'panel.open', panel: 'inspector', notice: 'Insert a hyperlink in the Inspector panel.' }), 'share'),
  callback('protectSelection', 'review', 'protection', RIBBON_TEXT.commands.protectSelection, (context) => context.actions.onProtectSelection(), 'lock'),
  callback('unprotect', 'review', 'protection', RIBBON_TEXT.commands.unprotect, (context) => context.actions.onUnprotectSelection(), 'lock'),
  intent('revisionLog', 'review', 'historyAudit', RIBBON_TEXT.commands.revisionLog, () => ({ type: 'panel.open', panel: 'history' }), 'history'),

  command('freezeTopRow', 'view', 'freezePanes', 'sheet.freeze.set', RIBBON_TEXT.commands.freezeTopRow, 'freeze', { pane: { kind: 'frozen', xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0, state: 'frozen' } }),
  command('freezeFirstColumn', 'view', 'freezePanes', 'sheet.freeze.set', RIBBON_TEXT.commands.freezeFirstColumn, 'freeze', { pane: { kind: 'frozen', xSplit: 1, ySplit: 0, startRow: 0, startColumn: 1, state: 'frozen' } }),
  callback('freezeAtSelection', 'view', 'freezePanes', RIBBON_TEXT.commands.freezeAtSelection, (context) => context.actions.onFreezeAtPrimary(), 'freeze'),
  command('unfreezeAll', 'view', 'freezePanes', 'sheet.freeze.set', RIBBON_TEXT.commands.unfreezeAll, 'freeze', { pane: { kind: 'none' } }),
  intent('zoomIn', 'view', 'zoom', RIBBON_TEXT.commands.zoomIn, () => ({ type: 'zoom.adjust', delta: 10 }), 'zoom-in'),
  intent('zoomOut', 'view', 'zoom', RIBBON_TEXT.commands.zoomOut, () => ({ type: 'zoom.adjust', delta: -10 }), 'zoom-out'),
  intent('zoomReset', 'view', 'zoom', RIBBON_TEXT.commands.zoomReset, () => ({ type: 'zoom.set', value: 100 })),
  intent('commandPalette', 'view', 'zoom', RIBBON_TEXT.commands.commandPalette, () => ({ type: 'command-palette.open' }), 'search'),
  intent('printPdf', 'view', 'printLayout', RIBBON_TEXT.commands.printPdf, () => ({ type: 'dialog.open', dialog: 'print-preview' }), 'printer'),
  callback('bandedRows', 'view', 'appearanceFiles', RIBBON_TEXT.commands.bandedRows, (context) => context.actions.onToggleBandedRows()),
  intent('settings', 'settings', 'settings', RIBBON_TEXT.commands.settings, () => ({ type: 'backstage.open', panel: 'options' }), 'sliders'),
  callback('exportXlsxView', 'view', 'appearanceFiles', RIBBON_TEXT.commands.exportXlsxView, (context) => context.actions.onExportXlsx()),
  callback('importXlsxView', 'view', 'appearanceFiles', RIBBON_TEXT.commands.importXlsxView, (context) => context.actions.onImportXlsx()),
] as const;

const commandById = new Map<RibbonCommandId, CommandDefinition>(RIBBON_COMMAND_CATALOG.map((definition) => [definition.id, definition] as const));
const groupById = new Map<RibbonGroupId, RibbonGroupDefinition>(RIBBON_GROUP_CATALOG.map((definition) => [definition.id, definition] as const));

export function getRibbonCommandDefinition(id: RibbonCommandId): CommandDefinition {
  const definition = commandById.get(id);
  if (!definition) throw new Error(`Unknown Ribbon command: ${id}`);
  return definition;
}

export function getRibbonGroupDefinition(id: RibbonGroupId): RibbonGroupDefinition {
  const definition = groupById.get(id);
  if (!definition) throw new Error(`Unknown Ribbon group: ${id}`);
  return definition;
}

export function listRibbonCommands(tab: RibbonCatalogTabId, context: RibbonCommandContext): readonly CommandDefinition[] {
  return RIBBON_COMMAND_CATALOG.filter((definition) => definition.tab === tab && (definition.when?.(context) ?? true));
}

export function isRibbonCommandEnabled(definition: CommandDefinition, context: RibbonCommandContext): boolean {
  if (context.disabled || !(definition.enabled?.(context) ?? true)) return false;
  return definition.build(context) !== undefined;
}

export function buildRibbonCommand(id: RibbonCommandId, context: RibbonCommandContext): RibbonCommandResult | undefined {
  const definition = getRibbonCommandDefinition(id);
  if (!isRibbonCommandEnabled(definition, context)) return undefined;
  return definition.build(context);
}
