import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { AppPhase, SidebarPanelId, UiSessionIntent } from '../types';
import type { BarcodeSymbology, DataChartPlotType, SheetTableModel } from '@react-sheets/core-model';
import { transformNumberFormatPrecision } from '@react-sheets/number-format';

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
  | 'pivotDesign'
  | 'tableSheetDesign'
  | 'ganttTask'
  | 'ganttProject'
  | 'ganttView'
  | 'ganttFormat'
  | 'reportSheetDesign'
  | 'tableDesign'
  | 'chartDesign'
  | 'chartFormat'
  | 'pictureFormat'
  | 'sparklineDesign';

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
  | 'editing'
  | 'illustrations'
  | 'insertSheets'
  | 'insertTables'
  | 'insertCharts'
  | 'insertDataCharts'
  | 'insertLinks'
  | 'insertControls'
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
  | 'pivotDesign'
  | 'tableSheetDesign'
  | 'ganttTask'
  | 'ganttProject'
  | 'ganttView'
  | 'ganttFormat'
  | 'reportSheetDesign'
  | 'tableDesign'
  | 'chartDesign'
  | 'chartFormat'
  | 'pictureFormat'
  | 'sparklineDesign';

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
  | 'borderTop'
  | 'borderBottom'
  | 'borderLeft'
  | 'borderRight'
  | 'borderNone'
  | 'borderOutside'
  | 'borderThickOutside'
  | 'borderInsideHorizontal'
  | 'borderInsideVertical'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'alignTop'
  | 'alignMiddle'
  | 'alignBottom'
  | 'indentIncrease'
  | 'indentDecrease'
  | 'wrapText'
  | 'textOrientation'
  | 'mergeCenter'
  | 'mergeCells'
  | 'mergeAcross'
  | 'unmergeCells'
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
  | 'insertCells'
  | 'deleteCells'
  | 'clearContents'
  | 'clearFormats'
  | 'clearAll'
  | 'clearCommentsNotes'
  | 'clearHyperlinks'
  | 'autoSum'
  | 'fillDown'
  | 'fillUp'
  | 'fillRight'
  | 'fillLeft'
  | 'fillSeries'
  | 'sortRange'
  | 'conditionalFormat'
  | 'cellTemplate'
  | 'cellEditor'
  | 'tableSheet'
  | 'ganttSheet'
  | 'reportSheet'
  | 'worksheetTable'
  | 'dataChart'
  | 'barcode'
  | 'picture'
  | 'camera'
  | 'formControls'
  | 'hyperlink'
  | 'checkbox'
  | 'textbox'
  | 'pivotTable'
  | 'chartBuilder'
  | 'sparkline'
  | 'pictureFormatPanel'
  | 'sparklineDesign'
  | 'shapesLines'
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
  | 'pivotFieldList'
  | 'tableSheetFieldList'
  | 'tableSheetColumnSettings'
  | 'ganttFieldMapping'
  | 'ganttCalendar'
  | 'ganttTimeline'
  | 'ganttDependencyStyle'
  | 'reportFieldBinding'
  | 'reportRenderMode'
  | 'reportPagination'
  | 'reportLayout'
  | 'tableName'
  | 'tableHeaderRow'
  | 'tableTotalRow'
  | 'tableFirstColumn'
  | 'tableLastColumn'
  | 'tableBandedRows'
  | 'tableBandedColumns'
  | 'tableFilterButton'
  | 'tableResize'
  | 'tableConvertToRange'
  | 'tableStyle'
  | 'chartElements'
  | 'chartFormatPanel'
  | 'chartSelectData';

export type RibbonTextKey = `groups.${RibbonGroupId}` | `commands.${RibbonCommandId}`;

export type RibbonDisplay = 'large' | 'medium' | 'small';

export type RibbonIconName =
  | 'align-center'
  | 'align-left'
  | 'align-right'
  | 'align-top'
  | 'align-middle'
  | 'align-bottom'
  | 'barcode'
  | 'borders'
  | 'bold'
  | 'calculator'
  | 'chart'
  | 'chart-column'
  | 'data-chart'
  | 'camera'
  | 'checkbox'
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
  | 'fill-up'
  | 'fill-right'
  | 'fill-left'
  | 'freeze'
  | 'function'
  | 'form-control'
  | 'history'
  | 'italic'
  | 'indent-decrease'
  | 'indent-increase'
  | 'layout'
  | 'link'
  | 'lock'
  | 'merge-cells'
  | 'minimize'
  | 'more-horizontal'
  | 'percent'
  | 'picture'
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
  | 'table-sheet'
  | 'gantt-sheet'
  | 'report-sheet'
  | 'table-pivot'
  | 'trash'
  | 'type'
  | 'textbox'
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
  horizontalAlignment?: 'left' | 'center' | 'right';
  verticalAlignment?: 'top' | 'middle' | 'bottom';
  textRotate?: number;
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
  onMerge: (operation: RibbonMergeOperation) => void;
  onFill: (direction: 'down' | 'up' | 'right' | 'left', mode?: 'copy' | 'series') => void;
  onFreezeAtPrimary: () => void;
  onCreateSheetTable: () => void;
  onOpenTableSettings: () => void;
  onToggleTableOption: (option: 'hasHeaderRow' | 'showFirstColumn' | 'showLastColumn' | 'showBandedRows' | 'showBandedColumns' | 'showFilterButton') => void;
  onConvertActiveTableToRange: () => void;
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
  onCreateAdvancedSheet: (kind: 'table-sheet' | 'gantt-sheet' | 'report-sheet') => void;
  onApplyBarcode: (symbology?: BarcodeSymbology) => void;
  onCreateDataChart: (type?: DataChartPlotType) => void;
  onCreateCamera: () => void;
  onCreateFormControl: () => void;
  onApplyCheckbox: () => void;
  onCreateTextBox: () => void;
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
  activeTableSheet?: { sheetId: string; viewId: string };
  activeGanttSheet?: { sheetId: string; viewId: string };
  activeReportSheet?: { sheetId: string; tableId?: string };
  activeTable?: { sheetId: string; tableId: string; table: SheetTableModel; resizeRange?: SheetTableModel['range'] };
  activeChart?: { sheetId: string; chartId: string };
  activePicture?: { sheetId: string; drawingId: string };
  activeSparkline?: { sheetId: string; sparklineId: string };
  actions: RibbonCommandActions;
  dispatchSessionIntent: (intent: UiSessionIntent) => void;
  sampleAutomationScript: string;
}

export type RibbonMergeOperation = 'center' | 'cells' | 'across' | 'unmerge';

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
export type RibbonSurfaceAppearance = 'large' | 'tile' | 'small' | 'split' | 'menu' | 'gallery' | 'state-control';
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
  /** Optional typed overflow menu owner. Members are rendered through this surface. */
  readonly menuId?: string;
  readonly ariaLabel?: string;
}

/** Stateful controls use this same catalog but emit canonical commands. */
export type RibbonControlId =
  | 'format-painter'
  | 'font-family'
  | 'font-size'
  | 'font-increase'
  | 'font-decrease'
  | 'font-borders-menu'
  | 'font-color'
  | 'fill-color'
  | 'number-format'
  | 'merge-menu'
  | 'cells-insert-menu'
  | 'cells-delete-menu'
  | 'cells-format-menu'
  | 'clear-menu'
  | 'column-width'
  | 'auto-fit-column-width'
  | 'hide-columns'
  | 'unhide-columns'
  | 'default-column-width';

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
    editing: 'groups.editing',
    illustrations: 'groups.illustrations',
    insertSheets: 'groups.insertSheets',
    insertTables: 'groups.insertTables',
    insertCharts: 'groups.insertCharts',
    insertDataCharts: 'groups.insertDataCharts',
    insertLinks: 'groups.insertLinks',
    insertControls: 'groups.insertControls',
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
    tableSheetDesign: 'groups.tableSheetDesign',
    ganttTask: 'groups.ganttTask',
    ganttProject: 'groups.ganttProject',
    ganttView: 'groups.ganttView',
    ganttFormat: 'groups.ganttFormat',
    reportSheetDesign: 'groups.reportSheetDesign',
    tableDesign: 'groups.tableDesign',
    chartDesign: 'groups.chartDesign',
    chartFormat: 'groups.chartFormat',
    pictureFormat: 'groups.pictureFormat',
    sparklineDesign: 'groups.sparklineDesign',
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
    borderTop: 'commands.borderTop',
    borderBottom: 'commands.borderBottom',
    borderLeft: 'commands.borderLeft',
    borderRight: 'commands.borderRight',
    borderNone: 'commands.borderNone',
    borderOutside: 'commands.borderOutside',
    borderThickOutside: 'commands.borderThickOutside',
    borderInsideHorizontal: 'commands.borderInsideHorizontal',
    borderInsideVertical: 'commands.borderInsideVertical',
    alignLeft: 'commands.alignLeft',
    alignCenter: 'commands.alignCenter',
    alignRight: 'commands.alignRight',
    alignTop: 'commands.alignTop',
    alignMiddle: 'commands.alignMiddle',
    alignBottom: 'commands.alignBottom',
    indentIncrease: 'commands.indentIncrease',
    indentDecrease: 'commands.indentDecrease',
    wrapText: 'commands.wrapText',
    textOrientation: 'commands.textOrientation',
    mergeCenter: 'commands.mergeCenter',
    mergeCells: 'commands.mergeCells',
    mergeAcross: 'commands.mergeAcross',
    unmergeCells: 'commands.unmergeCells',
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
    insertCells: 'commands.insertCells',
    deleteCells: 'commands.deleteCells',
    clearContents: 'commands.clearContents',
    clearFormats: 'commands.clearFormats',
    clearAll: 'commands.clearAll',
    clearCommentsNotes: 'commands.clearCommentsNotes',
    clearHyperlinks: 'commands.clearHyperlinks',
    autoSum: 'commands.autoSum',
    fillDown: 'commands.fillDown',
    fillUp: 'commands.fillUp',
    fillRight: 'commands.fillRight',
    fillLeft: 'commands.fillLeft',
    fillSeries: 'commands.fillSeries',
    sortRange: 'commands.sortRange',
    conditionalFormat: 'commands.conditionalFormat',
    cellTemplate: 'commands.cellTemplate',
    cellEditor: 'commands.cellEditor',
    tableSheet: 'commands.tableSheet',
    ganttSheet: 'commands.ganttSheet',
    reportSheet: 'commands.reportSheet',
    worksheetTable: 'commands.worksheetTable',
    dataChart: 'commands.dataChart',
    barcode: 'commands.barcode',
    picture: 'commands.picture',
    camera: 'commands.camera',
    formControls: 'commands.formControls',
    hyperlink: 'commands.hyperlink',
    checkbox: 'commands.checkbox',
    textbox: 'commands.textbox',
    pivotTable: 'commands.pivotTable',
    pivotRefresh: 'commands.pivotRefresh',
    pivotFieldList: 'commands.pivotFieldList',
    tableSheetFieldList: 'commands.tableSheetFieldList',
    tableSheetColumnSettings: 'commands.tableSheetColumnSettings',
    ganttFieldMapping: 'commands.ganttFieldMapping',
    ganttCalendar: 'commands.ganttCalendar',
    ganttTimeline: 'commands.ganttTimeline',
    ganttDependencyStyle: 'commands.ganttDependencyStyle',
    reportFieldBinding: 'commands.reportFieldBinding',
    reportRenderMode: 'commands.reportRenderMode',
    reportPagination: 'commands.reportPagination',
    reportLayout: 'commands.reportLayout',
    tableName: 'commands.tableName',
    tableHeaderRow: 'commands.tableHeaderRow',
    tableTotalRow: 'commands.tableTotalRow',
    tableFirstColumn: 'commands.tableFirstColumn',
    tableLastColumn: 'commands.tableLastColumn',
    tableBandedRows: 'commands.tableBandedRows',
    tableBandedColumns: 'commands.tableBandedColumns',
    tableFilterButton: 'commands.tableFilterButton',
    tableResize: 'commands.tableResize',
    tableConvertToRange: 'commands.tableConvertToRange',
    tableStyle: 'commands.tableStyle',
    chartElements: 'commands.chartElements',
    chartFormatPanel: 'commands.chartFormatPanel',
    chartSelectData: 'commands.chartSelectData',
    chartBuilder: 'commands.chartBuilder',
    sparkline: 'commands.sparkline',
    pictureFormatPanel: 'commands.pictureFormatPanel',
    sparklineDesign: 'commands.sparklineDesign',
    shapesLines: 'commands.shapesLines',
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
  group('insertSheets', 'insert', 10),
  group('insertTables', 'insert', 20),
  group('insertCharts', 'insert', 30),
  group('insertDataCharts', 'insert', 40),
  group('illustrations', 'insert', 50),
  group('insertLinks', 'insert', 60),
  group('insertControls', 'insert', 70),
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
  group('tableSheetDesign', 'tableSheetDesign', 10),
  group('ganttTask', 'ganttTask', 10),
  group('ganttProject', 'ganttProject', 20),
  group('ganttView', 'ganttView', 30),
  group('ganttFormat', 'ganttFormat', 40),
  group('reportSheetDesign', 'reportSheetDesign', 10),
  group('tableDesign', 'tableDesign', 10),
  group('chartDesign', 'chartDesign', 10),
  group('chartFormat', 'chartFormat', 20),
  group('pictureFormat', 'pictureFormat', 10),
  group('sparklineDesign', 'sparklineDesign', 10),
] as const;

const ribbonSurface = (
  tab: RibbonCatalogTabId,
  id: string,
  group: RibbonGroupId,
  order: number,
  appearance: RibbonSurfaceAppearance,
  commandId: RibbonCommandId | undefined,
  breakpoints: readonly RibbonSurfaceBreakpoint[] = ['wide', 'compact', 'narrow'],
  overflowTarget?: RibbonGroupId,
  menuId?: string,
): RibbonSurfaceDefinition => ({ id, tab, group, order, appearance, commandId, breakpoints, overflowTarget, menuId });

const homeControl = (
  id: RibbonControlId,
  group: RibbonGroupId,
  order: number,
  breakpoints: readonly RibbonSurfaceBreakpoint[] = ['wide', 'compact', 'narrow'],
  menuId?: string,
): RibbonSurfaceDefinition => ({ id: `control.${id}`, tab: 'home', group, controlId: id, order, appearance: 'small', breakpoints, menuId });

/** Single render catalogue for the Home tab. Components must not invent
 * command placements independently from this declaration. */
export const HOME_RIBBON_SURFACES: readonly RibbonSurfaceDefinition[] = [
  ribbonSurface('home', 'history.undo', 'history', 10, 'large', 'undo'),
  ribbonSurface('home', 'history.redo', 'history', 20, 'large', 'redo'),
  ribbonSurface('home', 'clipboard.paste', 'clipboard', 10, 'large', 'paste'),
  ribbonSurface('home', 'clipboard.cut', 'clipboard', 20, 'small', 'cut'),
  ribbonSurface('home', 'clipboard.copy', 'clipboard', 30, 'small', 'copy'),
  homeControl('format-painter', 'clipboard', 40),
  ribbonSurface('home', 'clipboard.paste-special', 'clipboard', 50, 'small', 'pasteSpecial'),
  homeControl('font-family', 'font', 10),
  homeControl('font-size', 'font', 20),
  homeControl('font-increase', 'font', 30),
  homeControl('font-decrease', 'font', 40),
  ribbonSurface('home', 'font.bold', 'font', 50, 'small', 'bold'),
  ribbonSurface('home', 'font.italic', 'font', 60, 'small', 'italic'),
  ribbonSurface('home', 'font.underline', 'font', 70, 'small', 'underline'),
  ribbonSurface('home', 'font.strikethrough', 'font', 80, 'small', 'strikethrough'),
  homeControl('font-borders-menu', 'font', 85),
  ribbonSurface('home', 'font.borders.all', 'font', 851, 'menu', 'allBorders', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.top', 'font', 852, 'menu', 'borderTop', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.bottom', 'font', 853, 'menu', 'borderBottom', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.left', 'font', 854, 'menu', 'borderLeft', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.right', 'font', 855, 'menu', 'borderRight', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.outside', 'font', 856, 'menu', 'borderOutside', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.thick-outside', 'font', 857, 'menu', 'borderThickOutside', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.inside-horizontal', 'font', 858, 'menu', 'borderInsideHorizontal', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.inside-vertical', 'font', 859, 'menu', 'borderInsideVertical', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  ribbonSurface('home', 'font.borders.none', 'font', 860, 'menu', 'borderNone', ['wide', 'compact', 'narrow'], undefined, 'control.font-borders-menu'),
  homeControl('font-color', 'font', 90),
  homeControl('fill-color', 'font', 100),
  ribbonSurface('home', 'alignment.left', 'alignment', 20, 'small', 'alignLeft'),
  ribbonSurface('home', 'alignment.center', 'alignment', 30, 'small', 'alignCenter'),
  ribbonSurface('home', 'alignment.right', 'alignment', 40, 'small', 'alignRight'),
  ribbonSurface('home', 'alignment.top', 'alignment', 45, 'small', 'alignTop'),
  ribbonSurface('home', 'alignment.middle', 'alignment', 46, 'small', 'alignMiddle'),
  ribbonSurface('home', 'alignment.bottom', 'alignment', 47, 'small', 'alignBottom'),
  ribbonSurface('home', 'alignment.indent-increase', 'alignment', 48, 'small', 'indentIncrease'),
  ribbonSurface('home', 'alignment.indent-decrease', 'alignment', 49, 'small', 'indentDecrease'),
  ribbonSurface('home', 'alignment.wrap', 'alignment', 50, 'small', 'wrapText'),
  ribbonSurface('home', 'alignment.orientation', 'alignment', 55, 'small', 'textOrientation'),
  homeControl('merge-menu', 'alignment', 60),
  ribbonSurface('home', 'alignment.merge-center', 'alignment', 61, 'menu', 'mergeCenter', ['wide', 'compact', 'narrow'], undefined, 'control.merge-menu'),
  ribbonSurface('home', 'alignment.merge-cells', 'alignment', 62, 'menu', 'mergeCells', ['wide', 'compact', 'narrow'], undefined, 'control.merge-menu'),
  ribbonSurface('home', 'alignment.merge-across', 'alignment', 63, 'menu', 'mergeAcross', ['wide', 'compact', 'narrow'], undefined, 'control.merge-menu'),
  ribbonSurface('home', 'alignment.unmerge', 'alignment', 64, 'menu', 'unmergeCells', ['wide', 'compact', 'narrow'], undefined, 'control.merge-menu'),
  homeControl('number-format', 'number', 10),
  ribbonSurface('home', 'number.currency', 'number', 20, 'small', 'numberFormatCurrency'),
  ribbonSurface('home', 'number.percent', 'number', 30, 'small', 'numberFormatPercent'),
  ribbonSurface('home', 'number.comma', 'number', 40, 'small', 'numberFormatComma'),
  ribbonSurface('home', 'number.decimal', 'number', 50, 'small', 'numberFormatDecimal'),
  ribbonSurface('home', 'number.decimal-increase', 'number', 60, 'small', 'numberFormatDecimalIncrease'),
  ribbonSurface('home', 'number.decimal-decrease', 'number', 70, 'small', 'numberFormatDecimalDecrease'),
  ribbonSurface('home', 'styles.conditional-format', 'styles', 10, 'tile', 'conditionalFormat'),
  ribbonSurface('home', 'styles.table', 'styles', 20, 'tile', 'formatAsTable'),
  ribbonSurface('home', 'styles.format-cells', 'styles', 30, 'tile', 'formatCells'),
  ribbonSurface('home', 'styles.validation', 'styles', 40, 'tile', 'dataValidation'),
  ribbonSurface('home', 'styles.template', 'styles', 50, 'tile', 'cellTemplate'),
  ribbonSurface('home', 'styles.editor', 'styles', 60, 'tile', 'cellEditor'),
  homeControl('cells-insert-menu', 'cells', 10),
  homeControl('cells-delete-menu', 'cells', 20),
  homeControl('cells-format-menu', 'cells', 30),
  ribbonSurface('home', 'cells.insert-row', 'cells', 40, 'menu', 'insertRowHome', ['wide', 'compact', 'narrow'], undefined, 'control.cells-insert-menu'),
  ribbonSurface('home', 'cells.insert-column', 'cells', 41, 'menu', 'insertColumnHome', ['wide', 'compact', 'narrow'], undefined, 'control.cells-insert-menu'),
  ribbonSurface('home', 'cells.insert-cells', 'cells', 42, 'menu', 'insertCells', ['wide', 'compact', 'narrow'], undefined, 'control.cells-insert-menu'),
  ribbonSurface('home', 'cells.delete-row', 'cells', 50, 'menu', 'deleteRow', ['wide', 'compact', 'narrow'], undefined, 'control.cells-delete-menu'),
  ribbonSurface('home', 'cells.delete-column', 'cells', 51, 'menu', 'deleteColumn', ['wide', 'compact', 'narrow'], undefined, 'control.cells-delete-menu'),
  ribbonSurface('home', 'cells.delete-cells', 'cells', 52, 'menu', 'deleteCells', ['wide', 'compact', 'narrow'], undefined, 'control.cells-delete-menu'),
  homeControl('column-width', 'cells', 60, ['wide', 'compact', 'narrow'], 'control.cells-format-menu'),
  homeControl('auto-fit-column-width', 'cells', 61, ['wide', 'compact', 'narrow'], 'control.cells-format-menu'),
  homeControl('hide-columns', 'cells', 62, ['wide', 'compact', 'narrow'], 'control.cells-format-menu'),
  homeControl('unhide-columns', 'cells', 63, ['wide', 'compact', 'narrow'], 'control.cells-format-menu'),
  homeControl('default-column-width', 'cells', 64, ['wide', 'compact', 'narrow'], 'control.cells-format-menu'),
  ribbonSurface('home', 'editing.autosum', 'editing', 60, 'small', 'autoSum'),
  ribbonSurface('home', 'editing.fill-down', 'editing', 65, 'small', 'fillDown'),
  ribbonSurface('home', 'editing.fill-up', 'editing', 66, 'small', 'fillUp'),
  ribbonSurface('home', 'editing.fill-right', 'editing', 67, 'small', 'fillRight'),
  ribbonSurface('home', 'editing.fill-left', 'editing', 68, 'small', 'fillLeft'),
  ribbonSurface('home', 'editing.fill-series', 'editing', 69, 'small', 'fillSeries'),
  ribbonSurface('home', 'editing.sort', 'editing', 70, 'small', 'sortRange'),
  ribbonSurface('home', 'editing.filter', 'editing', 80, 'small', 'filterSelection'),
  homeControl('clear-menu', 'editing', 90),
  ribbonSurface('home', 'editing.clear-contents', 'editing', 91, 'menu', 'clearContents', ['wide', 'compact', 'narrow'], undefined, 'control.clear-menu'),
  ribbonSurface('home', 'editing.clear-formats', 'editing', 92, 'menu', 'clearFormats', ['wide', 'compact', 'narrow'], undefined, 'control.clear-menu'),
  ribbonSurface('home', 'editing.clear-all', 'editing', 93, 'menu', 'clearAll', ['wide', 'compact', 'narrow'], undefined, 'control.clear-menu'),
  ribbonSurface('home', 'editing.clear-comments-notes', 'editing', 94, 'menu', 'clearCommentsNotes', ['wide', 'compact', 'narrow'], undefined, 'control.clear-menu'),
  ribbonSurface('home', 'editing.clear-hyperlinks', 'editing', 95, 'menu', 'clearHyperlinks', ['wide', 'compact', 'narrow'], undefined, 'control.clear-menu'),
  ribbonSurface('home', 'editing.find', 'editing', 100, 'small', 'findReplace'),
] as const;

export const INSERT_RIBBON_SURFACES: readonly RibbonSurfaceDefinition[] = [
  ribbonSurface('insert', 'sheets.table-sheet', 'insertSheets', 10, 'large', 'tableSheet'),
  ribbonSurface('insert', 'sheets.gantt-sheet', 'insertSheets', 20, 'large', 'ganttSheet'),
  ribbonSurface('insert', 'sheets.report-sheet', 'insertSheets', 30, 'large', 'reportSheet'),
  ribbonSurface('insert', 'tables.worksheet-table', 'insertTables', 10, 'split', 'worksheetTable'),
  ribbonSurface('insert', 'tables.pivot', 'insertTables', 20, 'large', 'pivotTable'),
  ribbonSurface('insert', 'charts.gallery', 'insertCharts', 10, 'gallery', 'chartBuilder'),
  ribbonSurface('insert', 'charts.barcode', 'insertCharts', 20, 'large', 'barcode'),
  ribbonSurface('insert', 'charts.sparkline', 'insertCharts', 30, 'split', 'sparkline'),
  ribbonSurface('insert', 'data-charts.insert', 'insertDataCharts', 10, 'large', 'dataChart'),
  ribbonSurface('insert', 'illustrations.picture', 'illustrations', 10, 'split', 'picture'),
  ribbonSurface('insert', 'illustrations.shape', 'illustrations', 20, 'split', 'shapesLines'),
  ribbonSurface('insert', 'illustrations.camera', 'illustrations', 30, 'large', 'camera'),
  ribbonSurface('insert', 'illustrations.controls', 'illustrations', 40, 'gallery', 'formControls'),
  ribbonSurface('insert', 'links.hyperlink', 'insertLinks', 10, 'large', 'hyperlink'),
  ribbonSurface('insert', 'controls.checkbox', 'insertControls', 10, 'large', 'checkbox'),
  ribbonSurface('insert', 'controls.textbox', 'insertControls', 20, 'large', 'textbox'),
] as const;

export const RIBBON_TAB_SURFACES: readonly RibbonSurfaceDefinition[] = [...HOME_RIBBON_SURFACES, ...INSERT_RIBBON_SURFACES];

export function getRibbonSurfaces(
  tab: RibbonCatalogTabId,
  groupId: RibbonGroupId,
  breakpoint: RibbonSurfaceBreakpoint,
): readonly RibbonSurfaceDefinition[] {
  return RIBBON_TAB_SURFACES
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

const precisionStyleCommand = (
  id: RibbonCommandId,
  labelKey: RibbonTextKey,
  icon: RibbonIconName,
  delta: 1 | -1,
): CommandDefinition => ({
  ...command(id, 'home', 'number', 'sheet.style.set', labelKey, icon),
  build: (context) => {
    const result = transformNumberFormatPrecision(context.cellStyle.numberFormat, delta);
    if (!result.ok) return undefined;
    return { type: 'command', descriptor: { commandId: 'sheet.style.set', params: { style: { numberFormat: result.format } } } };
  },
  enabled: (context) => {
    const result = transformNumberFormatPrecision(context.cellStyle.numberFormat, delta);
    return result.ok && !context.disabled && (!context.canExecute || context.canExecute('sheet.style.set', { style: { numberFormat: result.format } }));
  },
});

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
  command('allBorders', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.allBorders, 'borders', { placement: 'all', line: { style: 'thin', color: '#334155' } }),
  command('borderTop', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderTop, undefined, { placement: 'top', line: { style: 'thin', color: '#334155' } }),
  command('borderBottom', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderBottom, undefined, { placement: 'bottom', line: { style: 'thin', color: '#334155' } }),
  command('borderLeft', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderLeft, undefined, { placement: 'left', line: { style: 'thin', color: '#334155' } }),
  command('borderRight', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderRight, undefined, { placement: 'right', line: { style: 'thin', color: '#334155' } }),
  command('borderNone', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderNone, undefined, { placement: 'none' }),
  command('borderOutside', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderOutside, undefined, { placement: 'outside', line: { style: 'thin', color: '#334155' } }),
  command('borderThickOutside', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderThickOutside, undefined, { placement: 'thick-outside', line: { style: 'thick', color: '#334155' } }),
  command('borderInsideHorizontal', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderInsideHorizontal, undefined, { placement: 'inside-horizontal', line: { style: 'thin', color: '#334155' } }),
  command('borderInsideVertical', 'home', 'font', 'sheet.borders.set', RIBBON_TEXT.commands.borderInsideVertical, undefined, { placement: 'inside-vertical', line: { style: 'thin', color: '#334155' } }),
  styleCommand('alignLeft', RIBBON_TEXT.commands.alignLeft, 'align-left', () => ({ horizontalAlignment: 'left' }), (context) => context.cellStyle.horizontalAlignment === 'left', 'alignment'),
  styleCommand('alignCenter', RIBBON_TEXT.commands.alignCenter, 'align-center', () => ({ horizontalAlignment: 'center' }), (context) => context.cellStyle.horizontalAlignment === 'center', 'alignment'),
  styleCommand('alignRight', RIBBON_TEXT.commands.alignRight, 'align-right', () => ({ horizontalAlignment: 'right' }), (context) => context.cellStyle.horizontalAlignment === 'right', 'alignment'),
  styleCommand('alignTop', RIBBON_TEXT.commands.alignTop, 'align-top', () => ({ verticalAlignment: 'top' }), (context) => context.cellStyle.verticalAlignment === 'top', 'alignment'),
  styleCommand('alignMiddle', RIBBON_TEXT.commands.alignMiddle, 'align-middle', () => ({ verticalAlignment: 'middle' }), (context) => context.cellStyle.verticalAlignment === 'middle', 'alignment'),
  styleCommand('alignBottom', RIBBON_TEXT.commands.alignBottom, 'align-bottom', () => ({ verticalAlignment: 'bottom' }), (context) => context.cellStyle.verticalAlignment === 'bottom', 'alignment'),
  styleCommand('indentIncrease', RIBBON_TEXT.commands.indentIncrease, 'indent-increase', (context) => ({ indent: Math.min(250, Number(context.cellStyle.indent ?? 0) + 1) }), undefined, 'alignment'),
  styleCommand('indentDecrease', RIBBON_TEXT.commands.indentDecrease, 'indent-decrease', (context) => ({ indent: Math.max(0, Number(context.cellStyle.indent ?? 0) - 1) }), undefined, 'alignment'),
  styleCommand('wrapText', RIBBON_TEXT.commands.wrapText, 'wrap-text', (context) => ({ wrapText: !context.cellStyle.wrapText }), (context) => Boolean(context.cellStyle.wrapText)),
  styleCommand('textOrientation', RIBBON_TEXT.commands.textOrientation, 'type', (context) => ({ textRotate: context.cellStyle.textRotate === 45 ? 0 : 45 }), (context) => Boolean(context.cellStyle.textRotate), 'alignment'),
  callback('mergeCenter', 'home', 'alignment', RIBBON_TEXT.commands.mergeCenter, (context) => context.actions.onMerge('center'), 'merge-cells'),
  callback('mergeCells', 'home', 'alignment', RIBBON_TEXT.commands.mergeCells, (context) => context.actions.onMerge('cells'), 'merge-cells'),
  callback('mergeAcross', 'home', 'alignment', RIBBON_TEXT.commands.mergeAcross, (context) => context.actions.onMerge('across'), 'merge-cells'),
  callback('unmergeCells', 'home', 'alignment', RIBBON_TEXT.commands.unmergeCells, (context) => context.actions.onMerge('unmerge'), 'merge-cells'),
  intent('formatCells', 'home', 'number', RIBBON_TEXT.commands.formatCells, () => ({ type: 'dialog.open', dialog: 'format-cells' }), 'table'),
  command('numberFormatGeneral', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatGeneral, undefined, { style: { numberFormat: 'general' } }),
  command('numberFormatCurrency', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatCurrency, 'dollar-sign', { style: { numberFormat: '$#,##0' } }),
  command('numberFormatPercent', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatPercent, 'percent', { style: { numberFormat: '0%' } }),
  command('numberFormatComma', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatComma, 'comma', { style: { numberFormat: '#,##0' } }),
  command('numberFormatDecimal', 'home', 'number', 'sheet.style.set', RIBBON_TEXT.commands.numberFormatDecimal, 'decimal-increase', { style: { numberFormat: '0.00' } }),
  precisionStyleCommand('numberFormatDecimalIncrease', RIBBON_TEXT.commands.numberFormatDecimalIncrease, 'decimal-increase', 1),
  precisionStyleCommand('numberFormatDecimalDecrease', RIBBON_TEXT.commands.numberFormatDecimalDecrease, 'decimal-decrease', -1),
  command('insertRowHome', 'home', 'cells', 'sheet.rows.insert', RIBBON_TEXT.commands.insertRowHome, 'rows', { count: 1 }),
  command('insertColumnHome', 'home', 'cells', 'sheet.columns.insert', RIBBON_TEXT.commands.insertColumnHome, 'columns', { count: 1 }),
  command('deleteRow', 'home', 'cells', 'sheet.rows.delete', RIBBON_TEXT.commands.deleteRow, 'rows', { count: 1 }),
  command('deleteColumn', 'home', 'cells', 'sheet.columns.delete', RIBBON_TEXT.commands.deleteColumn, 'columns', { count: 1 }),
  intent('insertCells', 'home', 'cells', RIBBON_TEXT.commands.insertCells, () => ({ type: 'dialog.open', dialog: 'shift-cells', operation: 'insert' })),
  intent('deleteCells', 'home', 'cells', RIBBON_TEXT.commands.deleteCells, () => ({ type: 'dialog.open', dialog: 'shift-cells', operation: 'delete' })),
  // Delete/Clear Contents must never fall through to the range command's
  // historical default (`all`). The Home entry declares the semantic family so
  // keyboard, Ribbon and context-menu builders share the same contract.
  command('clearContents', 'home', 'editing', 'sheet.range.clear', RIBBON_TEXT.commands.clearContents, 'trash', { family: 'contents' }),
  command('clearFormats', 'home', 'editing', 'sheet.range.clear', RIBBON_TEXT.commands.clearFormats, undefined, { family: 'formats' }),
  command('clearAll', 'home', 'editing', 'sheet.range.clear', RIBBON_TEXT.commands.clearAll, undefined, { family: 'all' }),
  command('clearCommentsNotes', 'home', 'editing', 'sheet.range.clear', RIBBON_TEXT.commands.clearCommentsNotes, undefined, { family: 'comments-and-notes' }),
  command('clearHyperlinks', 'home', 'editing', 'sheet.range.clear', RIBBON_TEXT.commands.clearHyperlinks, undefined, { family: 'hyperlinks' }),
  callback('autoSum', 'home', 'editing', RIBBON_TEXT.commands.autoSum, (context) => context.actions.onAutoSum(), 'calculator'),
  callback('fillDown', 'home', 'editing', RIBBON_TEXT.commands.fillDown, (context) => context.actions.onFill('down'), 'fill-down'),
  callback('fillUp', 'home', 'editing', RIBBON_TEXT.commands.fillUp, (context) => context.actions.onFill('up'), 'fill-up'),
  callback('fillRight', 'home', 'editing', RIBBON_TEXT.commands.fillRight, (context) => context.actions.onFill('right'), 'fill-right'),
  callback('fillLeft', 'home', 'editing', RIBBON_TEXT.commands.fillLeft, (context) => context.actions.onFill('left'), 'fill-left'),
  callback('fillSeries', 'home', 'editing', RIBBON_TEXT.commands.fillSeries, (context) => context.actions.onFill('down', 'series'), 'sort'),
  intent('sortRange', 'home', 'editing', RIBBON_TEXT.commands.sortRange, () => ({ type: 'dialog.open', dialog: 'sort-dialog' }), 'sort'),
  intent('conditionalFormat', 'home', 'editing', RIBBON_TEXT.commands.conditionalFormat, () => ({ type: 'panel.open', panel: 'conditionalFormat' }), 'sparkles'),
  intent('cellTemplate', 'home', 'styles', RIBBON_TEXT.commands.cellTemplate, () => ({ type: 'dialog.open', dialog: 'cell-template' }), 'star'),
  intent('cellEditor', 'home', 'styles', RIBBON_TEXT.commands.cellEditor, () => ({ type: 'dialog.open', dialog: 'cell-editor' }), 'sliders'),

  callback('tableSheet', 'insert', 'insertSheets', RIBBON_TEXT.commands.tableSheet, (context) => context.actions.onCreateAdvancedSheet('table-sheet'), 'table-sheet'),
  callback('ganttSheet', 'insert', 'insertSheets', RIBBON_TEXT.commands.ganttSheet, (context) => context.actions.onCreateAdvancedSheet('gantt-sheet'), 'gantt-sheet'),
  callback('reportSheet', 'insert', 'insertSheets', RIBBON_TEXT.commands.reportSheet, (context) => context.actions.onCreateAdvancedSheet('report-sheet'), 'report-sheet'),
  callback('worksheetTable', 'insert', 'insertTables', RIBBON_TEXT.commands.worksheetTable, (context) => context.actions.onCreateSheetTable(), 'table'),
  callback('barcode', 'insert', 'insertCharts', RIBBON_TEXT.commands.barcode, (context) => context.actions.onApplyBarcode(), 'barcode'),
  callback('dataChart', 'insert', 'insertDataCharts', RIBBON_TEXT.commands.dataChart, (context) => context.actions.onCreateDataChart(), 'data-chart'),
  intent('picture', 'insert', 'illustrations', RIBBON_TEXT.commands.picture, () => ({ type: 'dialog.open', dialog: 'insert-picture' }), 'picture'),
  callback('camera', 'insert', 'illustrations', RIBBON_TEXT.commands.camera, (context) => context.actions.onCreateCamera(), 'camera'),
  callback('formControls', 'insert', 'illustrations', RIBBON_TEXT.commands.formControls, (context) => context.actions.onCreateFormControl(), 'form-control'),
  intent('hyperlink', 'insert', 'insertLinks', RIBBON_TEXT.commands.hyperlink, () => ({ type: 'panel.open', panel: 'inspector', notice: '在检查器中设置超链接。' }), 'link'),
  callback('checkbox', 'insert', 'insertControls', RIBBON_TEXT.commands.checkbox, (context) => context.actions.onApplyCheckbox(), 'checkbox'),
  callback('textbox', 'insert', 'insertControls', RIBBON_TEXT.commands.textbox, (context) => context.actions.onCreateTextBox(), 'textbox'),

  {
    ...callback('pivotTable', 'insert', 'insertTables', RIBBON_TEXT.commands.pivotTable, (context) => context.openCreatePivotDialog?.(), 'table-pivot'),
    enabled: (context) => Boolean(context.openCreatePivotDialog),
  },
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
  {
    ...intent('tableSheetFieldList', 'tableSheetDesign', 'tableSheetDesign', RIBBON_TEXT.commands.tableSheetFieldList, () => ({ type: 'panel.open', panel: 'data' }), 'table'),
    enabled: (context) => Boolean(context.activeTableSheet),
  },
  {
    ...intent('tableSheetColumnSettings', 'tableSheetDesign', 'tableSheetDesign', RIBBON_TEXT.commands.tableSheetColumnSettings, () => ({ type: 'panel.open', panel: 'data' }), 'sliders'),
    enabled: (context) => Boolean(context.activeTableSheet),
  },
  {
    ...intent('ganttFieldMapping', 'ganttTask', 'ganttTask', RIBBON_TEXT.commands.ganttFieldMapping, () => ({ type: 'panel.open', panel: 'data' }), 'table'),
    enabled: (context) => Boolean(context.activeGanttSheet),
  },
  {
    ...intent('ganttCalendar', 'ganttProject', 'ganttProject', RIBBON_TEXT.commands.ganttCalendar, () => ({ type: 'panel.open', panel: 'data' }), 'sliders'),
    enabled: (context) => Boolean(context.activeGanttSheet),
  },
  {
    ...intent('ganttTimeline', 'ganttView', 'ganttView', RIBBON_TEXT.commands.ganttTimeline, () => ({ type: 'panel.open', panel: 'data' }), 'layout'),
    enabled: (context) => Boolean(context.activeGanttSheet),
  },
  {
    ...intent('ganttDependencyStyle', 'ganttFormat', 'ganttFormat', RIBBON_TEXT.commands.ganttDependencyStyle, () => ({ type: 'panel.open', panel: 'data' }), 'sparkles'),
    enabled: (context) => Boolean(context.activeGanttSheet),
  },
  {
    ...intent('reportFieldBinding', 'reportSheetDesign', 'reportSheetDesign', RIBBON_TEXT.commands.reportFieldBinding, () => ({ type: 'panel.open', panel: 'data' }), 'table'),
    enabled: (context) => Boolean(context.activeReportSheet),
  },
  {
    ...intent('reportRenderMode', 'reportSheetDesign', 'reportSheetDesign', RIBBON_TEXT.commands.reportRenderMode, () => ({ type: 'panel.open', panel: 'data' }), 'layout'),
    enabled: (context) => Boolean(context.activeReportSheet),
  },
  {
    ...intent('reportPagination', 'reportSheetDesign', 'reportSheetDesign', RIBBON_TEXT.commands.reportPagination, () => ({ type: 'panel.open', panel: 'data' }), 'sliders'),
    enabled: (context) => Boolean(context.activeReportSheet),
  },
  {
    ...intent('reportLayout', 'reportSheetDesign', 'reportSheetDesign', RIBBON_TEXT.commands.reportLayout, () => ({ type: 'panel.open', panel: 'data' }), 'layout'),
    enabled: (context) => Boolean(context.activeReportSheet),
  },
  {
    ...intent('tableName', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableName, () => ({ type: 'panel.open', panel: 'data' }), 'table'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableHeaderRow', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableHeaderRow, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), hasHeaderRow: !context.activeTable.table.hasHeaderRow, showFilterButton: !context.activeTable.table.hasHeaderRow ? context.activeTable.table.showFilterButton : false, autoFilter: undefined } }
      : undefined, 'rows'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableTotalRow', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableTotalRow, (context) => context.activeTable
      ? { commandId: 'sheetTable.toggleTotalRow', params: { sheetId: context.activeTable.sheetId, tableId: context.activeTable.tableId, enabled: !context.activeTable.table.hasTotalRow } }
      : undefined, 'rows'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableFirstColumn', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableFirstColumn, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), showFirstColumn: !context.activeTable.table.showFirstColumn } }
      : undefined, 'columns'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableLastColumn', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableLastColumn, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), showLastColumn: !context.activeTable.table.showLastColumn } }
      : undefined, 'columns'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableBandedRows', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableBandedRows, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), showBandedRows: !context.activeTable.table.showBandedRows } }
      : undefined, 'rows'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableBandedColumns', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableBandedColumns, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), showBandedColumns: !context.activeTable.table.showBandedColumns } }
      : undefined, 'columns'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableFilterButton', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableFilterButton, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), showFilterButton: !context.activeTable.table.showFilterButton, autoFilter: undefined } }
      : undefined, 'filter'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...dynamicCommand('tableResize', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableResize, (context) => context.activeTable
      ? { commandId: 'sheetTable.update', params: { ...structuredClone(context.activeTable.table), range: structuredClone(context.activeTable.resizeRange ?? context.activeTable.table.range) } }
      : undefined, 'layout'),
    enabled: (context) => Boolean(context.activeTable?.resizeRange),
  },
  {
    ...dynamicCommand('tableConvertToRange', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableConvertToRange, (context) => context.activeTable
      ? { commandId: 'sheetTable.convertToRange', params: { sheetId: context.activeTable.sheetId, tableId: context.activeTable.tableId } }
      : undefined, 'trash'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...intent('tableStyle', 'tableDesign', 'tableDesign', RIBBON_TEXT.commands.tableStyle, () => ({ type: 'panel.open', panel: 'data' }), 'sparkles'),
    enabled: (context) => Boolean(context.activeTable),
  },
  {
    ...intent('chartElements', 'chartDesign', 'chartDesign', RIBBON_TEXT.commands.chartElements, () => ({ type: 'panel.open', panel: 'chart' }), 'chart'),
    enabled: (context) => Boolean(context.activeChart),
  },
  {
    ...intent('chartSelectData', 'chartDesign', 'chartDesign', RIBBON_TEXT.commands.chartSelectData, () => ({ type: 'panel.open', panel: 'chart' }), 'table'),
    enabled: (context) => Boolean(context.activeChart),
  },
  {
    ...intent('chartFormatPanel', 'chartFormat', 'chartFormat', RIBBON_TEXT.commands.chartFormatPanel, () => ({ type: 'panel.open', panel: 'chart' }), 'sparkles'),
    enabled: (context) => Boolean(context.activeChart),
  },
  {
    ...intent('sparklineDesign', 'sparklineDesign', 'sparklineDesign', RIBBON_TEXT.commands.sparklineDesign, () => ({ type: 'panel.open', panel: 'sparkline' }), 'sparkline'),
    enabled: (context) => Boolean(context.activeSparkline),
  },
  {
    ...intent('pictureFormatPanel', 'pictureFormat', 'pictureFormat', RIBBON_TEXT.commands.pictureFormatPanel, () => ({ type: 'panel.open', panel: 'picture' }), 'picture'),
    enabled: (context) => Boolean(context.activePicture),
  },
  intent('chartBuilder', 'insert', 'insertCharts', RIBBON_TEXT.commands.chartBuilder, () => ({ type: 'panel.open', panel: 'chart' }), 'chart-column'),
  intent('sparkline', 'insert', 'insertCharts', RIBBON_TEXT.commands.sparkline, () => ({ type: 'panel.open', panel: 'sparkline' }), 'sparkline'),
  intent('shapesLines', 'insert', 'illustrations', RIBBON_TEXT.commands.shapesLines, () => ({ type: 'panel.open', panel: 'shape' }), 'shape-square'),

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
  intent('findReplace', 'data', 'findTransform', RIBBON_TEXT.commands.findReplace, () => ({ type: 'dialog.open', dialog: 'find-replace', findMode: 'replace' }), 'search'),
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
