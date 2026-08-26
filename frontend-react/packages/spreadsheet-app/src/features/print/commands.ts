import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { RangeRef } from '@react-sheets/core-model';
import {
  buildPrintSnapshot,
} from './layout';
import type { PrintAreaSetCommandParams, PrintPreviewCommandParams } from './layout';
import {
  DEFAULT_PAGE_SETUP,
  getPrintDocument,
  normalizePrintDocument,
  replacePrintDocument,
  type PageSetup,
  type PrintDocument,
  type PrintLayout,
  type PrintPageBreak,
  type PrintTitleSpan,
} from './index';

export interface PrintPageSetupCommandParams {
  sheetId: string;
  /** Canonical persisted setup. `layout` is the user-facing command shape. */
  pageSetup?: PageSetup;
  layout?: PrintLayout;
  /** Optional title spans for direct host calls; omitted preserves current titles. */
  repeatRows?: PrintTitleSpan | null;
  repeatColumns?: PrintTitleSpan | null;
}

export interface PrintTitlesSetCommandParams {
  sheetId: string;
  repeatRows?: PrintTitleSpan | null;
  repeatColumns?: PrintTitleSpan | null;
}

export interface PrintScaleSetCommandParams {
  sheetId: string;
  scale: number;
  fitToWidth?: number | null;
  fitToHeight?: number | null;
}

export interface PrintToggleCommandParams {
  sheetId: string;
  enabled: boolean;
}

export interface PrintPageBreakSetCommandParams {
  sheetId: string;
  pageBreak: PrintPageBreak;
}

export interface PrintPageBreakRemoveCommandParams {
  sheetId: string;
  pageBreak: PrintPageBreak;
}

export interface PageLayoutMarginsSetParams {
  sheetId: string;
  margins: PageSetup['margins'];
}

export interface PageLayoutOrientationSetParams {
  sheetId: string;
  orientation: PageSetup['orientation'];
}

export interface PageLayoutPaperSizeSetParams {
  sheetId: string;
  paperSize: PageSetup['paperSize'];
}

export interface PageLayoutScaleToFitSetParams {
  sheetId: string;
  scale: number;
  fitToWidth?: number | null;
  fitToHeight?: number | null;
}

export interface PageLayoutToggleSetParams {
  sheetId: string;
  enabled: boolean;
}

export interface PageLayoutTitlesSetParams {
  sheetId: string;
  repeatRows?: PrintTitleSpan | null;
  repeatColumns?: PrintTitleSpan | null;
}

export interface PageLayoutAreaSetParams {
  sheetId: string;
  range: RangeRef;
}

interface PageLayoutAreaClearParams {
  sheetId: string;
  printAreas?: PrintDocument['printAreas'];
}

export interface PageLayoutBreakParams {
  sheetId: string;
  pageBreak: PrintPageBreak;
}

interface PageLayoutBreakClearParams {
  sheetId: string;
  pageBreaks?: PrintDocument['pageBreaks'];
}

interface PageLayoutSetupDetailSetParams {
  sheetId: string;
  pageSetup: PageSetup;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isRange(value: unknown): value is RangeRef {
  if (!isRecord(value)) return false;
  return typeof value.sheetId === 'string'
    && Number.isInteger(value.startRow) && Number.isInteger(value.endRow)
    && Number.isInteger(value.startColumn) && Number.isInteger(value.endColumn)
    && Number(value.startRow) >= 0 && Number(value.endRow) >= Number(value.startRow)
    && Number(value.startColumn) >= 0 && Number(value.endColumn) >= Number(value.startColumn);
}

function isPageSetup(value: unknown): value is PageSetup {
  if (!isRecord(value) || !isRecord(value.margins)) return false;
  const margins = value.margins;
  return ['top', 'right', 'bottom', 'left', 'header', 'footer'].every((key) => typeof margins[key] === 'number' && Number.isFinite(margins[key]))
    && ['a3', 'a4', 'letter', 'legal', 'custom'].includes(String(value.paperSize))
    && ['portrait', 'landscape'].includes(String(value.orientation))
    && typeof value.scale === 'number' && value.scale > 0 && value.scale <= 400
    && typeof value.printGridlines === 'boolean'
    && typeof value.printHeadings === 'boolean'
    && typeof value.centerHorizontally === 'boolean'
    && typeof value.centerVertically === 'boolean';
}

function isPrintTitleSpan(value: unknown): value is PrintTitleSpan {
  return isRecord(value)
    && typeof value.start === 'number' && Number.isSafeInteger(value.start) && value.start >= 0
    && typeof value.end === 'number' && Number.isSafeInteger(value.end) && value.end >= value.start;
}

function isOptionalPrintTitleSpan(value: unknown): value is PrintTitleSpan | null {
  return value === undefined || value === null || isPrintTitleSpan(value);
}

function isPrintTitlesSet(value: unknown): value is PrintTitlesSetCommandParams {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && (Object.prototype.hasOwnProperty.call(value, 'repeatRows') || Object.prototype.hasOwnProperty.call(value, 'repeatColumns'))
    && isOptionalPrintTitleSpan(value.repeatRows)
    && isOptionalPrintTitleSpan(value.repeatColumns);
}

function isPrintScaleSet(value: unknown): value is PrintScaleSetCommandParams {
  return isRecord(value)
    && typeof value.sheetId === 'string'
    && typeof value.scale === 'number' && Number.isFinite(value.scale) && value.scale > 0 && value.scale <= 400
    && (value.fitToWidth === undefined || value.fitToWidth === null || (typeof value.fitToWidth === 'number' && Number.isSafeInteger(value.fitToWidth) && value.fitToWidth > 0))
    && (value.fitToHeight === undefined || value.fitToHeight === null || (typeof value.fitToHeight === 'number' && Number.isSafeInteger(value.fitToHeight) && value.fitToHeight > 0));
}

function isPrintToggle(value: unknown): value is PrintToggleCommandParams {
  return isRecord(value) && typeof value.sheetId === 'string' && typeof value.enabled === 'boolean';
}

function isPageLayoutMarginsSet(value: unknown): value is PageLayoutMarginsSetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isPageSetup({ ...DEFAULT_PAGE_SETUP, margins: value.margins } as PageSetup);
}

function isPageLayoutOrientationSet(value: unknown): value is PageLayoutOrientationSetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && (value.orientation === 'portrait' || value.orientation === 'landscape');
}

function isPageLayoutPaperSizeSet(value: unknown): value is PageLayoutPaperSizeSetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && ['letter', 'a4', 'a3', 'legal', 'custom'].includes(String(value.paperSize));
}

function isPageLayoutSetupDetailSet(value: unknown): value is PageLayoutSetupDetailSetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isPageSetup(value.pageSetup);
}

function isPageLayoutAreaSet(value: unknown): value is PageLayoutAreaSetParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRange(value.range) && value.range.sheetId === value.sheetId;
}

function isPageLayoutBreak(value: unknown): value is PageLayoutBreakParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isRecord(value.pageBreak) && value.pageBreak.sheetId === value.sheetId
    && (Number.isInteger(value.pageBreak.row) !== Number.isInteger(value.pageBreak.column));
}

function isPrintAreaList(value: unknown, sheetId: string): value is PrintDocument['printAreas'] {
  return Array.isArray(value) && value.every((entry) => isRecord(entry)
    && entry.sheetId === sheetId && isRange(entry.range) && entry.range.sheetId === sheetId);
}

function isPageBreakList(value: unknown, sheetId: string): value is PrintDocument['pageBreaks'] {
  return Array.isArray(value) && value.every((entry) => isRecord(entry)
    && entry.sheetId === sheetId && (Number.isInteger(entry.row) !== Number.isInteger(entry.column)));
}

function isPageLayoutAreaClear(value: unknown): value is PageLayoutAreaClearParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.printAreas === undefined || isPrintAreaList(value.printAreas, value.sheetId));
}

function isPageLayoutBreakClear(value: unknown): value is PageLayoutBreakClearParams {
  return isRecord(value) && typeof value.sheetId === 'string'
    && (value.pageBreaks === undefined || isPageBreakList(value.pageBreaks, value.sheetId));
}

function isPrintDocument(value: unknown): value is PrintDocument {
  if (!isRecord(value) || value.schema !== 'PrintDocument' || typeof value.unitId !== 'string' || typeof value.sheetId !== 'string') return false;
  if (!isPageSetup(value.pageSetup) || !Array.isArray(value.printAreas) || !Array.isArray(value.pageBreaks)) return false;
  return value.printAreas.every((entry) => isRecord(entry) && typeof entry.sheetId === 'string' && isRange(entry.range))
    && value.pageBreaks.every((entry) => isRecord(entry) && typeof entry.sheetId === 'string' && (Number.isInteger(entry.row) !== Number.isInteger(entry.column)));
}

function pageSetupFromParams(params: PrintPageSetupCommandParams): PageSetup {
  if (params.pageSetup) return params.pageSetup;
  if (params.layout) {
    const margin = params.layout.margin;
    const paperSize = params.layout.paper === 'A3' ? 'a3' : params.layout.paper === 'Letter' ? 'letter' : params.layout.paper === 'Legal' ? 'legal' : 'a4';
    return {
      paperSize,
      orientation: params.layout.orientation,
      margins: { top: margin.top * 72 / 25.4, right: margin.right * 72 / 25.4, bottom: margin.bottom * 72 / 25.4, left: margin.left * 72 / 25.4, header: 36, footer: 36 },
      scale: params.layout.scale ?? 100,
      fitToWidth: params.layout.fitToWidth ? 1 : undefined,
      fitToHeight: params.layout.fitToHeight ? 1 : undefined,
      printGridlines: params.layout.printGridlines ?? false,
      printHeadings: params.layout.printHeadings ?? false,
      centerHorizontally: params.layout.centerHorizontally ?? false,
      centerVertically: params.layout.centerVertically ?? false,
      headerText: params.layout.headerText,
      footerText: params.layout.footerText,
    };
  }
  throw new Error('pageLayout.pageSetup.set requires pageSetup');
}

function nextPrintTitles(
  previous: PrintDocument,
  params: PrintPageSetupCommandParams,
): Pick<PrintDocument, 'repeatRows' | 'repeatColumns'> {
  let repeatRows = previous.repeatRows;
  let repeatColumns = previous.repeatColumns;
  if (params.repeatRows !== undefined) repeatRows = params.repeatRows ?? undefined;
  if (params.repeatColumns !== undefined) repeatColumns = params.repeatColumns ?? undefined;
  if (params.layout?.repeatRows) {
    if (params.layout.repeatRows.sheetId !== params.sheetId) throw new Error('Print title rows must target the command sheet');
    repeatRows = { start: params.layout.repeatRows.startRow, end: params.layout.repeatRows.endRow };
  }
  if (params.layout?.repeatColumns) {
    if (params.layout.repeatColumns.sheetId !== params.sheetId) throw new Error('Print title columns must target the command sheet');
    repeatColumns = { start: params.layout.repeatColumns.startColumn, end: params.layout.repeatColumns.endColumn };
  }
  return { repeatRows, repeatColumns };
}

function pageSetupWithTitles(previous: PrintDocument, params: PrintPageSetupCommandParams): PrintDocument {
  return {
    ...previous,
    pageSetup: structuredClone(pageSetupFromParams(params)),
    ...nextPrintTitles(previous, params),
  };
}

function samePageBreak(left: PrintPageBreak, right: PrintPageBreak): boolean {
  return left.sheetId === right.sheetId && left.row === right.row && left.column === right.column;
}

export function registerPrintCommands(registry: CommandRegistry): void {
  registry.registerMutation<PageLayoutSetupDetailSetParams>({ id: 'pageLayout.pageSetupDetail.set', handler: (item, context) => { if (!isPageLayoutSetupDetailSet(item.params)) throw new Error('pageLayout.pageSetupDetail.set requires pageSetup'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: structuredClone(item.params.pageSetup) }); }, metadata: { schema: { name: 'PageLayoutSetupDetailSetMutation', validate: isPageLayoutSetupDetailSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.pageSetupDetail.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutMarginsSetParams>({ id: 'pageLayout.margins.set', handler: (item, context) => { if (!isPageLayoutMarginsSet(item.params)) throw new Error('pageLayout.margins.set requires margins'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: { ...document.pageSetup, margins: structuredClone(item.params.margins) } }); }, metadata: { schema: { name: 'PageLayoutMarginsSetMutation', validate: isPageLayoutMarginsSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.margins.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutOrientationSetParams>({ id: 'pageLayout.orientation.set', handler: (item, context) => { if (!isPageLayoutOrientationSet(item.params)) throw new Error('pageLayout.orientation.set requires orientation'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: { ...document.pageSetup, orientation: item.params.orientation } }); }, metadata: { schema: { name: 'PageLayoutOrientationSetMutation', validate: isPageLayoutOrientationSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.orientation.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutPaperSizeSetParams>({ id: 'pageLayout.paperSize.set', handler: (item, context) => { if (!isPageLayoutPaperSizeSet(item.params)) throw new Error('pageLayout.paperSize.set requires paperSize'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: { ...document.pageSetup, paperSize: item.params.paperSize } }); }, metadata: { schema: { name: 'PageLayoutPaperSizeSetMutation', validate: isPageLayoutPaperSizeSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.paperSize.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintScaleSetCommandParams>({ id: 'pageLayout.scaleToFit.set', handler: (item, context) => { if (!isPrintScaleSet(item.params)) throw new Error('pageLayout.scaleToFit.set requires scale'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: { ...document.pageSetup, scale: item.params.scale, fitToWidth: item.params.fitToWidth ?? undefined, fitToHeight: item.params.fitToHeight ?? undefined } }); }, metadata: { schema: { name: 'PageLayoutScaleToFitSetMutation', validate: isPrintScaleSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.scaleToFit.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintTitlesSetCommandParams>({ id: 'pageLayout.printTitles.set', handler: (item, context) => { if (!isPrintTitlesSet(item.params)) throw new Error('pageLayout.printTitles.set requires title spans'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, ...(Object.prototype.hasOwnProperty.call(item.params, 'repeatRows') ? { repeatRows: item.params.repeatRows ?? undefined } : {}), ...(Object.prototype.hasOwnProperty.call(item.params, 'repeatColumns') ? { repeatColumns: item.params.repeatColumns ?? undefined } : {}) }); }, metadata: { schema: { name: 'PageLayoutPrintTitlesSetMutation', validate: isPrintTitlesSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.printTitles.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutAreaSetParams>({ id: 'pageLayout.printArea.set', handler: (item, context) => { if (!isPageLayoutAreaSet(item.params)) throw new Error('pageLayout.printArea.set requires a same-sheet range'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, printAreas: [{ sheetId: item.params.sheetId, range: structuredClone(item.params.range) }] }); }, metadata: { schema: { name: 'PageLayoutPrintAreaSetMutation', validate: isPageLayoutAreaSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [params.range], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.printArea.clear'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutAreaClearParams>({ id: 'pageLayout.printArea.clear', handler: (item, context) => { if (!isPageLayoutAreaClear(item.params)) throw new Error('pageLayout.printArea.clear requires a sheetId and valid print areas'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, printAreas: item.params.printAreas === undefined ? [] : structuredClone(item.params.printAreas) }); }, metadata: { schema: { name: 'PageLayoutPrintAreaClearMutation', validate: isPageLayoutAreaClear }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.printArea.clear'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutBreakParams>({ id: 'pageLayout.pageBreak.insert', handler: (item, context) => { if (!isPageLayoutBreak(item.params)) throw new Error('pageLayout.pageBreak.insert requires a page break'); const document = getPrintDocument(context.workbook, item.params.sheetId); const pageBreaks = document.pageBreaks.filter((entry) => !samePageBreak(entry, item.params.pageBreak)); pageBreaks.push(structuredClone(item.params.pageBreak)); replacePrintDocument(context.workbook, { ...document, pageBreaks }); }, metadata: { schema: { name: 'PageLayoutPageBreakInsertMutation', validate: isPageLayoutBreak }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.pageBreak.remove'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutBreakParams>({ id: 'pageLayout.pageBreak.remove', handler: (item, context) => { if (!isPageLayoutBreak(item.params)) throw new Error('pageLayout.pageBreak.remove requires a page break'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageBreaks: document.pageBreaks.filter((entry) => !samePageBreak(entry, item.params.pageBreak)) }); }, metadata: { schema: { name: 'PageLayoutPageBreakRemoveMutation', validate: isPageLayoutBreak }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.pageBreak.insert'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PageLayoutBreakClearParams>({ id: 'pageLayout.pageBreak.clear', handler: (item, context) => { if (!isPageLayoutBreakClear(item.params)) throw new Error('pageLayout.pageBreak.clear requires a sheetId and valid page breaks'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageBreaks: item.params.pageBreaks === undefined ? [] : structuredClone(item.params.pageBreaks) }); }, metadata: { schema: { name: 'PageLayoutPageBreakClearMutation', validate: isPageLayoutBreakClear }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.pageBreak.clear'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.printGridlines.set', handler: (item, context) => { if (!isPrintToggle(item.params)) throw new Error('pageLayout.printGridlines.set requires enabled'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: { ...document.pageSetup, printGridlines: item.params.enabled } }); }, metadata: { schema: { name: 'PageLayoutPrintGridlinesSetMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.printGridlines.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.printHeadings.set', handler: (item, context) => { if (!isPrintToggle(item.params)) throw new Error('pageLayout.printHeadings.set requires enabled'); const document = getPrintDocument(context.workbook, item.params.sheetId); replacePrintDocument(context.workbook, { ...document, pageSetup: { ...document.pageSetup, printHeadings: item.params.enabled } }); }, metadata: { schema: { name: 'PageLayoutPrintHeadingsSetMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.printHeadings.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.viewGridlines.set', handler: (item, context) => { context.workbook.getSheet(item.params.sheetId).showGridlines = item.params.enabled; }, metadata: { schema: { name: 'PageLayoutViewGridlinesMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.viewGridlines.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.viewHeadings.set', handler: (item, context) => { context.workbook.getSheet(item.params.sheetId).showHeaders = item.params.enabled; }, metadata: { schema: { name: 'PageLayoutViewHeadingsMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.viewHeadings.set'], minCount: 1, maxCount: 1 } } });

  registry.registerCommand<PrintPreviewCommandParams>({
    id: 'print.preview',
    execute(params, context): CommandResult {
      const sheetId = params.sheetId ?? context.workbook.primarySheetId;
      const snapshot = buildPrintSnapshot(context.workbook, sheetId, params.layout, params.range);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [snapshot.printArea] };
    },
  });

  registry.registerCommand<PrintPreviewCommandParams>({
    id: 'print.export',
    execute(params, context): CommandResult {
      const sheetId = params.sheetId ?? context.workbook.primarySheetId;
      const snapshot = buildPrintSnapshot(context.workbook, sheetId, params.layout, params.range);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [snapshot.printArea] };
    },
  });

  registry.registerCommand<PrintPageSetupCommandParams>({
    id: 'pageLayout.pageSetup.set',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = pageSetupWithTitles(previous, params);
      context.applyMutation({
        id: 'pageLayout.pageSetupDetail.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, pageSetup: normalizePrintDocument(next).pageSetup }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.pageSetupDetail.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, pageSetup: normalizePrintDocument(previous).pageSetup }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintTitlesSetCommandParams>({
    id: 'pageLayout.printTitles.set',
    execute(params, context): CommandResult {
      if (!isPrintTitlesSet(params)) throw new Error('pageLayout.printTitles.set requires repeatRows and/or repeatColumns');
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = {
        ...previous,
        ...(Object.prototype.hasOwnProperty.call(params, 'repeatRows')
          ? { repeatRows: params.repeatRows ?? undefined }
          : {}),
        ...(Object.prototype.hasOwnProperty.call(params, 'repeatColumns')
          ? { repeatColumns: params.repeatColumns ?? undefined }
          : {}),
      };
      context.applyMutation({
        id: 'pageLayout.printTitles.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, repeatRows: params.repeatRows ?? null, repeatColumns: params.repeatColumns ?? null }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.printTitles.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, repeatRows: previous.repeatRows ?? null, repeatColumns: previous.repeatColumns ?? null }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'pageLayout.printTitles.clear',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = { ...previous, repeatRows: undefined, repeatColumns: undefined };
      context.applyMutation({
        id: 'pageLayout.printTitles.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, repeatRows: null, repeatColumns: null }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.printTitles.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, repeatRows: previous.repeatRows ?? null, repeatColumns: previous.repeatColumns ?? null }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintScaleSetCommandParams>({
    id: 'pageLayout.scaleToFit.set',
    execute(params, context): CommandResult {
      if (!isPrintScaleSet(params)) throw new Error('pageLayout.scaleToFit.set requires a scale between 1 and 400');
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = {
        ...previous,
        pageSetup: {
          ...previous.pageSetup,
          scale: params.scale,
          ...(params.fitToWidth === undefined ? {} : { fitToWidth: params.fitToWidth ?? undefined }),
          ...(params.fitToHeight === undefined ? {} : { fitToHeight: params.fitToHeight ?? undefined }),
        },
      };
      context.applyMutation({
        id: 'pageLayout.scaleToFit.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, scale: params.scale, fitToWidth: params.fitToWidth ?? null, fitToHeight: params.fitToHeight ?? null }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.scaleToFit.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, scale: previous.pageSetup.scale, fitToWidth: previous.pageSetup.fitToWidth ?? null, fitToHeight: previous.pageSetup.fitToHeight ?? null }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.printGridlines.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.printGridlines.set requires enabled');
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = { ...previous, pageSetup: { ...previous.pageSetup, printGridlines: params.enabled } };
      context.applyMutation({
        id: 'pageLayout.printGridlines.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, enabled: params.enabled }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.printGridlines.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, enabled: previous.pageSetup.printGridlines }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.printHeadings.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.printHeadings.set requires enabled');
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = { ...previous, pageSetup: { ...previous.pageSetup, printHeadings: params.enabled } };
      context.applyMutation({
        id: 'pageLayout.printHeadings.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, enabled: params.enabled }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.printHeadings.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, enabled: previous.pageSetup.printHeadings }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.viewGridlines.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.viewGridlines.set requires enabled');
      const previous = context.workbook.getSheet(params.sheetId).showGridlines;
      context.applyMutation({
        id: 'pageLayout.viewGridlines.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params, affectedRanges: [],
        inverse: [{ id: 'pageLayout.viewGridlines.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, enabled: previous }, affectedRanges: [] }],
        apply: () => { context.workbook.getSheet(params.sheetId).showGridlines = params.enabled; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.viewHeadings.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.viewHeadings.set requires enabled');
      const previous = context.workbook.getSheet(params.sheetId).showHeaders;
      context.applyMutation({
        id: 'pageLayout.viewHeadings.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params, affectedRanges: [],
        inverse: [{ id: 'pageLayout.viewHeadings.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, enabled: previous }, affectedRanges: [] }],
        apply: () => { context.workbook.getSheet(params.sheetId).showHeaders = params.enabled; },
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintAreaSetCommandParams>({
    id: 'pageLayout.printArea.set',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, printAreas: [{ sheetId: params.sheetId, range: structuredClone(params.range) }] };
      context.applyMutation({
        id: 'pageLayout.printArea.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, range: structuredClone(params.range) }, affectedRanges: [params.range],
        inverse: [{ id: 'pageLayout.printArea.clear', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, printAreas: structuredClone(previous.printAreas) }, affectedRanges: [params.range] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [params.range] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'pageLayout.printArea.clear',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, printAreas: [] };
      context.applyMutation({
        id: 'pageLayout.printArea.clear', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.printArea.clear', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, printAreas: structuredClone(previous.printAreas) }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintPageBreakSetCommandParams>({
    id: 'pageLayout.pageBreak.insert',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const nextBreak = structuredClone(params.pageBreak);
      const pageBreaks = previous.pageBreaks.filter((item) => !(item.sheetId === nextBreak.sheetId && ((nextBreak.row !== undefined && item.row === nextBreak.row) || (nextBreak.column !== undefined && item.column === nextBreak.column))));
      pageBreaks.push(nextBreak);
      const next = { ...previous, pageBreaks };
      context.applyMutation({
        id: 'pageLayout.pageBreak.insert', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, pageBreak: nextBreak }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.pageBreak.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, pageBreak: nextBreak }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintPageBreakRemoveCommandParams>({
    id: 'pageLayout.pageBreak.remove',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, pageBreaks: previous.pageBreaks.filter((item) => !samePageBreak(item, params.pageBreak)) };
      const previousBreak = previous.pageBreaks.find((entry) => samePageBreak(entry, params.pageBreak));
      if (!previousBreak) throw new Error('pageLayout.pageBreak.remove requires an existing page break');
      context.applyMutation({
        id: 'pageLayout.pageBreak.remove', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, pageBreak: structuredClone(params.pageBreak) }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.pageBreak.insert', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, pageBreak: structuredClone(previousBreak) }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'pageLayout.pageBreak.clear',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, pageBreaks: [] };
      context.applyMutation({
        id: 'pageLayout.pageBreak.clear', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId }, affectedRanges: [],
        inverse: [{ id: 'pageLayout.pageBreak.clear', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, pageBreaks: structuredClone(previous.pageBreaks) }, affectedRanges: [] }],
        apply: () => replacePrintDocument(context.workbook, next),
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

}
