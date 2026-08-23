import type { CommandContext, CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { RangeRef } from '@react-sheets/core-model';
import {
  buildPrintSnapshot,
} from './layout';
import type { PrintAreaSetCommandParams, PrintPreviewCommandParams } from './layout';
import {
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

export interface PrintDocumentReplaceCommandParams {
  document: PrintDocument;
}

interface PrintDocumentMutationParams {
  sheetId: string;
  document: PrintDocument;
}

function sheetRange(sheetId: string): RangeRef[] {
  return [{ sheetId, startRow: 0, endRow: Number.MAX_SAFE_INTEGER, startColumn: 0, endColumn: Number.MAX_SAFE_INTEGER }];
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

function isPrintDocument(value: unknown): value is PrintDocument {
  if (!isRecord(value) || value.schema !== 'PrintDocument' || typeof value.unitId !== 'string' || typeof value.sheetId !== 'string') return false;
  if (!isPageSetup(value.pageSetup) || !Array.isArray(value.printAreas) || !Array.isArray(value.pageBreaks)) return false;
  return value.printAreas.every((entry) => isRecord(entry) && typeof entry.sheetId === 'string' && isRange(entry.range))
    && value.pageBreaks.every((entry) => isRecord(entry) && typeof entry.sheetId === 'string' && (Number.isInteger(entry.row) !== Number.isInteger(entry.column)));
}

function isDocumentMutation(value: unknown): value is PrintDocumentMutationParams {
  return isRecord(value) && typeof value.sheetId === 'string' && isPrintDocument(value.document) && value.document.sheetId === value.sheetId;
}

function applyPrintPageSetup(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.pageSetup.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.pageSetup.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applyPrintAreaSet(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.area.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.area.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applyPrintAreaClear(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.area.clear', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.area.clear', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applyPrintPageBreakSet(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.pageBreak.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.pageBreak.set', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applyPrintPageBreakRemove(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.pageBreak.remove', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.pageBreak.remove', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applyPrintPageBreaksClear(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.pageBreaks.clear', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.pageBreaks.clear', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applyPrintDocumentReplace(context: CommandContext, sheetId: string, next: PrintDocument, previous: PrintDocument): void {
  const affectedRanges = sheetRange(sheetId);
  context.applyMutation({ id: 'print.document.replace', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(next) }, affectedRanges, inverse: [{ id: 'print.document.replace', unitId: context.workbook.unitId, sheetId, params: { sheetId, document: normalizePrintDocument(previous) }, affectedRanges }], apply: () => replacePrintDocument(context.workbook, next) });
}

function applySheetViewToggle(
  context: CommandContext,
  mutationId: 'pageLayout.gridlines.view.set' | 'pageLayout.headings.view.set',
  params: PrintToggleCommandParams,
  previous: boolean,
): void {
  const affectedRanges = sheetRange(params.sheetId);
  context.applyMutation({
    id: mutationId,
    unitId: context.workbook.unitId,
    sheetId: params.sheetId,
    params,
    affectedRanges,
    inverse: [{
      id: mutationId,
      unitId: context.workbook.unitId,
      sheetId: params.sheetId,
      params: { sheetId: params.sheetId, enabled: previous },
      affectedRanges,
    }],
    apply: () => {
      const sheet = context.workbook.getSheet(params.sheetId);
      if (mutationId === 'pageLayout.gridlines.view.set') sheet.showGridlines = params.enabled;
      else sheet.showHeaders = params.enabled;
    },
  });
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
  throw new Error('print.pageSetup requires pageSetup');
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
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.pageSetup.set', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintPageSetupMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.pageSetup.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.area.set', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintAreaSetMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.area.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.area.clear', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintAreaClearMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.area.clear'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.pageBreak.set', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintPageBreakSetMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.pageBreak.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.pageBreak.remove', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintPageBreakRemoveMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.pageBreak.remove'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.pageBreaks.clear', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintPageBreaksClearMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.pageBreaks.clear'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintDocumentMutationParams>({ id: 'print.document.replace', handler: (item, context) => replacePrintDocument(context.workbook, item.params.document), metadata: { schema: { name: 'PrintDocumentReplaceMutation', validate: isDocumentMutation }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['print.document.replace'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.gridlines.view.set', handler: (item, context) => { context.workbook.getSheet(item.params.sheetId).showGridlines = item.params.enabled; }, metadata: { schema: { name: 'PageLayoutGridlinesViewMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.gridlines.view.set'], minCount: 1, maxCount: 1 } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.headings.view.set', handler: (item, context) => { context.workbook.getSheet(item.params.sheetId).showHeaders = item.params.enabled; }, metadata: { schema: { name: 'PageLayoutHeadingsViewMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => sheetRange(params.sheetId), mode: 'exact' }, inversePolicy: { allowedMutationIds: ['pageLayout.headings.view.set'], minCount: 1, maxCount: 1 } } });

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
    id: 'print.pageSetup',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = pageSetupWithTitles(previous, params);
      applyPrintPageSetup(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintTitlesSetCommandParams>({
    id: 'print.titles.set',
    execute(params, context): CommandResult {
      if (!isPrintTitlesSet(params)) throw new Error('print.titles.set requires repeatRows and/or repeatColumns');
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
      applyPrintPageSetup(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'print.titles.clear',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = { ...previous, repeatRows: undefined, repeatColumns: undefined };
      applyPrintPageSetup(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintScaleSetCommandParams>({
    id: 'print.scale.set',
    execute(params, context): CommandResult {
      if (!isPrintScaleSet(params)) throw new Error('print.scale.set requires a scale between 1 and 400');
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
      applyPrintPageSetup(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'print.gridlines.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('print.gridlines.set requires enabled');
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = { ...previous, pageSetup: { ...previous.pageSetup, printGridlines: params.enabled } };
      applyPrintPageSetup(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'print.headings.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('print.headings.set requires enabled');
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next: PrintDocument = { ...previous, pageSetup: { ...previous.pageSetup, printHeadings: params.enabled } };
      applyPrintPageSetup(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.gridlines.view.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.gridlines.view.set requires enabled');
      const previous = context.workbook.getSheet(params.sheetId).showGridlines;
      applySheetViewToggle(context, 'pageLayout.gridlines.view.set', params, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.headings.view.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.headings.view.set requires enabled');
      const previous = context.workbook.getSheet(params.sheetId).showHeaders;
      applySheetViewToggle(context, 'pageLayout.headings.view.set', params, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintAreaSetCommandParams>({
    id: 'print.area.set',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, printAreas: [{ sheetId: params.sheetId, range: structuredClone(params.range) }] };
      applyPrintAreaSet(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [params.range] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'print.area.clear',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, printAreas: [] };
      applyPrintAreaClear(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintPageBreakSetCommandParams>({
    id: 'print.pageBreak.set',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const nextBreak = structuredClone(params.pageBreak);
      const pageBreaks = previous.pageBreaks.filter((item) => !(item.sheetId === nextBreak.sheetId && ((nextBreak.row !== undefined && item.row === nextBreak.row) || (nextBreak.column !== undefined && item.column === nextBreak.column))));
      pageBreaks.push(nextBreak);
      const next = { ...previous, pageBreaks };
      applyPrintPageBreakSet(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintPageBreakRemoveCommandParams>({
    id: 'print.pageBreak.remove',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, pageBreaks: previous.pageBreaks.filter((item) => !samePageBreak(item, params.pageBreak)) };
      applyPrintPageBreakRemove(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'print.pageBreaks.clear',
    execute(params, context): CommandResult {
      const previous = getPrintDocument(context.workbook, params.sheetId);
      const next = { ...previous, pageBreaks: [] };
      applyPrintPageBreaksClear(context, params.sheetId, next, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(params.sheetId) };
    },
  });

  registry.registerCommand<PrintDocumentReplaceCommandParams>({
    id: 'print.document.replace',
    execute(params, context): CommandResult {
      const document = normalizePrintDocument(params.document);
      const previous = getPrintDocument(context.workbook, document.sheetId);
      applyPrintDocumentReplace(context, document.sheetId, document, previous);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: sheetRange(document.sheetId) };
    },
  });
}
