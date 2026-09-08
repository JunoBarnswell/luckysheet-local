import type { CommandRegistry, CommandResult } from '@react-sheets/command-runtime';
import type { RangeRef } from '@react-sheets/core-model';
import { resolvePrintArea } from './layout';
import type { PrintAreaSetCommandParams, PrintPreviewCommandParams } from './layout';
import {
  DEFAULT_PAGE_SETUP,
  type PageSetup,
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
}

export interface PageLayoutBreakParams {
  sheetId: string;
  pageBreak: PrintPageBreak;
}

interface PageLayoutBreakClearParams {
  sheetId: string;
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

function isPageLayoutAreaClear(value: unknown): value is PageLayoutAreaClearParams {
  return isRecord(value) && typeof value.sheetId === 'string';
}

function isPageLayoutBreakClear(value: unknown): value is PageLayoutBreakClearParams {
  return isRecord(value) && typeof value.sheetId === 'string';
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

function printTitlesFromPageSetupParams(params: PrintPageSetupCommandParams): PrintTitlesSetCommandParams | undefined {
  const titleParams: PrintTitlesSetCommandParams = { sheetId: params.sheetId };
  let hasTitles = false;
  if (params.repeatRows !== undefined) {
    titleParams.repeatRows = params.repeatRows;
    hasTitles = true;
  }
  if (params.repeatColumns !== undefined) {
    titleParams.repeatColumns = params.repeatColumns;
    hasTitles = true;
  }
  if (params.layout?.repeatRows) {
    if (params.layout.repeatRows.sheetId !== params.sheetId) throw new Error('Print title rows must target the command sheet');
    titleParams.repeatRows = { start: params.layout.repeatRows.startRow, end: params.layout.repeatRows.endRow };
    hasTitles = true;
  }
  if (params.layout?.repeatColumns) {
    if (params.layout.repeatColumns.sheetId !== params.sheetId) throw new Error('Print title columns must target the command sheet');
    titleParams.repeatColumns = { start: params.layout.repeatColumns.startColumn, end: params.layout.repeatColumns.endColumn };
    hasTitles = true;
  }
  return hasTitles ? titleParams : undefined;
}

export function registerPrintCommands(registry: CommandRegistry): void {
  registry.registerMutation<PageLayoutSetupDetailSetParams>({ id: 'pageLayout.pageSetupDetail.set', metadata: { schema: { name: 'PageLayoutSetupDetailSetMutation', validate: isPageLayoutSetupDetailSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutMarginsSetParams>({ id: 'pageLayout.margins.set', metadata: { schema: { name: 'PageLayoutMarginsSetMutation', validate: isPageLayoutMarginsSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutOrientationSetParams>({ id: 'pageLayout.orientation.set', metadata: { schema: { name: 'PageLayoutOrientationSetMutation', validate: isPageLayoutOrientationSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutPaperSizeSetParams>({ id: 'pageLayout.paperSize.set', metadata: { schema: { name: 'PageLayoutPaperSizeSetMutation', validate: isPageLayoutPaperSizeSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PrintScaleSetCommandParams>({ id: 'pageLayout.scaleToFit.set', metadata: { schema: { name: 'PageLayoutScaleToFitSetMutation', validate: isPrintScaleSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PrintTitlesSetCommandParams>({ id: 'pageLayout.printTitles.set', metadata: { schema: { name: 'PageLayoutPrintTitlesSetMutation', validate: isPrintTitlesSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutAreaSetParams>({ id: 'pageLayout.printArea.set', metadata: { schema: { name: 'PageLayoutPrintAreaSetMutation', validate: isPageLayoutAreaSet }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: (params) => [params.range], mode: 'exact' } } });
  registry.registerMutation<PageLayoutAreaClearParams>({ id: 'pageLayout.printArea.clear', metadata: { schema: { name: 'PageLayoutPrintAreaClearMutation', validate: isPageLayoutAreaClear }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutBreakParams>({ id: 'pageLayout.pageBreak.insert', metadata: { schema: { name: 'PageLayoutPageBreakInsertMutation', validate: isPageLayoutBreak }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutBreakParams>({ id: 'pageLayout.pageBreak.remove', metadata: { schema: { name: 'PageLayoutPageBreakRemoveMutation', validate: isPageLayoutBreak }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PageLayoutBreakClearParams>({ id: 'pageLayout.pageBreak.clear', metadata: { schema: { name: 'PageLayoutPageBreakClearMutation', validate: isPageLayoutBreakClear }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.printGridlines.set', metadata: { schema: { name: 'PageLayoutPrintGridlinesSetMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.printHeadings.set', metadata: { schema: { name: 'PageLayoutPrintHeadingsSetMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.viewGridlines.set', metadata: { schema: { name: 'PageLayoutViewGridlinesMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });
  registry.registerMutation<PrintToggleCommandParams>({ id: 'pageLayout.viewHeadings.set', metadata: { schema: { name: 'PageLayoutViewHeadingsMutation', validate: isPrintToggle }, permission: { capability: 'print.layout.write', roles: ['owner', 'editor'] }, affectedRanges: { resolve: () => [], mode: 'exact' } } });

  registry.registerCommand<PrintPreviewCommandParams>({
    id: 'print.preview',
    execute(params, context): CommandResult {
      const sheetId = params.sheetId ?? context.workbook.primarySheetId;
      const sheet = context.workbook.getSheet(sheetId);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [resolvePrintArea(sheet, params.range)] };
    },
  });

  registry.registerCommand<PrintPreviewCommandParams>({
    id: 'print.export',
    execute(params, context): CommandResult {
      const sheetId = params.sheetId ?? context.workbook.primarySheetId;
      const sheet = context.workbook.getSheet(sheetId);
      return { operationId: context.operationId, mutationCount: 0, affectedRanges: [resolvePrintArea(sheet, params.range)] };
    },
  });

  registry.registerCommand<PrintPageSetupCommandParams>({
    id: 'pageLayout.pageSetup.set',
    execute(params, context): CommandResult {
      context.applyMutation({
        id: 'pageLayout.pageSetupDetail.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, pageSetup: structuredClone(pageSetupFromParams(params)) }, affectedRanges: [],
      });
      const titleParams = printTitlesFromPageSetupParams(params);
      if (titleParams) context.executeCommand('pageLayout.printTitles.set', titleParams);
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintTitlesSetCommandParams>({
    id: 'pageLayout.printTitles.set',
    execute(params, context): CommandResult {
      if (!isPrintTitlesSet(params)) throw new Error('pageLayout.printTitles.set requires repeatRows and/or repeatColumns');
      const mutationParams: PrintTitlesSetCommandParams = {
        sheetId: params.sheetId,
        ...(Object.prototype.hasOwnProperty.call(params, 'repeatRows') ? { repeatRows: params.repeatRows ?? null } : {}),
        ...(Object.prototype.hasOwnProperty.call(params, 'repeatColumns') ? { repeatColumns: params.repeatColumns ?? null } : {}),
      };
      context.applyMutation({
        id: 'pageLayout.printTitles.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: mutationParams, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'pageLayout.printTitles.clear',
    execute(params, context): CommandResult {
      context.applyMutation({
        id: 'pageLayout.printTitles.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, repeatRows: null, repeatColumns: null }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintScaleSetCommandParams>({
    id: 'pageLayout.scaleToFit.set',
    execute(params, context): CommandResult {
      if (!isPrintScaleSet(params)) throw new Error('pageLayout.scaleToFit.set requires a scale between 1 and 400');
      const mutationParams: PrintScaleSetCommandParams = {
        sheetId: params.sheetId,
        scale: params.scale,
        ...(params.fitToWidth === undefined ? {} : { fitToWidth: params.fitToWidth }),
        ...(params.fitToHeight === undefined ? {} : { fitToHeight: params.fitToHeight }),
      };
      context.applyMutation({
        id: 'pageLayout.scaleToFit.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: mutationParams, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.printGridlines.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.printGridlines.set requires enabled');
      context.applyMutation({
        id: 'pageLayout.printGridlines.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, enabled: params.enabled }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.printHeadings.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.printHeadings.set requires enabled');
      context.applyMutation({
        id: 'pageLayout.printHeadings.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, enabled: params.enabled }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.viewGridlines.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.viewGridlines.set requires enabled');
      context.applyMutation({
        id: 'pageLayout.viewGridlines.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintToggleCommandParams>({
    id: 'pageLayout.viewHeadings.set',
    execute(params, context): CommandResult {
      if (!isPrintToggle(params)) throw new Error('pageLayout.viewHeadings.set requires enabled');
      context.applyMutation({
        id: 'pageLayout.viewHeadings.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintAreaSetCommandParams>({
    id: 'pageLayout.printArea.set',
    execute(params, context): CommandResult {
      context.applyMutation({
        id: 'pageLayout.printArea.set', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, range: structuredClone(params.range) }, affectedRanges: [params.range],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [params.range] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'pageLayout.printArea.clear',
    execute(params, context): CommandResult {
      context.applyMutation({
        id: 'pageLayout.printArea.clear', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintPageBreakSetCommandParams>({
    id: 'pageLayout.pageBreak.insert',
    execute(params, context): CommandResult {
      const nextBreak = structuredClone(params.pageBreak);
      context.applyMutation({
        id: 'pageLayout.pageBreak.insert', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, pageBreak: nextBreak }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<PrintPageBreakRemoveCommandParams>({
    id: 'pageLayout.pageBreak.remove',
    execute(params, context): CommandResult {
      context.applyMutation({
        id: 'pageLayout.pageBreak.remove', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId, pageBreak: structuredClone(params.pageBreak) }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

  registry.registerCommand<{ sheetId: string }>({
    id: 'pageLayout.pageBreak.clear',
    execute(params, context): CommandResult {
      context.applyMutation({
        id: 'pageLayout.pageBreak.clear', unitId: context.workbook.unitId, sheetId: params.sheetId,
        params: { sheetId: params.sheetId }, affectedRanges: [],
      });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges: [] };
    },
  });

}
