import type { CellStyle, WorkbookTableModel } from '@react-sheets/core-model';
import type { PrintLayout } from '@react-sheets/pro-features';
import {
  copyRangeToClipboardData,
  formatTsv,
  normalizeRangeRef,
  parseTsv,
} from '@react-sheets/sheet-features';
import { buildXlsxArchiveBase64 } from '@react-sheets/pro-features';
import type { SpreadsheetApplication } from './application';
import { columnLabel, usedRangeOfSheet } from './application-helpers';
import type { SidebarPanelId } from './types';

/** Ribbon / chrome command IDs — UI calls app.execute(id, params) */
export type UiCommandId =
  | 'ui.history.undo'
  | 'ui.history.redo'
  | 'ui.clipboard.cut'
  | 'ui.clipboard.copy'
  | 'ui.clipboard.paste'
  | 'sheet.style.toggle'
  | 'sheet.style.set'
  | 'sheet.merge.set'
  | 'sheet.merge.remove'
  | 'sheet.range.clear'
  | 'sheet.formula.autosum'
  | 'sheet.sort.multi'
  | 'sheet.rows.insert'
  | 'sheet.rows.delete'
  | 'sheet.columns.insert'
  | 'sheet.columns.delete'
  | 'sheet.row.hide'
  | 'sheet.column.hide'
  | 'sheet.rows.unhide.all'
  | 'sheet.columns.unhide.all'
  | 'matrix.transpose'
  | 'matrix.flip'
  | 'sheet.splitColumn'
  | 'sheet.banded.set'
  | 'sheet.freeze.set'
  | 'sheet.filter.remove'
  | 'sheet.filter.set'
  | 'ui.file.export-xlsx'
  | 'ui.file.import-xlsx'
  | 'ui.dialog.open'
  | 'ui.panel.open'
  | 'ui.zoom.set'
  | 'ui.zoom.adjust'
  | 'ui.freeze.atPrimary'
  | 'table.create'
  | 'ui.notice';

export interface StyleToggleParams {
  property: keyof CellStyle | 'horizontalAlignment' | 'verticalAlignment';
  value?: unknown;
}

export interface StyleSetParams {
  style: Partial<CellStyle>;
}

export interface DialogOpenParams {
  dialog: 'function-wizard' | 'sort-dialog' | 'find-replace' | 'print-preview' | 'goto' | 'paste-special' | 'format-cells' | 'shift-cells';
  findQuery?: string;
}

export interface PanelOpenParams {
  panel: SidebarPanelId;
  notice?: string;
}

export interface ZoomAdjustParams {
  delta?: number;
  value?: number;
}

const UI_COMMAND_PREFIXES = ['ui.', 'table.create'];

export function isUiCommand(commandId: string): boolean {
  return UI_COMMAND_PREFIXES.some((prefix) => commandId.startsWith(prefix)) || commandId === 'sheet.formula.autosum';
}

export function executeUiCommand(app: SpreadsheetApplication, commandId: string, params?: unknown): boolean {
  const activeSheetId = app.getActiveSheetId();
  const selection = app.getSelection();
  const primaryRange = app.getPrimaryRange();
  const sheet = app.getWorkbook().getSheet(activeSheetId);

  const applyStyleToPrimary = (style: Partial<CellStyle>) => {
    app.runCommand('sheet.style.set', { sheetId: activeSheetId, range: primaryRange, style });
  };
  const readStyle = (): CellStyle | undefined =>
    sheet.cells.get(selection.primaryRowIndex, selection.primaryColumnIndex)?.style;

  switch (commandId) {
    case 'ui.history.undo':
      app.undo();
      return true;
    case 'ui.history.redo':
      app.redo();
      return true;
    case 'ui.clipboard.copy': {
      const data = copyRangeToClipboardData(app.getWorkbook(), primaryRange);
      app.setClipboard({ ...data, isCut: false });
      void navigator.clipboard?.writeText(formatTsv(data.values));
      app.notify('Copied to clipboard');
      return true;
    }
    case 'ui.clipboard.cut': {
      const data = copyRangeToClipboardData(app.getWorkbook(), primaryRange);
      app.setClipboard({ ...data, isCut: true });
      void navigator.clipboard?.writeText(formatTsv(data.values));
      app.notify('Cut to clipboard');
      return true;
    }
    case 'ui.clipboard.paste': {
      const pasteParams = params as { mode?: 'all' | 'values' | 'formats' | 'formulas' | 'transpose' } | undefined;
      const internal = app.getClipboard();
      if (internal) {
        app.runCommand('sheet.range.paste', {
          sheetId: activeSheetId,
          targetOrigin: { row: selection.primaryRowIndex, column: selection.primaryColumnIndex },
          clipboard: internal,
          mode: pasteParams?.mode ?? 'all',
        });
        if (internal.isCut) {
          app.runCommand('sheet.range.clear', { sheetId: activeSheetId, range: internal.range, mode: 'contents' });
          app.clearClipboard();
        }
        app.syncDraftFromPrimary();
        app.notify('Pasted from clipboard');
        return true;
      }
      void navigator.clipboard
        ?.readText()
        .then((text) => {
          const parsed = parseTsv(text);
          if (parsed.length === 0) return;
          const clipboard: import('@react-sheets/sheet-features').ClipboardData = {
            range: primaryRange,
            values: parsed,
          };
          app.runCommand('sheet.range.paste', {
            sheetId: activeSheetId,
            targetOrigin: { row: selection.primaryRowIndex, column: selection.primaryColumnIndex },
            clipboard,
            mode: pasteParams?.mode ?? 'all',
          });
          app.syncDraftFromPrimary();
          app.notify('Pasted from clipboard');
        })
        .catch(() => app.notify('Clipboard unavailable'));
      return true;
    }
    case 'sheet.style.toggle': {
      const toggle = params as StyleToggleParams;
      const current = readStyle();
      switch (toggle.property) {
        case 'bold':
          applyStyleToPrimary({ bold: !current?.bold });
          break;
        case 'italic':
          applyStyleToPrimary({ italic: !current?.italic });
          break;
        case 'underline':
          applyStyleToPrimary({ underline: !current?.underline });
          break;
        case 'strikethrough':
          applyStyleToPrimary({ strikethrough: !current?.strikethrough });
          break;
        case 'wrapText':
          applyStyleToPrimary({ wrapText: !current?.wrapText });
          break;
        case 'horizontalAlignment':
          applyStyleToPrimary({ horizontalAlignment: toggle.value as CellStyle['horizontalAlignment'] });
          break;
        case 'verticalAlignment':
          applyStyleToPrimary({ verticalAlignment: toggle.value as CellStyle['verticalAlignment'] });
          break;
        default:
          break;
      }
      return true;
    }
    case 'sheet.style.set':
      applyStyleToPrimary((params as StyleSetParams).style);
      return true;
    case 'sheet.merge.set': {
      const range = primaryRange;
      if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
        app.runCommand('sheet.merge.set', {
          sheetId: activeSheetId,
          range: { ...range, endColumn: range.endColumn + 1 },
        });
      } else {
        app.runCommand('sheet.merge.set', { sheetId: activeSheetId, range });
      }
      return true;
    }
    case 'sheet.merge.remove': {
      for (const range of selection.ranges) {
        app.runCommand('sheet.merge.remove', { sheetId: activeSheetId, range });
      }
      return true;
    }
    case 'sheet.range.clear':
      app.runCommand('sheet.range.clear', {
        sheetId: activeSheetId,
        range: primaryRange,
        mode: (params as { mode?: 'formats' | 'contents' | 'all' })?.mode,
      });
      app.syncDraftFromPrimary();
      return true;
    case 'sheet.formula.autosum': {
      let sumStart = selection.primaryRowIndex - 1;
      while (sumStart >= 0) {
        const above = sheet.cells.get(sumStart, selection.primaryColumnIndex);
        if (!above || typeof above.value !== 'number') break;
        sumStart -= 1;
      }
      sumStart += 1;
      if (sumStart < selection.primaryRowIndex) {
        const label = columnLabel(selection.primaryColumnIndex);
        const formula = '=SUM(' + label + (sumStart + 1) + ':' + label + selection.primaryRowIndex + ')';
        app.setFormulaDraft(formula);
        app.commitFormula(formula);
      }
      return true;
    }
    case 'sheet.sort.multi': {
      const sortParams = params as { ascending?: boolean } | undefined;
      const ascending = sortParams?.ascending ?? true;
      const selectedSheet = app.getSelectedSheet();
      app.runCommand('sheet.sort.multi', {
        sheetId: activeSheetId,
        range:
          primaryRange.endRow > primaryRange.startRow || primaryRange.endColumn > primaryRange.startColumn
            ? primaryRange
            : normalizeRangeRef({
                sheetId: activeSheetId,
                startRow: 0,
                endRow: Math.min(sheet.rowCount - 1, 30),
                startColumn: 0,
                endColumn: Math.min(selectedSheet.columnCount - 1, 6),
              }),
        criteria: [{ column: selection.primaryColumnIndex, ascending }],
        hasHeader: true,
      });
      return true;
    }
    case 'sheet.rows.insert':
      app.insertRowsAtPrimary((params as { count?: number })?.count ?? 1);
      return true;
    case 'sheet.rows.delete':
      app.deleteRowsAtPrimary();
      return true;
    case 'sheet.columns.insert':
      app.insertColumnsAtPrimary((params as { count?: number })?.count ?? 1);
      return true;
    case 'sheet.columns.delete':
      app.deleteColumnsAtPrimary();
      return true;
    case 'sheet.row.hide':
      app.hideRowsAtPrimary();
      return true;
    case 'sheet.column.hide':
      app.hideColumnsAtPrimary();
      return true;
    case 'sheet.rows.unhide.all':
    case 'sheet.columns.unhide.all':
      app.unhideAll();
      return true;
    case 'matrix.transpose':
      app.transposeSelection();
      return true;
    case 'matrix.flip':
      app.flipSelection((params as { axis: 'h' | 'v' }).axis);
      return true;
    case 'sheet.splitColumn':
      app.splitByDelimiter(String(params ?? ','));
      return true;
    case 'sheet.banded.set':
      app.toggleBandedRows();
      return true;
    case 'sheet.freeze.set': {
      const freezeParams = params as { freeze?: import('@react-sheets/core-model').FreezeModel } | import('@react-sheets/core-model').FreezeModel;
      const freeze = freezeParams && typeof freezeParams === 'object' && 'freeze' in freezeParams
        ? freezeParams.freeze!
        : freezeParams as import('@react-sheets/core-model').FreezeModel;
      app.runCommand('sheet.freeze.set', { sheetId: activeSheetId, freeze });
      return true;
    }
    case 'ui.freeze.atPrimary':
      app.freezeAtPrimary();
      return true;
    case 'sheet.filter.remove':
      app.clearFilter();
      return true;
    case 'sheet.filter.set':
      app.applyFilterSelection();
      return true;
    case 'ui.file.export-xlsx': {
      void (async () => {
        try {
          const base64 = buildXlsxArchiveBase64(app.getWorkbook().snapshot());
          const link = document.createElement('a');
          link.href = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
          link.download = (app.getWorkbook().name || 'workbook') + '.xlsx';
          link.click();
          app.notify('Workbook exported as .xlsx');
        } catch {
          app.notify('Export failed');
        }
      })();
      return true;
    }
    case 'ui.file.import-xlsx': {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.xlsx';
      input.onchange = () => {
        const file = input.files?.[0];
        if (!file) return;
        void file
          .arrayBuffer()
          .then((buffer) => {
            let binary = '';
            const bytes = new Uint8Array(buffer);
            const chunkSize = 0x8000;
            for (let i = 0; i < bytes.length; i += chunkSize) {
              binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
            }
            return btoa(binary);
          })
          .then((base64) => app.importXlsxBase64(base64))
          .catch(() => app.notify('Import failed'));
      };
      input.click();
      return true;
    }
    case 'ui.dialog.open': {
      const dialogParams = params as DialogOpenParams;
      app.openDialog(dialogParams.dialog, dialogParams.findQuery);
      return true;
    }
    case 'ui.panel.open': {
      const panelParams = params as PanelOpenParams;
      app.setActivePanel(panelParams.panel);
      if (panelParams.notice) app.notify(panelParams.notice);
      return true;
    }
    case 'ui.zoom.adjust': {
      const zoomParams = params as ZoomAdjustParams;
      if (zoomParams.value != null) app.setZoom(zoomParams.value);
      else app.setZoom(app.getZoom() + (zoomParams.delta ?? 0));
      return true;
    }
    case 'ui.zoom.set':
      app.setZoom((params as ZoomAdjustParams).value ?? 100);
      return true;
    case 'table.create':
      app.createDataTableFromSelection();
      return true;
    case 'ui.notice':
      app.notify(String(params ?? ''));
      return true;
    default:
      return false;
  }
}

export type { PrintLayout };
