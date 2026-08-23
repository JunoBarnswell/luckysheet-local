import type { RibbonAction } from './ribbon-actions';

export interface RibbonCommandDispatch {
  commandId: string;
  params?: unknown;
}

/** RibbonAction 字符串 → 统一 Command ID（禁止第二套执行路径） */
export function mapRibbonAction(action: RibbonAction, payload?: unknown): RibbonCommandDispatch {
  switch (action) {
    case 'undo':
      return { commandId: 'ui.history.undo' };
    case 'redo':
      return { commandId: 'ui.history.redo' };
    case 'copy':
      return { commandId: 'ui.clipboard.copy' };
    case 'cut':
      return { commandId: 'ui.clipboard.cut' };
    case 'paste':
      return { commandId: 'ui.clipboard.paste' };
    case 'bold':
      return { commandId: 'sheet.style.toggle', params: { property: 'bold' } };
    case 'italic':
      return { commandId: 'sheet.style.toggle', params: { property: 'italic' } };
    case 'underline':
      return { commandId: 'sheet.style.toggle', params: { property: 'underline' } };
    case 'strikethrough':
      return { commandId: 'sheet.style.toggle', params: { property: 'strikethrough' } };
    case 'align-left':
      return { commandId: 'sheet.style.toggle', params: { property: 'horizontalAlignment', value: 'left' } };
    case 'align-center':
      return { commandId: 'sheet.style.toggle', params: { property: 'horizontalAlignment', value: 'center' } };
    case 'align-right':
      return { commandId: 'sheet.style.toggle', params: { property: 'horizontalAlignment', value: 'right' } };
    case 'v-align-top':
      return { commandId: 'sheet.style.toggle', params: { property: 'verticalAlignment', value: 'top' } };
    case 'v-align-middle':
      return { commandId: 'sheet.style.toggle', params: { property: 'verticalAlignment', value: 'middle' } };
    case 'v-align-bottom':
      return { commandId: 'sheet.style.toggle', params: { property: 'verticalAlignment', value: 'bottom' } };
    case 'wrap-text':
      return { commandId: 'sheet.style.toggle', params: { property: 'wrapText' } };
    case 'merge-cells':
      return { commandId: 'sheet.merge.set' };
    case 'unmerge-cells':
      return { commandId: 'sheet.merge.remove' };
    case 'textColor':
      return { commandId: 'sheet.style.set', params: { style: { textColor: String(payload ?? '#111827') } } };
    case 'background':
      return { commandId: 'sheet.style.set', params: { style: { background: String(payload ?? '#ffffff') } } };
    case 'numberFormat':
      return { commandId: 'sheet.style.set', params: { style: { numberFormat: String(payload ?? 'general') } } };
    case 'format-currency':
      return { commandId: 'sheet.style.set', params: { style: { numberFormat: '$#,##0.00' } } };
    case 'format-percent':
      return { commandId: 'sheet.style.set', params: { style: { numberFormat: '0.00%' } } };
    case 'clear-range':
      return { commandId: 'sheet.range.clear' };
    case 'clear-formats':
      return { commandId: 'sheet.range.clear', params: { mode: 'formats' } };
    case 'autosum':
      return { commandId: 'sheet.formula.autosum' };
    case 'function-wizard':
      return { commandId: 'ui.dialog.open', params: { dialog: 'function-wizard' } };
    case 'sort-dialog':
      return { commandId: 'ui.dialog.open', params: { dialog: 'sort-dialog' } };
    case 'sort-asc':
      return { commandId: 'sheet.sort.multi', params: { ascending: true } };
    case 'sort-desc':
      return { commandId: 'sheet.sort.multi', params: { ascending: false } };
    case 'insert-row':
      return { commandId: 'sheet.rows.insert' };
    case 'insert-column':
      return { commandId: 'sheet.columns.insert' };
    case 'delete-row':
      return { commandId: 'sheet.rows.delete' };
    case 'delete-column':
      return { commandId: 'sheet.columns.delete' };
    case 'hide-row':
      return { commandId: 'sheet.row.hide' };
    case 'hide-column':
      return { commandId: 'sheet.column.hide' };
    case 'unhide-all':
      return { commandId: 'sheet.rows.unhide.all' };
    case 'transpose':
      return { commandId: 'matrix.transpose' };
    case 'flip-h':
      return { commandId: 'matrix.flip', params: { axis: 'h' } };
    case 'flip-v':
      return { commandId: 'matrix.flip', params: { axis: 'v' } };
    case 'split-column':
      return { commandId: 'sheet.splitColumn', params: payload ?? ',' };
    case 'banded-toggle':
      return { commandId: 'sheet.banded.set' };
    case 'freeze-top-row':
      return { commandId: 'sheet.freeze.set', params: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 } };
    case 'freeze-first-col':
      return { commandId: 'sheet.freeze.set', params: { xSplit: 1, ySplit: 0, startRow: 0, startColumn: 1 } };
    case 'freeze-at-primary':
      return { commandId: 'sheet.freeze.set', params: payload };
    case 'unfreeze':
      return { commandId: 'sheet.freeze.set', params: { xSplit: 0, ySplit: 0, startRow: 0, startColumn: 0 } };
    case 'filter-clear':
      return { commandId: 'sheet.filter.remove' };
    case 'apply-filter-selection':
      return { commandId: 'sheet.filter.set' };
    case 'export-xlsx':
      return { commandId: 'ui.file.export-xlsx' };
    case 'import-xlsx':
      return { commandId: 'ui.file.import-xlsx' };
    case 'find-replace':
      return { commandId: 'ui.dialog.open', params: { dialog: 'find-replace', findQuery: payload } };
    case 'zoom-in':
      return { commandId: 'ui.zoom.adjust', params: { delta: 10 } };
    case 'zoom-out':
      return { commandId: 'ui.zoom.adjust', params: { delta: -10 } };
    case 'zoom-100':
      return { commandId: 'ui.zoom.set', params: { value: 100 } };
    case 'open-chart':
      return { commandId: 'ui.panel.open', params: { panel: 'chart' } };
    case 'open-pivot':
      return { commandId: 'ui.panel.open', params: { panel: 'pivot' } };
    case 'open-data-table':
      return { commandId: 'ui.panel.open', params: { panel: 'data' } };
    case 'create-data-table':
      return { commandId: 'table.create' };
    case 'open-shape':
      return { commandId: 'ui.panel.open', params: { panel: 'shape' } };
    case 'open-sparkline':
      return { commandId: 'ui.panel.open', params: { panel: 'sparkline' } };
    case 'open-conditional-format':
      return { commandId: 'ui.panel.open', params: { panel: 'conditionalFormat' } };
    case 'open-data-validation':
      return { commandId: 'ui.panel.open', params: { panel: 'dataValidation' } };
    case 'open-history':
      return { commandId: 'ui.panel.open', params: { panel: 'history' } };
    case 'open-print':
      return { commandId: 'ui.dialog.open', params: { dialog: 'print-preview' } };
    case 'open-comments':
      return { commandId: 'ui.panel.open', params: { panel: 'inspector', notice: 'Select a cell and use Review tools for comments.' } };
    default:
      return { commandId: 'ui.notice', params: `Unsupported action: ${action}` };
  }
}
