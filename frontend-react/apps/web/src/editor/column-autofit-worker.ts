import { DEFAULT_RENDER_THEME, hasMeasurableCellContent, measureCellAutoFit, type CellRenderStyle } from '@react-sheets/render-engine';
import type { CellPhoneticMetadata, RichTextRun } from '@react-sheets/core-model';
import type { AutoFitWorkerRequest } from './column-autofit-protocol';

type AutoFitCell = { column: number; value: string; style?: CellRenderStyle; richText?: RichTextRun[]; phonetic?: CellPhoneticMetadata; filterButton?: boolean };

const tasks = new Map<string, { maxima: Map<number, number> }>();
const canvas = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(1, 1) : null;
const context = canvas?.getContext('2d') ?? null;
const workerScope = typeof self === 'undefined' ? undefined : self;

if (workerScope) workerScope.onmessage = (event: MessageEvent<AutoFitWorkerRequest>) => {
  const request = event.data;
  if (request.kind === 'start') {
    tasks.set(request.taskId, { maxima: new Map(request.columns.map((column) => [column, 8])) });
    return;
  }
  if (request.kind === 'cancel') {
    tasks.delete(request.taskId);
    return;
  }
  const task = tasks.get(request.taskId);
  if (!task) return;
  if (request.kind === 'chunk') {
    const { block } = request;
    for (let index = 0; index < block.values.length; index += 1) {
      const styleIndex = block.styleIndexes[index] ?? 0;
      const cell: AutoFitCell = {
        column: block.columns[index] ?? 0,
        value: block.values[index] ?? '',
        style: styleIndex > 0 ? block.styles[styleIndex - 1] : undefined,
        richText: block.richTexts[index],
        phonetic: block.phonetics[index],
        filterButton: block.filterButtons[index] === 1,
      };
      if (!hasMeasurableCellContent(cell)) continue;
      if (!context) throw new Error('Canvas text measurement is unavailable in AutoFit worker');
      const width = measureCellAutoFit(context, { value: cell.value, displayValue: cell.value, style: cell.style, richText: cell.richText, phonetic: cell.phonetic }, DEFAULT_RENDER_THEME, undefined, cell.filterButton).widthPx;
      task.maxima.set(cell.column, Math.max(task.maxima.get(cell.column) ?? 8, width));
    }
    return;
  }
  tasks.delete(request.taskId);
  workerScope.postMessage({ kind: 'complete', taskId: request.taskId, widths: [...task.maxima].map(([column, widthPx]) => ({ column, widthPx })) });
};
