import { DEFAULT_RENDER_THEME, measureCellAutoFit, type CellRenderStyle } from '@react-sheets/render-engine';

type AutoFitCell = { column: number; value: string; style?: CellRenderStyle; filterButton?: boolean };
type Request =
  | { kind: 'start'; taskId: string; columns: number[] }
  | { kind: 'chunk'; taskId: string; cells: AutoFitCell[] }
  | { kind: 'finish'; taskId: string }
  | { kind: 'cancel'; taskId: string };

const tasks = new Map<string, { maxima: Map<number, number> }>();
const canvas = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(1, 1) : null;
const context = canvas?.getContext('2d') ?? null;

self.onmessage = (event: MessageEvent<Request>) => {
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
    for (const cell of request.cells) {
      const width = context
        ? measureCellAutoFit(context, { value: cell.value, displayValue: cell.value, style: cell.style }, DEFAULT_RENDER_THEME, undefined, cell.filterButton).widthPx
        : fallbackWidth(cell);
      task.maxima.set(cell.column, Math.max(task.maxima.get(cell.column) ?? 8, width));
    }
    return;
  }
  tasks.delete(request.taskId);
  self.postMessage({ kind: 'complete', taskId: request.taskId, widths: [...task.maxima].map(([column, widthPx]) => ({ column, widthPx })) });
};

function fallbackWidth(cell: AutoFitCell): number {
  const size = cell.style?.fontSizePx ?? 13;
  const units = [...cell.value].reduce((total, character) => total + (/[^\u0000-\u00ff]/.test(character) ? 1 : 0.55), 0);
  return Math.ceil(units * size + (cell.style?.padding ?? DEFAULT_RENDER_THEME.cellPadding) * 2 + (cell.filterButton ? 18 : 0));
}

