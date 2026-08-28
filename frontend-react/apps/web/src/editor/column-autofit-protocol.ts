import type { CellRenderStyle } from '@react-sheets/render-engine';
import type { CellPhoneticMetadata, RichTextRun } from '@react-sheets/core-model';

export interface AutoFitCellInput {
  column: number;
  value: string;
  style?: CellRenderStyle;
  richText?: RichTextRun[];
  phonetic?: CellPhoneticMetadata;
  filterButton?: boolean;
}

/**
 * Transfer-oriented AutoFit payload.  Coordinate and flag data travel in
 * typed arrays, while styles are interned once per block instead of repeated
 * on every occupied cell.
 */
export interface AutoFitBlock {
  columns: Uint32Array;
  values: string[];
  styleIndexes: Uint16Array;
  filterButtons: Uint8Array;
  styles: CellRenderStyle[];
  richTexts: Array<RichTextRun[] | undefined>;
  phonetics: Array<CellPhoneticMetadata | undefined>;
}

export type AutoFitWorkerRequest =
  | { kind: 'start'; taskId: string; columns: number[] }
  | { kind: 'chunk'; taskId: string; block: AutoFitBlock }
  | { kind: 'finish'; taskId: string }
  | { kind: 'cancel'; taskId: string };

export function createAutoFitBlock(cells: readonly AutoFitCellInput[]): AutoFitBlock {
  const columns = new Uint32Array(cells.length);
  const styleIndexes = new Uint16Array(cells.length);
  const filterButtons = new Uint8Array(cells.length);
  const values = new Array<string>(cells.length);
  const styles: CellRenderStyle[] = [];
  const richTexts = new Array<RichTextRun[] | undefined>(cells.length);
  const phonetics = new Array<CellPhoneticMetadata | undefined>(cells.length);
  const styleIndexesByKey = new Map<string, number>();
  for (let index = 0; index < cells.length; index += 1) {
    const cell = cells[index]!;
    columns[index] = cell.column;
    values[index] = cell.value;
    richTexts[index] = cell.richText;
    phonetics[index] = cell.phonetic;
    if (cell.filterButton) filterButtons[index] = 1;
    if (!cell.style) continue;
    const key = JSON.stringify(cell.style);
    let styleIndex = styleIndexesByKey.get(key);
    if (styleIndex === undefined) {
      styleIndex = styles.length + 1;
      if (styleIndex > 0xffff) throw new Error('AUTOFIT_STYLE_BLOCK_OVERFLOW');
      styleIndexesByKey.set(key, styleIndex);
      styles.push(cell.style);
    }
    styleIndexes[index] = styleIndex;
  }
  return { columns, values, styleIndexes, filterButtons, styles, richTexts, phonetics };
}

export function autoFitBlockTransferables(block: AutoFitBlock): Transferable[] {
  return [block.columns.buffer, block.styleIndexes.buffer, block.filterButtons.buffer];
}
