import type { RichTextRunStyle } from '@react-sheets/core-model';
import type { CellEditCaret, CellEditDraft } from './contracts';

export type RichTextBooleanStyleKey = 'bold' | 'italic' | 'underline' | 'strikethrough';

export function richTextSelectionHasFlag(draft: CellEditDraft, caret: CellEditCaret, key: RichTextBooleanStyleKey): boolean {
  if (draft.kind !== 'rich-text') return false;
  const start = Math.min(caret.start, caret.end);
  const end = Math.max(caret.start, caret.end);
  if (start === end) return false;
  let offset = 0;
  let found = false;
  for (const run of draft.runs) {
    const runStart = offset;
    const runEnd = offset + run.text.length;
    offset = runEnd;
    if (runEnd <= start || runStart >= end) continue;
    found = true;
    if (run.style?.[key] !== true) return false;
  }
  return found;
}

export function richTextSelectionStyle(draft: CellEditDraft, caret: CellEditCaret): Partial<RichTextRunStyle> {
  if (draft.kind !== 'rich-text') return {};
  const start = Math.min(caret.start, caret.end);
  const end = Math.max(caret.start, caret.end);
  if (start === end) return {};
  let offset = 0;
  const selected = draft.runs.filter((run) => {
    const runStart = offset;
    const runEnd = offset + run.text.length;
    offset = runEnd;
    return runEnd > start && runStart < end;
  });
  if (selected.length === 0) return {};
  const keys: Array<keyof RichTextRunStyle> = ['fontFamily', 'fontSizePx', 'bold', 'italic', 'underline', 'strikethrough', 'textColor', 'verticalAlignment'];
  const result: Partial<RichTextRunStyle> = {};
  for (const key of keys) {
    const value = selected[0]?.style?.[key];
    if (value !== undefined && selected.every((run) => run.style?.[key] === value)) Object.assign(result, { [key]: value });
  }
  return result;
}

