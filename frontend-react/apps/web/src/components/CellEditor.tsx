import React, { useEffect, useMemo, useRef, type CSSProperties } from 'react';
import { CheckToggle, RichTextInput, Stack, Textarea } from '@react-sheets/ui-system';
import type { CellStyle } from '@react-sheets/core-model';
import type { CellContentLayoutResult } from '@react-sheets/render-engine';
import type { CellEditController, CellEditDraft, CellEditorSurfaceDescriptor } from '@react-sheets/spreadsheet-app';
import { toCanonicalKeyGesture } from '../editor/cell-edit-gesture';

export interface CellEditorProps {
  editorSurface: CellEditorSurfaceDescriptor;
  cellEdit: CellEditController;
  draft: CellEditDraft;
  cellStyle?: CellStyle;
  caret: { start: number; end: number };
  layout?: CellContentLayoutResult | null;
}

/** In-cell DOM surface. CellEditDomain owns every editing decision. */
export function CellEditor({ editorSurface, cellEdit, cellStyle, draft, caret, layout }: CellEditorProps): React.ReactElement {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const focusedRef = useRef(false);
  const editorStyle = useMemo<CSSProperties>(() => ({
    color: cellStyle?.textColor,
    fontFamily: cellStyle?.fontFamily,
    fontSize: cellStyle?.fontSizePx,
    fontStyle: cellStyle?.italic ? 'italic' : undefined,
    fontWeight: cellStyle?.bold ? 700 : undefined,
    textAlign: cellStyle?.horizontalAlignment === 'center' ? 'center' : cellStyle?.horizontalAlignment === 'right' ? 'right' : 'left',
    ...(layout ? { font: layout.font, lineHeight: `${layout.lineHeightPx}px` } : {}),
  }), [cellStyle, layout]);

  useEffect(() => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    if (!focusedRef.current) {
      textarea.focus({ preventScroll: true });
      focusedRef.current = true;
    }
    // Direct typing opens the editor with its first character as the initial
    // draft. Explicitly place the caret after it; browsers otherwise retain a
    // start-of-text selection during this focus transition and produce
    // `ello-h` / `1+2=` style reordered input.
    const start = Math.max(0, Math.min(textarea.value.length, caret.start));
    const end = Math.max(start, Math.min(textarea.value.length, caret.end));
    textarea.setSelectionRange(start, end);
  }, [draft.text, caret.start, caret.end]);

  const handleKeyDown = (event: React.KeyboardEvent<HTMLElement>) => {
    const result = cellEdit.dispatch({ type: 'keyboard', gesture: toCanonicalKeyGesture(event) });
    if (result.preventDefault) event.preventDefault();
    event.stopPropagation();
  };

  const focusGridEditor = () => {
    const status = cellEdit.getSnapshot().status;
    cellEdit.dispatch({ type: 'surface.focus', surface: 'grid' });
    if (status === 'enter') cellEdit.dispatch({ type: 'status.toggle' });
  };

  return (
    <Stack gap="none" className="h-full w-full" onPointerDown={focusGridEditor} onFocusCapture={focusGridEditor}>
      {editorSurface.kind === 'checkbox' ? (
        <CheckToggle
          aria-label="Cell editor checkbox"
          checked={/^true$/i.test(draft.text)}
          className="h-full w-full justify-center"
          onChange={(event) => {
            const text = event.currentTarget.checked ? 'TRUE' : 'FALSE';
            cellEdit.dispatch({ type: 'text.replace', text, caret: { start: text.length, end: text.length } });
          }}
          onKeyDown={handleKeyDown}
        />
      ) : draft.kind === 'rich-text' ? (
        <RichTextInput
          ariaLabel="Cell editor"
          caret={caret}
          runs={draft.runs}
          style={editorStyle}
          onCaretChange={(start, end) => cellEdit.dispatch({ type: 'caret.set', caret: { start, end } })}
          onCompositionStart={() => cellEdit.dispatch({ type: 'composition.start' })}
          onCompositionUpdate={(event) => cellEdit.dispatch({ type: 'composition.update', text: event.currentTarget.textContent ?? '' })}
          onCompositionEnd={(event) => {
            const value = event.currentTarget.textContent ?? '';
            cellEdit.dispatch({ type: 'composition.end', text: value, caret: cellEdit.getSnapshot().session?.caret ?? { start: value.length, end: value.length } });
          }}
          onInput={(value, start, end) => cellEdit.dispatch({ type: 'text.replace', text: value, caret: { start, end } })}
          onKeyDown={handleKeyDown}
        />
      ) : (
        <Textarea
        ref={textareaRef}
        style={editorStyle}
        aria-label="Cell editor"
        inputMode={editorSurface.inputMode ?? 'text'}
        autoCapitalize={editorSurface.autoCapitalize}
        data-pointer-gesture-owner="cell-editor"
        className={`h-full min-h-0 w-full resize-none rounded-none border-0 bg-transparent px-1 py-0 text-[13px] leading-[inherit] text-slate-800 outline-none focus:border-0 focus:ring-0 ${layout?.requiresInternalScroll ? 'overflow-auto' : 'overflow-visible'}`}
        value={draft.text}
        onCompositionStart={() => cellEdit.dispatch({ type: 'composition.start' })}
        onCompositionUpdate={(event) => cellEdit.dispatch({ type: 'composition.update', text: event.currentTarget.value })}
        onCompositionEnd={(event) => cellEdit.dispatch({ type: 'composition.end', text: event.currentTarget.value, caret: { start: event.currentTarget.selectionStart, end: event.currentTarget.selectionEnd } })}
        onChange={(event) => {
          cellEdit.dispatch({ type: 'text.replace', text: event.target.value, caret: { start: event.target.selectionStart, end: event.target.selectionEnd } });
        }}
        onSelect={(event) => cellEdit.dispatch({ type: 'caret.set', caret: { start: event.currentTarget.selectionStart, end: event.currentTarget.selectionEnd } })}
        onKeyDown={handleKeyDown}
        />
      )}
    </Stack>
  );
}
