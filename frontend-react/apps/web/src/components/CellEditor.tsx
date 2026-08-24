import React, { useEffect, useMemo, useRef, type CSSProperties } from 'react';
import { Stack, Textarea } from '@react-sheets/ui-system';
import type { CellStyle } from '@react-sheets/core-model';

export interface CellEditorProps {
  initialText: string;
  cellStyle?: CellStyle;
  onChange: (value: string) => void;
  onCommit: (moveAfter?: 'down' | 'up' | 'left' | 'right' | 'none') => void;
  onCancel: () => void;
  /** 向草稿插入引用文本(编辑中点击单元格/F4 由画布侧转发) */
  onInsertRef?: (refText: string) => void;
}

/**
 * 行内浮动编辑器。多行受控输入,回车提交(Shift+Enter 反向),
 * Tab 横向提交,Escape 取消。
 */
export function CellEditor({ cellStyle, initialText, onChange, onCommit, onCancel }: CellEditorProps): React.ReactElement {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  const composingRef = useRef(false);
  const editorStyle = useMemo<CSSProperties>(() => ({
    color: cellStyle?.textColor,
    fontFamily: cellStyle?.fontFamily,
    fontSize: cellStyle?.fontSizePx,
    fontStyle: cellStyle?.italic ? 'italic' : undefined,
    fontWeight: cellStyle?.bold ? 700 : undefined,
    textAlign: cellStyle?.horizontalAlignment === 'center' ? 'center' : cellStyle?.horizontalAlignment === 'right' ? 'right' : 'left',
  }), [cellStyle]);

  useEffect(() => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    textarea.focus({ preventScroll: true });
    // Direct typing opens the editor with its first character as the initial
    // draft. Explicitly place the caret after it; browsers otherwise retain a
    // start-of-text selection during this focus transition and produce
    // `ello-h` / `1+2=` style reordered input.
    const end = textarea.value.length;
    textarea.setSelectionRange(end, end);
  }, []);

  return (
    <Stack gap="none" className="h-full w-full">
      <Textarea
        ref={textareaRef}
        style={editorStyle}
        aria-label="Cell editor"
        className="h-full min-h-0 w-full resize-none overflow-hidden rounded-none border-0 bg-transparent px-1 py-0 text-[13px] leading-[inherit] text-slate-800 outline-none focus:border-0 focus:ring-0"
        value={initialText}
        onCompositionStart={() => { composingRef.current = true; }}
        onCompositionEnd={() => { composingRef.current = false; }}
        onChange={(event) => {
          onChange(event.target.value);
        }}
        onKeyDown={(event) => {
          if (event.nativeEvent.isComposing || composingRef.current) {
            event.stopPropagation();
            return;
          }
          if (event.key === 'Enter' && !event.shiftKey) {
            event.preventDefault();
            onCommit('down');
          } else if (event.key === 'Enter' && event.shiftKey) {
            event.preventDefault();
            onCommit('up');
          } else if (event.key === 'Tab') {
            event.preventDefault();
            onCommit(event.shiftKey ? 'left' : 'right');
          } else if (event.key === 'Escape') {
            event.preventDefault();
            onCancel();
          }
          event.stopPropagation();
        }}
      />
    </Stack>
  );
}
