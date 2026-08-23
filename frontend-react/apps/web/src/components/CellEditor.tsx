import React, { useEffect, useRef } from 'react';
import { Stack } from '@react-sheets/ui-system';

export interface CellEditorProps {
  initialText: string;
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
export function CellEditor({ initialText, onChange, onCommit, onCancel }: CellEditorProps): React.ReactElement {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    const textarea = textareaRef.current;
    if (!textarea) return;
    textarea.style.height = '0px';
    textarea.style.height = Math.max(28, textarea.scrollHeight + 2) + 'px';
  }, [initialText]);

  return (
    <Stack gap="none" className="w-full">
      <textarea
        ref={textareaRef}
        aria-label="Cell editor"
        className="w-full resize-none border-0 bg-transparent px-1 py-0.5 text-[13px] leading-5 text-slate-800 outline-none"
        value={initialText}
        onChange={(event) => {
          onChange(event.target.value);
          event.currentTarget.style.height = '0px';
          event.currentTarget.style.height = Math.max(28, event.currentTarget.scrollHeight + 2) + 'px';
        }}
        onKeyDown={(event) => {
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
