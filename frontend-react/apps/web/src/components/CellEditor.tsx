import React, { useEffect, useRef, useState } from 'react';
import type { Rect } from '@react-sheets/render-engine';

export interface CellEditorProps {
  initialValue: string;
  rect: Rect;
  onCommit: (value: string) => void;
  onCancel: () => void;
  onNavigate?: (direction: 'down' | 'up' | 'left' | 'right') => void;
}

export function CellEditor({
  initialValue,
  rect,
  onCommit,
  onCancel,
  onNavigate,
}: CellEditorProps) {
  const [value, setValue] = useState(initialValue);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    setValue(initialValue);
    if (textareaRef.current) {
      textareaRef.current.focus();
      textareaRef.current.select();
    }
  }, [initialValue]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    if (e.key === 'Enter') {
      if (e.shiftKey) {
        // Multiline support in cell editor
        return;
      }
      e.preventDefault();
      onCommit(value);
      onNavigate?.('down');
    } else if (e.key === 'Escape') {
      e.preventDefault();
      onCancel();
    } else if (e.key === 'Tab') {
      e.preventDefault();
      onCommit(value);
      onNavigate?.(e.shiftKey ? 'left' : 'right');
    }
  };

  const handleBlur = () => {
    onCommit(value);
  };

  const minWidth = Math.max(rect.width, 100);
  const minHeight = Math.max(rect.height, 28);

  return (
    <div
      className="absolute z-30 shadow-lg"
      style={{
        left: rect.x - 1,
        top: rect.y - 1,
        minWidth,
        minHeight,
      }}
    >
      <textarea
        ref={textareaRef}
        value={value}
        onChange={(e) => setValue(e.target.value)}
        onKeyDown={handleKeyDown}
        onBlur={handleBlur}
        rows={1}
        className="w-full resize-none overflow-hidden rounded-xs border-2 border-blue-600 bg-white p-1 text-xs font-medium text-slate-900 outline-hidden font-sans"
        style={{
          minWidth,
          minHeight,
        }}
      />
    </div>
  );
}
