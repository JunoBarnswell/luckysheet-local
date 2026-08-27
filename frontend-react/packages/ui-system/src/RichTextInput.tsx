import { useLayoutEffect, useRef, type CompositionEvent, type CSSProperties, type FormEvent, type KeyboardEvent, type ReactElement } from 'react';
import { cn } from './cn';

export interface RichTextInputRun {
  text: string;
  style?: {
    fontFamily?: string;
    fontSizePx?: number;
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strikethrough?: boolean;
    textColor?: string;
    verticalAlignment?: 'baseline' | 'superscript' | 'subscript';
  };
}

export interface RichTextInputProps {
  ariaLabel: string;
  caret: { start: number; end: number };
  className?: string;
  runs: readonly RichTextInputRun[];
  style?: CSSProperties;
  onCaretChange(start: number, end: number): void;
  onCompositionStart(event: CompositionEvent<HTMLElement>): void;
  onCompositionUpdate(event: CompositionEvent<HTMLElement>): void;
  onCompositionEnd(event: CompositionEvent<HTMLElement>): void;
  onInput(value: string, start: number, end: number): void;
  onKeyDown(event: KeyboardEvent<HTMLElement>): void;
}

function selectionOffsets(root: HTMLElement): { start: number; end: number } {
  const selection = root.ownerDocument.getSelection();
  if (!selection || selection.rangeCount === 0) return { start: root.textContent?.length ?? 0, end: root.textContent?.length ?? 0 };
  const range = selection.getRangeAt(0);
  if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) return { start: 0, end: 0 };
  const beforeStart = range.cloneRange();
  beforeStart.selectNodeContents(root);
  beforeStart.setEnd(range.startContainer, range.startOffset);
  const beforeEnd = range.cloneRange();
  beforeEnd.selectNodeContents(root);
  beforeEnd.setEnd(range.endContainer, range.endOffset);
  return { start: beforeStart.toString().length, end: beforeEnd.toString().length };
}

function textPosition(root: HTMLElement, requestedOffset: number): { node: Node; offset: number } {
  const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let remaining = Math.max(0, requestedOffset);
  let node = walker.nextNode();
  while (node) {
    const length = node.textContent?.length ?? 0;
    if (remaining <= length) return { node, offset: remaining };
    remaining -= length;
    node = walker.nextNode();
  }
  return { node: root, offset: root.childNodes.length };
}

function applyCaret(root: HTMLElement, caret: RichTextInputProps['caret']): void {
  const selection = root.ownerDocument.getSelection();
  if (!selection) return;
  const start = textPosition(root, caret.start);
  const end = textPosition(root, caret.end);
  const range = root.ownerDocument.createRange();
  range.setStart(start.node, start.offset);
  range.setEnd(end.node, end.offset);
  selection.removeAllRanges();
  selection.addRange(range);
}

function runStyle(run: RichTextInputRun): CSSProperties {
  const style = run.style;
  return {
    fontFamily: style?.fontFamily,
    fontSize: style?.fontSizePx,
    fontStyle: style?.italic ? 'italic' : undefined,
    fontWeight: style?.bold ? 700 : undefined,
    textDecoration: [style?.underline ? 'underline' : '', style?.strikethrough ? 'line-through' : ''].filter(Boolean).join(' ') || undefined,
    color: style?.textColor,
    verticalAlign: style?.verticalAlignment === 'superscript' ? 'super' : style?.verticalAlignment === 'subscript' ? 'sub' : undefined,
  };
}

export function RichTextInput({ ariaLabel, caret, className, runs, style, onCaretChange, onCompositionStart, onCompositionUpdate, onCompositionEnd, onInput, onKeyDown }: RichTextInputProps): ReactElement {
  const rootRef = useRef<HTMLDivElement | null>(null);
  const focusedRef = useRef(false);

  useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    if (!focusedRef.current) {
      root.focus({ preventScroll: true });
      focusedRef.current = true;
    }
    applyCaret(root, caret);
  }, [caret.start, caret.end, runs]);

  const publishCaret = () => {
    const root = rootRef.current;
    if (!root) return;
    const selection = selectionOffsets(root);
    onCaretChange(selection.start, selection.end);
  };

  const publishInput = (event: FormEvent<HTMLDivElement>) => {
    const selection = selectionOffsets(event.currentTarget);
    onInput(event.currentTarget.textContent ?? '', selection.start, selection.end);
  };

  return (
    <div
      ref={rootRef}
      aria-label={ariaLabel}
      contentEditable
      suppressContentEditableWarning
      role="textbox"
      aria-multiline="true"
      data-pointer-gesture-owner="cell-editor"
      className={cn('h-full min-h-0 w-full overflow-auto whitespace-pre-wrap break-words px-1 py-0 outline-none', className)}
      style={style}
      onCompositionStart={onCompositionStart}
      onCompositionUpdate={onCompositionUpdate}
      onCompositionEnd={onCompositionEnd}
      onInput={publishInput}
      onKeyDown={onKeyDown}
      onKeyUp={publishCaret}
      onMouseUp={publishCaret}
    >
      {runs.map((run, index) => <span key={`${index}:${run.text.length}`} style={runStyle(run)}>{run.text}</span>)}
    </div>
  );
}
