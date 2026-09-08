import { useEffect, useLayoutEffect, useRef, useState, type ChangeEvent, type FormEvent, type KeyboardEvent } from 'react';
import { Button } from './Button';
import { Box, Inline } from './layout';
import { TextInput } from './TextInput';
import { Textarea } from './Textarea';

export interface FormulaBarLabels {
  selectedCell: string;
  nameBoxOptions: string;
  formulaInput: string;
  insertFunction: string;
  cancel: string;
  apply: string;
  placeholder: string;
  applyHint: string;
}

export interface FormulaBarProps {
  cellName: string;
  disabled: boolean;
  formula: string;
  caret?: { start: number; end: number };
  labels: FormulaBarLabels;
  onCancel: () => void;
  onFocusFormula?: () => void;
  onFormulaKeyDown?: (event: KeyboardEvent<HTMLTextAreaElement>) => void;
  onChange: (value: string) => void;
  onCommit: () => void;
  onCaretChange?: (start: number, end: number) => void;
  onCompositionStart?: (value: string, start: number, end: number) => void;
  onCompositionUpdate?: (value: string, start: number, end: number) => void;
  onCompositionEnd?: (value: string, start: number, end: number) => void;
  onNameBoxCommit?: (value: string) => void;
  onOpenNameManager?: () => void;
  onOpenWizard?: () => void;
}

export function FormulaBar({
  cellName,
  disabled,
  formula,
  caret,
  labels,
  onCancel,
  onChange,
  onCommit,
  onCaretChange,
  onCompositionStart,
  onCompositionUpdate,
  onCompositionEnd,
  onFocusFormula,
  onFormulaKeyDown,
  onNameBoxCommit,
  onOpenNameManager,
  onOpenWizard,
}: FormulaBarProps) {
  const [nameDraft, setNameDraft] = useState(cellName);
  const formulaInputRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    setNameDraft(cellName);
  }, [cellName]);

  useLayoutEffect(() => {
    const input = formulaInputRef.current;
    if (!input || !caret || input.ownerDocument.activeElement !== input) return;
    const start = Math.max(0, Math.min(input.value.length, caret.start));
    const end = Math.max(0, Math.min(input.value.length, caret.end));
    input.setSelectionRange(start, end);
  }, [caret?.start, caret?.end, formula]);

  const handleSubmit = (event: FormEvent<HTMLElement>) => {
    event.preventDefault();
    onCommit();
  };

  const commitNameBox = () => {
    if (!onNameBoxCommit || disabled) return;
    const next = nameDraft.trim();
    if (next && next !== cellName) onNameBoxCommit(next);
    else setNameDraft(cellName);
  };

  const handleNameKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      commitNameBox();
    } else if (event.key === 'Escape') {
      event.preventDefault();
      setNameDraft(cellName);
    }
  };

  return (
    <Box
      as="form"
      aria-label="Formula bar"
      className="flex h-[32px] items-center gap-0 border-y border-[#D1D1D1] bg-white pl-1 pr-2"
      data-testid="formula-bar"
      onSubmit={handleSubmit}
    >
      <TextInput
        aria-label={labels.selectedCell}
        containerClassName="w-[156px] flex-none"
        className="!h-[28px] !min-h-0 !w-full flex-none rounded-none border-[#D1D1D1] bg-white text-center font-sans text-[13px] text-[#242424]"
        data-testid="name-box"
        disabled={disabled}
        readOnly={!onNameBoxCommit}
        value={nameDraft}
        onBlur={commitNameBox}
        onChange={(event: ChangeEvent<HTMLInputElement>) => setNameDraft(event.target.value)}
        onKeyDown={handleNameKeyDown}
      />
      <Button
        aria-label={labels.nameBoxOptions}
        icon="more-vertical"
        iconOnly
        size="xs"
        variant="ghost"
        className="!h-[28px] !min-h-0 !w-7 rounded-none text-[#242424]"
        onClick={onOpenNameManager}
      />
      <Button
        aria-label={labels.cancel}
        data-testid="formula-cancel"
        disabled={disabled}
        icon="x"
        iconOnly
        onClick={onCancel}
        size="xs"
        variant="ghost"
        className="!h-[28px] !min-h-0 !w-7 rounded-none text-[#242424]"
      />
      <Button
        aria-label={labels.apply}
        data-testid="formula-apply"
        disabled={disabled}
        icon="check"
        iconOnly
        onClick={onCommit}
        size="xs"
        variant="ghost"
        className="!h-[28px] !min-h-0 !w-7 rounded-none text-[#242424]"
      />
      <Button aria-label={labels.insertFunction} disabled={disabled} icon="function" onClick={onOpenWizard} size="xs" variant="ghost" className="!h-[28px] !min-h-0 !w-9 rounded-none text-[#107C41]">fx</Button>
      <Inline gap="none" className="min-w-0 flex-1">
        <Textarea
          ref={formulaInputRef}
          aria-label={labels.formulaInput}
          className="!h-[28px] !min-h-0 !w-full resize-none overflow-auto rounded-none border-[#D1D1D1] px-2 py-1 font-sans text-[13px] leading-4"
          data-testid="formula-input"
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLTextAreaElement>) => {
            onChange(event.target.value);
            onCaretChange?.(event.target.selectionStart ?? event.target.value.length, event.target.selectionEnd ?? event.target.value.length);
          }}
          onFocus={onFocusFormula}
          onSelect={(event) => onCaretChange?.(event.currentTarget.selectionStart ?? 0, event.currentTarget.selectionEnd ?? 0)}
          onCompositionStart={(event) => onCompositionStart?.(event.currentTarget.value, event.currentTarget.selectionStart ?? 0, event.currentTarget.selectionEnd ?? 0)}
          onCompositionUpdate={(event) => onCompositionUpdate?.(event.currentTarget.value, event.currentTarget.selectionStart ?? 0, event.currentTarget.selectionEnd ?? 0)}
          onCompositionEnd={(event) => onCompositionEnd?.(event.currentTarget.value, event.currentTarget.selectionStart ?? event.currentTarget.value.length, event.currentTarget.selectionEnd ?? event.currentTarget.value.length)}
          onKeyDown={onFormulaKeyDown}
          placeholder=""
          value={formula}
        />
      </Inline>
    </Box>
  );
}
