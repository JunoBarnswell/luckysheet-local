import { useEffect, useState, type ChangeEvent, type FormEvent, type KeyboardEvent } from 'react';
import { Button } from './Button';
import { Box, Inline } from './layout';
import { TextInput } from './TextInput';

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
  labels: FormulaBarLabels;
  onCancel: () => void;
  onFocusFormula?: () => void;
  onChange: (value: string) => void;
  onCommit: () => void;
  onNameBoxCommit?: (value: string) => void;
  onOpenNameManager?: () => void;
  onOpenWizard?: () => void;
}

export function FormulaBar({
  cellName,
  disabled,
  formula,
  labels,
  onCancel,
  onChange,
  onCommit,
  onFocusFormula,
  onNameBoxCommit,
  onOpenNameManager,
  onOpenWizard,
}: FormulaBarProps) {
  const [nameDraft, setNameDraft] = useState(cellName);

  useEffect(() => {
    setNameDraft(cellName);
  }, [cellName]);

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

  const handleFormulaKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onCancel();
    } else if (event.key === 'Enter') {
      event.preventDefault();
      onCommit();
    }
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
      className="flex h-[37px] items-center gap-0 border-y border-[#d9d9d9] border-t-[#eeeeee] bg-white px-1"
      data-testid="formula-bar"
      onSubmit={handleSubmit}
    >
      <TextInput
        aria-label={labels.selectedCell}
        containerClassName="w-[132px] flex-none"
        className="!h-[30px] !min-h-0 !w-full flex-none rounded-[3px] border-[#d9d9d9] bg-white text-center font-sans text-[13px] text-slate-900"
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
        className="!h-[30px] !min-h-0 !w-7 rounded-none text-[#6b6b6b]"
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
        className="!h-[30px] !min-h-0 !w-7 rounded-none text-slate-500"
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
        className="!h-[30px] !min-h-0 !w-7 rounded-none text-slate-500"
      />
      <Button aria-label={labels.insertFunction} disabled={disabled} icon="function" onClick={onOpenWizard} size="xs" variant="ghost" className="!h-[30px] !min-h-0 !w-8 rounded-none text-[#2572bc]">fx</Button>
      <Inline gap="none" className="min-w-0 flex-1">
        <TextInput
          aria-label={labels.formulaInput}
          containerClassName="min-w-0 flex-1"
          className="!h-[30px] !min-h-0 !w-full rounded-[3px] border-[#d9d9d9] px-2 font-sans text-[13px]"
          data-testid="formula-input"
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLInputElement>) => onChange(event.target.value)}
          onFocus={onFocusFormula}
          onKeyDown={handleFormulaKeyDown}
          placeholder=""
          value={formula}
        />
      </Inline>
    </Box>
  );
}
