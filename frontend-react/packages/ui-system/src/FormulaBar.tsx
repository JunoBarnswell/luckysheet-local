import type { ChangeEvent, FormEvent, KeyboardEvent } from 'react';
import { Button } from './Button';
import { Box, Inline, Kbd, Text } from './layout';
import { TextInput } from './TextInput';

export interface FormulaBarLabels {
  selectedCell: string;
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
  onChange: (value: string) => void;
  onCommit: () => void;
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
  onOpenWizard,
}: FormulaBarProps) {
  const handleSubmit = (event: FormEvent<HTMLElement>) => {
    event.preventDefault();
    onCommit();
  };

  const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault();
      onCancel();
    } else if (event.key === 'Enter') {
      event.preventDefault();
      onCommit();
    }
  };

  return (
    <Box
      as="form"
      aria-label="Formula bar"
      className="flex h-10 items-center gap-2 border-b border-slate-200 bg-white px-4"
      data-testid="formula-bar"
      onSubmit={handleSubmit}
    >
      <TextInput
        aria-label={labels.selectedCell}
        className="w-20 text-center font-mono text-xs font-bold text-slate-800 bg-slate-50 border-slate-200"
        data-testid="name-box"
        disabled={disabled}
        readOnly
        value={cellName}
      />
      <Inline gap="sm" className="min-w-0 flex-1">
        <Button aria-label={labels.insertFunction} disabled={disabled} icon="function" onClick={onOpenWizard} size="sm" variant="outline">fx</Button>
        <TextInput
          aria-label={labels.formulaInput}
          className="font-mono text-xs"
          data-testid="formula-input"
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLInputElement>) => onChange(event.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={labels.placeholder}
          value={formula}
        />
      </Inline>
      <Inline gap="xs" className="shrink-0">
        <Button
          aria-label={labels.cancel}
          data-testid="formula-cancel"
          disabled={disabled}
          icon="x"
          iconOnly
          onClick={onCancel}
          size="sm"
          variant="ghost"
        />
        <Button
          aria-label={labels.apply}
          data-testid="formula-apply"
          disabled={disabled}
          icon="check"
          iconOnly
          onClick={onCommit}
          size="sm"
          variant="soft"
        />
        <Inline gap="xs" className="hidden pl-2 lg:flex">
          <Text size="xs" tone="subtle">{labels.applyHint}</Text>
          <Kbd>Enter</Kbd>
        </Inline>
      </Inline>
    </Box>
  );
}
