import type { ChangeEvent, FormEvent, KeyboardEvent } from 'react';
import { Box, Button, Icon, Inline, Kbd, Text, TextInput } from '@react-sheets/ui-system';
import type { WorkspacePhase } from '../state/workspace';

export interface FormulaBarProps {
  cellName: string;
  disabled: boolean;
  formula: string;
  onCancel: () => void;
  onChange: (value: string) => void;
  onCommit: () => void;
  onOpenWizard?: () => void;
  phase: WorkspacePhase;
}

export function FormulaBar({
  cellName,
  disabled,
  formula,
  onCancel,
  onChange,
  onCommit,
  onOpenWizard,
  phase,
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
    <Box as="form" aria-label="Formula bar" className="flex h-11 items-center gap-2 border-b border-slate-200 bg-white px-4" onSubmit={handleSubmit}>
      <TextInput
        aria-label="Selected cell"
        className="w-20 text-center font-mono text-xs font-bold text-slate-800 bg-slate-50 border-slate-200"
        disabled={disabled}
        readOnly
        value={cellName}
      />
      <Inline gap="sm" className="min-w-0 flex-1">
        <button
          type="button"
          onClick={onOpenWizard}
          title="Insert Function Wizard (fx)"
          disabled={disabled}
          className="flex h-8 shrink-0 items-center gap-1 rounded-md border border-slate-200 bg-slate-50 px-2.5 text-slate-600 transition-colors hover:bg-blue-50 hover:text-blue-600 focus:outline-hidden"
        >
          <Icon name="function" size="xs" />
          <span className="text-xs font-bold font-serif italic">fx</span>
        </button>
        <TextInput
          aria-label="Formula input"
          className="font-mono text-xs"
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLInputElement>) => onChange(event.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={phase === 'empty' ? 'Create a sheet to start editing' : 'Enter a value or formula (=SUM, =IF, ...)'}
          value={formula}
        />
      </Inline>
      <Inline gap="xs" className="shrink-0">
        <Button
          aria-label="Cancel formula edit"
          disabled={disabled}
          icon="x"
          iconOnly
          onClick={onCancel}
          size="sm"
          variant="ghost"
        />
        <Button
          aria-label="Apply formula"
          disabled={disabled}
          icon="check"
          iconOnly
          onClick={onCommit}
          size="sm"
          variant="soft"
        />
        <Inline gap="xs" className="hidden pl-2 lg:flex">
          <Text size="xs" tone="subtle">Apply</Text>
          <Kbd>Enter</Kbd>
        </Inline>
      </Inline>
    </Box>
  );
}
