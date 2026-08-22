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
  phase: WorkspacePhase;
}

export function FormulaBar({ cellName, disabled, formula, onCancel, onChange, onCommit, phase }: FormulaBarProps) {
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
    <Box as="form" aria-label="Formula bar" className="flex h-12 items-center gap-2 px-5" onSubmit={handleSubmit}>
      <TextInput aria-label="Selected cell" className="w-20 text-center font-mono text-xs font-semibold" disabled={disabled} readOnly value={cellName} />
      <Inline gap="sm" className="min-w-0 flex-1">
        <Inline gap="xs" className="h-9 shrink-0 rounded-lg border border-line bg-slate-50 px-3 text-muted">
          <Icon name="function" size="sm" />
          <Text size="sm" weight="semibold" tone="muted">fx</Text>
        </Inline>
        <TextInput
          aria-label="Formula input"
          className="font-mono text-xs"
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLInputElement>) => onChange(event.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={phase === 'empty' ? 'Create a sheet to start editing' : 'Enter a value or formula'}
          value={formula}
        />
      </Inline>
      <Inline gap="xs" className="shrink-0">
        <Button aria-label="Cancel formula edit" disabled={disabled} icon="x" iconOnly onClick={onCancel} size="sm" variant="ghost" />
        <Button aria-label="Apply formula" disabled={disabled} icon="check" iconOnly onClick={onCommit} size="sm" variant="soft" />
        <Inline gap="xs" className="hidden pl-2 lg:flex">
          <Text size="xs" tone="subtle">Apply</Text>
          <Kbd>Enter</Kbd>
        </Inline>
      </Inline>
    </Box>
  );
}
