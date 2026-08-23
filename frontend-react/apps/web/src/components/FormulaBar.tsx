import type { ChangeEvent, FormEvent, KeyboardEvent } from 'react';
import { Box, Button, Inline, Kbd, Text, TextInput } from '@react-sheets/ui-system';
import type { WorkspacePhase } from '../state/workspace';
import type { Locale } from '../i18n';

export interface FormulaBarProps {
  cellName: string;
  disabled: boolean;
  formula: string;
  locale: Locale;
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
  locale,
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
    <Box as="form" aria-label="Formula bar" className="flex h-10 items-center gap-2 border-b border-slate-200 bg-white px-4" onSubmit={handleSubmit}>
      <TextInput
          aria-label={locale === 'zh-CN' ? '选中单元格' : 'Selected cell'}
        className="w-20 text-center font-mono text-xs font-bold text-slate-800 bg-slate-50 border-slate-200"
        disabled={disabled}
        readOnly
        value={cellName}
      />
      <Inline gap="sm" className="min-w-0 flex-1">
        <Button aria-label={locale === 'zh-CN' ? '插入函数向导' : 'Insert Function Wizard'} disabled={disabled} icon="function" onClick={onOpenWizard} size="sm" variant="outline">fx</Button>
        <TextInput
          aria-label={locale === 'zh-CN' ? '公式输入' : 'Formula input'}
          className="font-mono text-xs"
          disabled={disabled}
          onChange={(event: ChangeEvent<HTMLInputElement>) => onChange(event.target.value)}
          onKeyDown={handleKeyDown}
          placeholder={phase === 'empty' ? (locale === 'zh-CN' ? '创建工作表后开始编辑' : 'Create a sheet to start editing') : (locale === 'zh-CN' ? '输入值或公式 (=SUM, =IF, ...)' : 'Enter a value or formula (=SUM, =IF, ...)')}
          value={formula}
        />
      </Inline>
      <Inline gap="xs" className="shrink-0">
        <Button
          aria-label={locale === 'zh-CN' ? '取消公式编辑' : 'Cancel formula edit'}
          disabled={disabled}
          icon="x"
          iconOnly
          onClick={onCancel}
          size="sm"
          variant="ghost"
        />
        <Button
          aria-label={locale === 'zh-CN' ? '应用公式' : 'Apply formula'}
          disabled={disabled}
          icon="check"
          iconOnly
          onClick={onCommit}
          size="sm"
          variant="soft"
        />
        <Inline gap="xs" className="hidden pl-2 lg:flex">
          <Text size="xs" tone="subtle">{locale === 'zh-CN' ? '应用' : 'Apply'}</Text>
          <Kbd>Enter</Kbd>
        </Inline>
      </Inline>
    </Box>
  );
}
