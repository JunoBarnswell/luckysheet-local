import { FormulaBar as FormulaBarShell, type FormulaBarProps as ShellFormulaBarProps } from '@react-sheets/ui-system';
import type { AppPhase } from '@react-sheets/spreadsheet-app';
import { formulaBarLabels, type Locale } from '../i18n';

export interface FormulaBarProps {
  cellName: string;
  disabled: boolean;
  formula: string;
  locale: Locale;
  onCancel: () => void;
  onBeginEdit: () => void;
  onChange: (value: string) => void;
  onCommit: () => void;
  composing?: boolean;
  onCaretChange?: (start: number, end: number) => void;
  onCompositionStart?: () => void;
  onCompositionUpdate?: (text: string) => void;
  onCompositionEnd?: () => void;
  onNameBoxCommit?: (value: string) => void;
  onOpenNameManager?: () => void;
  onOpenWizard?: () => void;
  phase: AppPhase;
}

export function FormulaBar({ locale, phase, onBeginEdit, ...props }: FormulaBarProps) {
  const shellProps: ShellFormulaBarProps = {
    ...props,
    onFocusFormula: onBeginEdit,
    labels: formulaBarLabels(locale, phase),
  };
  return <FormulaBarShell {...shellProps} />;
}
