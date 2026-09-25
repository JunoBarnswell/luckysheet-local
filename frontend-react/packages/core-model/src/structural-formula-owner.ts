import type { RangeRef } from './index';

export type StructuralFormulaRuleField = 'value1' | 'value2' | 'formula1' | 'formula2' | 'listSource.formula';

export type StructuralFormulaRule = {
  id: string;
  sheetId: string;
  ranges: RangeRef[];
  type?: string;
  operator?: string;
  value1?: string | number;
  value2?: string | number;
  formula1?: string;
  formula2?: string;
  listSource?: { kind: 'values'; values: string[] } | { kind: 'range'; range: RangeRef } | { kind: 'formula'; formula: string };
};

export function structuralRuleFormulaFields(rule: StructuralFormulaRule): Map<StructuralFormulaRuleField, string> {
  const formulas = new Map<StructuralFormulaRuleField, string>();
  if (rule.operator === 'formula' && typeof rule.value1 === 'string') formulas.set('value1', rule.value1);
  else {
    if (typeof rule.value1 === 'string' && rule.value1.trim().startsWith('=')) formulas.set('value1', rule.value1);
    if (typeof rule.value2 === 'string' && rule.value2.trim().startsWith('=')) formulas.set('value2', rule.value2);
  }
  if (rule.formula1 && (rule.formula1.trim().startsWith('=') || rule.operator === 'formula' || rule.type === 'custom')) {
    formulas.set('formula1', rule.formula1);
  }
  if (rule.formula2 && (rule.formula2.trim().startsWith('=') || rule.type === 'custom')) formulas.set('formula2', rule.formula2);
  if (rule.listSource?.kind === 'formula') formulas.set('listSource.formula', rule.listSource.formula);
  return formulas;
}
