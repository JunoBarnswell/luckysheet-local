export type WorkbookCalculationContextAction = 'rebuild' | 'sync-defined-names' | 'sync-tables';

export interface WorkbookCalculationContextEffect {
  readonly kind: 'calculation-context';
  readonly action: WorkbookCalculationContextAction;
}

export const CALCULATION_CONTEXT_EFFECTS = {
  rebuild: { kind: 'calculation-context', action: 'rebuild' },
  syncDefinedNames: { kind: 'calculation-context', action: 'sync-defined-names' },
  syncTables: { kind: 'calculation-context', action: 'sync-tables' },
} as const satisfies Record<string, WorkbookCalculationContextEffect>;

export function isWorkbookCalculationContextEffect(value: unknown): value is WorkbookCalculationContextEffect {
  if (typeof value !== 'object' || value === null) return false;
  const effect = value as Record<string, unknown>;
  return Object.keys(effect).length === 2
    && effect.kind === 'calculation-context'
    && (effect.action === 'rebuild' || effect.action === 'sync-defined-names' || effect.action === 'sync-tables');
}
