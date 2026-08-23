import { createFormulaError, type FormulaError } from './values';

/**
 * Formula capabilities which are not safe to expose until their complete
 * calculation and spill semantics are implemented.
 *
 * Function names are used as capability ids deliberately. This keeps the
 * evaluator boundary data-only and avoids a second mapping that could drift
 * from the function registry.
 */
export type FormulaCapability = 'GROUPBY' | 'PIVOTBY';

export interface FormulaCapabilities {
  /** Return whether a gated formula capability may execute. */
  readonly isEnabled: (capability: FormulaCapability) => boolean;
  /** Optional diagnostic supplied by the host when a capability is disabled. */
  readonly reason?: (capability: FormulaCapability) => string | undefined;
}

const DEFAULT_DISABLED_REASON = 'calculation capability is not implemented';

/**
 * Formula functions are fail-closed by default. A host must explicitly inject
 * a provider before GROUPBY/PIVOTBY can reach their implementation.
 */
export const DEFAULT_FORMULA_CAPABILITIES: FormulaCapabilities = Object.freeze({
  isEnabled: () => false,
  reason: () => DEFAULT_DISABLED_REASON,
});

export interface FormulaCapabilityOptions {
  readonly enabled?: Iterable<FormulaCapability>;
  readonly reasons?: Partial<Record<FormulaCapability, string>>;
}

/** Create an immutable, explicit capability provider for a formula host. */
export function createFormulaCapabilities(options: FormulaCapabilityOptions = {}): FormulaCapabilities {
  const enabled = new Set(options.enabled ?? []);
  const reasons = { ...options.reasons };
  return {
    isEnabled: (capability) => enabled.has(capability),
    reason: (capability) => reasons[capability] ?? DEFAULT_DISABLED_REASON,
  };
}

const FUNCTION_CAPABILITIES: Readonly<Record<string, FormulaCapability>> = Object.freeze({
  GROUPBY: 'GROUPBY',
  PIVOTBY: 'PIVOTBY',
});

export function getFormulaCapability(name: string): FormulaCapability | undefined {
  return FUNCTION_CAPABILITIES[name.trim().toUpperCase()];
}

export function isFormulaCapabilityEnabled(
  capabilities: FormulaCapabilities | undefined,
  capability: FormulaCapability,
): boolean {
  return (capabilities ?? DEFAULT_FORMULA_CAPABILITIES).isEnabled(capability);
}

export function createFormulaCapabilityError(
  name: string,
  capability: FormulaCapability,
  capabilities?: FormulaCapabilities,
): FormulaError {
  const reason = (capabilities ?? DEFAULT_FORMULA_CAPABILITIES).reason?.(capability) ?? DEFAULT_DISABLED_REASON;
  return createFormulaError(
    '#BLOCKED!',
    `${name.trim().toUpperCase()} is disabled: ${reason}`,
  );
}
