/** M18 Capability flags — Solver/DAX/Python 仅评估入口，默认关闭 */
export type PlatformCapability =
  | 'solver'
  | 'dax'
  | 'python-runtime'
  | 'ai-copilot'
  | 'what-if'
  | 'groupby-pivotby';

export const DISABLED_STUB_FUNCTIONS = new Set(['GROUPBY', 'PIVOTBY']);

export interface CapabilityDescriptor {
  id: PlatformCapability;
  enabled: boolean;
  reason?: string;
  evaluateUrl?: string;
}

export class CapabilityRegistry {
  private readonly capabilities = new Map<PlatformCapability, CapabilityDescriptor>();

  constructor() {
    this.register({ id: 'what-if', enabled: true });
    this.register({
      id: 'groupby-pivotby',
      enabled: false,
      reason: 'GROUPBY/PIVOTBY are disabled until their full calculation and spill semantics are implemented',
    });
    this.register({ id: 'solver', enabled: false, reason: 'Requires external Solver engine' });
    this.register({ id: 'dax', enabled: false, reason: 'Requires Data Model runtime' });
    this.register({ id: 'python-runtime', enabled: false, reason: 'Requires Python sandbox service' });
    this.register({ id: 'ai-copilot', enabled: false, reason: 'Requires AI service' });
  }

  register(descriptor: CapabilityDescriptor): void {
    this.capabilities.set(descriptor.id, descriptor);
  }

  isEnabled(id: PlatformCapability): boolean {
    return this.capabilities.get(id)?.enabled ?? false;
  }

  get(id: PlatformCapability): CapabilityDescriptor | undefined {
    return this.capabilities.get(id);
  }

  list(): CapabilityDescriptor[] {
    return [...this.capabilities.values()];
  }

  /** 评估入口 — 返回是否可启用及原因 */
  evaluate(id: PlatformCapability): { canEnable: boolean; reason?: string } {
    const cap = this.capabilities.get(id);
    if (!cap) return { canEnable: false, reason: 'Unknown capability' };
    if (cap.enabled) return { canEnable: true };
    return { canEnable: false, reason: cap.reason };
  }

  evaluateFormulaFunction(name: string): { enabled: boolean; reason?: string } {
    const normalized = name.trim().toUpperCase();
    if (DISABLED_STUB_FUNCTIONS.has(normalized)) {
      return { enabled: false, reason: 'GROUPBY/PIVOTBY require full calculation and spill semantics' };
    }
    return { enabled: true };
  }
}

export function createDefaultCapabilityRegistry(): CapabilityRegistry {
  return new CapabilityRegistry();
}

export * from './what-if';
export { registerExtendedCommands } from './commands';
export * from './runtime';
