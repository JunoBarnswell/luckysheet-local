import {
  DEFAULT_FACADE_DSL_LIMITS,
  parseFacadeScript,
  type FacadeDslLimits,
  type FacadeProgram,
} from './dsl';

/**
 * Limits for the data-only Facade DSL.  The old implementation attempted to
 * secure JavaScript by scanning source text for forbidden regular expressions;
 * that is not a security boundary.  The parser is now the boundary and only
 * these finite resource limits remain here.
 */
export interface SandboxPolicy extends FacadeDslLimits {
  maxDurationMs: number;
}

export const DEFAULT_SANDBOX_POLICY: SandboxPolicy = {
  ...DEFAULT_FACADE_DSL_LIMITS,
  maxDurationMs: 5000,
};

export class ScriptSandbox {
  constructor(private readonly policy: SandboxPolicy = DEFAULT_SANDBOX_POLICY) {}

  parse(source: string): FacadeProgram {
    try {
      return parseFacadeScript(source, this.policy);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(`Automation script blocked by the Facade DSL: ${message}`);
    }
  }

  /** Validate the complete program before any mutation is submitted. */
  assertAllowed(source: string): void {
    this.parse(source);
  }

  getLimits(): FacadeDslLimits {
    return this.policy;
  }

  getTimeoutMs(): number {
    return this.policy.maxDurationMs;
  }
}
