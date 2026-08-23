export interface SandboxPolicy {
  maxDurationMs: number;
  allowNetwork: boolean;
  allowFileSystem: boolean;
  blockedPatterns: RegExp[];
}

export const DEFAULT_SANDBOX_POLICY: SandboxPolicy = {
  maxDurationMs: 5000,
  allowNetwork: false,
  allowFileSystem: false,
  blockedPatterns: [
    /\bfetch\s*\(/,
    /\bXMLHttpRequest\b/,
    /\brequire\s*\(/,
    /\bimport\s*\(/,
    /\bprocess\b/,
    /\bfs\b/,
    /\beval\s*\(/,
    /\bFunction\s*\(/,
  ],
};

/** 脚本沙箱 — 无任意 FS/DB、网络许可、超时、审计 */
export class ScriptSandbox {
  constructor(private readonly policy: SandboxPolicy = DEFAULT_SANDBOX_POLICY) {}

  assertAllowed(source: string): void {
    for (const pattern of this.policy.blockedPatterns) {
      if (pattern.test(source)) {
        throw new Error(`Script blocked by sandbox policy: ${pattern.source}`);
      }
    }
  }

  getTimeoutMs(): number {
    return this.policy.maxDurationMs;
  }
}
