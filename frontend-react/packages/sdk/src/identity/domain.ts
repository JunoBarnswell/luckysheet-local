import type { AuthContext } from '../auth/contract';
import { SdkError } from '../error';
import type { IdentityService, IdentitySnapshot } from './contract';

/** Owns publication of verified identity; credentials and administration are separate domains. */
export class IdentityDomain {
  private snapshot: IdentitySnapshot = Object.freeze({ phase: 'unverified', context: null, displayName: null });
  private readonly listeners = new Set<() => void>();
  private retired = false;
  readonly service: IdentityService = Object.freeze({
    getSnapshot: () => this.snapshot,
    subscribe: (listener: () => void) => {
      if (this.retired) throw new SdkError('RUNTIME_DISPOSED', 'identity.subscribe', 'Identity is retired.', 'Create a new SDK.');
      this.listeners.add(listener);
      return () => { this.listeners.delete(listener); };
    },
  });
  publish(context: AuthContext | null, displayName: string | null): void {
    if (this.retired) return;
    if (context?.contextId === this.snapshot.context?.contextId && displayName === this.snapshot.displayName) return;
    this.snapshot = Object.freeze({ phase: context ? 'verified' : 'unverified', context, displayName });
    for (const listener of this.listeners) listener();
  }
  dispose(): void {
    if (this.retired) return;
    this.retired = true;
    this.snapshot = Object.freeze({ phase: 'retired', context: null, displayName: null });
    for (const listener of this.listeners) listener();
    this.listeners.clear();
  }
}
