import type { AuthContext } from '../auth/contract';

/** A read-only projection of the identity verified by the spreadsheet server. */
export interface IdentitySnapshot {
  readonly phase: 'unverified' | 'verified' | 'retired';
  readonly context: AuthContext | null;
  readonly displayName: string | null;
}
export interface IdentityService {
  getSnapshot(): IdentitySnapshot;
  subscribe(listener: () => void): () => void;
}
