import { SdkError } from '../error';

/** Private, tab-scoped credential source. The public API never returns a capability. */
export class GuestCapabilityDomain {
  private token: string | null = null;
  private path: string | null = null;
  private unitId: string | null = null;
  private disposed = false;
  constructor(private readonly onRetired: () => void) {
    if (typeof window === 'undefined') return;
    const fragment = new URLSearchParams(window.location.hash.slice(1));
    const capability = fragment.get('share')?.trim();
    if (fragment.has('share') || new URLSearchParams(window.location.search).has('share')) {
      fragment.delete('share');
      const clean = new URL(window.location.href);
      clean.hash = fragment.toString();
      clean.searchParams.delete('share');
      window.history.replaceState(window.history.state, '', clean.pathname + clean.search + clean.hash);
    }
    const match = /^\/workbooks\/([^/]+)$/.exec(window.location.pathname);
    if (!match) {
      if (capability) throw new SdkError('INVALID_ARGUMENT', 'auth.guest.capture', 'A share capability requires a workbook route.', 'Use the original workbook share link.');
      return;
    }
    try {
      this.path = window.location.pathname;
      this.unitId = decodeURIComponent(match[1]!);
      if (!this.unitId.trim()) throw new Error('Workbook identity is empty');
      if (capability) window.sessionStorage.setItem(`share:${this.path}`, capability);
      this.token = capability || window.sessionStorage.getItem(`share:${this.path}`);
    } catch (cause) {
      this.token = null;
      throw new SdkError('SERVICE_UNAVAILABLE', 'auth.guest.capture', 'The tab-scoped share credential could not be captured.', 'Enable tab storage and reopen the original share link.', { cause });
    }
  }
  get scope(): string | null { return this.read() ? this.unitId : null; }
  read(): string | null {
    if (this.disposed) throw new SdkError('RUNTIME_DISPOSED', 'auth.guest', 'The share credential source is retired.', 'Create a new SDK with the original share link.');
    if (this.path && typeof window !== 'undefined' && window.location.pathname !== this.path) this.revoke();
    return this.token;
  }
  revoke(): void {
    const hadCapability = this.token !== null;
    this.token = null;
    try { if (this.path && typeof window !== 'undefined') window.sessionStorage.removeItem(`share:${this.path}`); }
    catch (cause) {
      throw new SdkError('SERVICE_UNAVAILABLE', 'auth.guest.revoke', 'The tab-scoped share credential could not be cleared.', 'Close this tab before reopening a share link.', { cause });
    } finally { if (hadCapability) this.onRetired(); }
  }
  dispose(): void { this.disposed = true; this.token = null; }
}
