import { AuthDomain } from './auth/domain';
import type { AuthOptions, AuthSession } from './auth/contract';
import type { IdentityService } from './identity/contract';
import type { UserAdministrationService } from './users/contract';
import { ApplicationRuntime, type WorkbooksActions } from './internal/runtime';
import { SdkError } from './error';

export interface SpreadsheetSdk {
  readonly auth: AuthSession;
  readonly identity: IdentityService;
  readonly users: UserAdministrationService;
  readonly workbooks: WorkbooksActions;
  dispose(): Promise<void>;
}
const runtimes = new WeakMap<SpreadsheetSdk, ApplicationRuntime>();
/** SDK-internal lookup; absent from package exports. */
export function runtimeFor(sdk: SpreadsheetSdk): ApplicationRuntime {
  const runtime = runtimes.get(sdk);
  if (!runtime) throw new SdkError('RUNTIME_DISPOSED', 'runtime', 'SDK 实例不存在。', '请使用 createSpreadsheetSdk 创建实例。');
  return runtime;
}
export function createSpreadsheetSdk(options: AuthOptions = {}): SpreadsheetSdk {
  const auth = new AuthDomain(options);
  const runtime = new ApplicationRuntime(auth);
  const releaseSdk = runtime.acquire();
  const sdk: SpreadsheetSdk = Object.freeze({
    auth: auth.session, identity: auth.identity, users: auth.users, get workbooks() { return runtime.catalog; },
    dispose: async () => { releaseSdk(); try { await runtime.dispose(); } finally { auth.dispose(); runtimes.delete(sdk); } },
  });
  runtimes.set(sdk, runtime);
  return sdk;
}
