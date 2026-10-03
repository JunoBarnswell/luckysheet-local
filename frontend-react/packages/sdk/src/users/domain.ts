import type { AuthSnapshot } from '../auth/contract';
import { SdkError } from '../error';
import type { LocalUser, UserAdministrationService } from './contract';

interface AdministrationPort {
  snapshot(): AuthSnapshot;
  request(path: string, method: string, body?: object): Promise<Response>;
}
function record(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}
/** Administrative actions use the credential owner's contextual transport and real Java authority. */
export class UserAdministrationDomain {
  private retired = false;
  readonly service: UserAdministrationService;
  constructor(private readonly port: AdministrationPort) {
    this.service = Object.freeze({
      listUsers: () => this.listUsers(),
      createUser: input => this.perform('users.createUser', async request => {
        if (!input || ![input.username, input.displayName, input.password].every(value => typeof value === 'string' && value.trim())) throw this.invalid('users.createUser');
        await request('/api/admin/users', 'POST', { username: input.username, displayName: input.displayName, password: input.password });
      }),
      setUserEnabled: (userId, enabled) => this.perform('users.setUserEnabled', async request => {
        if (typeof userId !== 'string' || !userId.trim() || typeof enabled !== 'boolean') throw this.invalid('users.setUserEnabled');
        await request(`/api/admin/users/${encodeURIComponent(userId)}`, 'PATCH', { enabled });
      }),
      resetPassword: (userId, password) => this.perform('users.resetPassword', async request => {
        if (typeof userId !== 'string' || !userId.trim() || typeof password !== 'string' || !password.trim()) throw this.invalid('users.resetPassword');
        await request(`/api/admin/users/${encodeURIComponent(userId)}/password`, 'POST', { password });
      }),
    } satisfies UserAdministrationService);
  }
  private invalid(operation: string): SdkError {
    return new SdkError('INVALID_ARGUMENT', operation, 'User administration input is invalid.', 'Provide a valid user identity and required fields.');
  }
  private async perform<T>(operation: string, action: (request: AdministrationPort['request']) => Promise<T>): Promise<T> {
    try {
      if (this.retired) throw new SdkError('RUNTIME_DISPOSED', operation, 'User administration is retired.', 'Create a new SDK.');
      const initial = this.port.snapshot();
      const assert = () => {
        if (this.retired) throw new SdkError('RUNTIME_DISPOSED', operation, 'User administration is retired.', 'Create a new SDK.');
        const current = this.port.snapshot();
        if (initial.context?.contextId !== current.context?.contextId) throw new SdkError('STALE_OPERATION', operation, 'The administrative identity changed.', 'Use the current verified identity.');
        if (current.phase !== 'authenticated' || !current.capabilities.canManageUsers) throw new SdkError('FORBIDDEN', operation, 'User administration requires an administrator.', 'Sign in with user administration permission.');
      };
      assert();
      const request: AdministrationPort['request'] = async (path, method, body) => {
        assert();
        const response = await this.port.request(path, method, body);
        assert();
        return response;
      };
      const result = await action(request);
      assert();
      return result;
    } catch (cause) {
      if (cause instanceof SdkError) throw cause;
      throw new SdkError('SERVICE_UNAVAILABLE', operation, 'User administration failed.', 'Check the server connection and retry using the current identity.', { cause });
    }
  }
  private listUsers(): Promise<readonly LocalUser[]> {
    return this.perform('users.listUsers', async request => {
      let value: unknown;
      try { value = await (await request('/api/admin/users', 'GET')).json(); }
      catch (cause) {
        if (cause instanceof SdkError) throw cause;
        throw new SdkError('CONTRACT_INVALID', 'users.listUsers', 'The user list is not valid JSON.', 'Repair the server response contract.', { cause });
      }
      if (!Array.isArray(value) || value.some(user => !record(user) || typeof user.id !== 'string' || !user.id
        || typeof user.username !== 'string' || typeof user.displayName !== 'string' || typeof user.enabled !== 'boolean' || typeof user.admin !== 'boolean')) {
        throw new SdkError('CONTRACT_INVALID', 'users.listUsers', 'The user list contract is invalid.', 'Repair the server user list contract.');
      }
      return Object.freeze(value.map(user => Object.freeze({ id: user.id, username: user.username, displayName: user.displayName, enabled: user.enabled, admin: user.admin })));
    });
  }
  dispose(): void { this.retired = true; }
}
