import type { AuthContext } from './contract';
import { SdkError } from '../error';

export function authContext(value: unknown, principal: string): AuthContext {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw invalid();
  const raw = value as Record<string, unknown>;
  for (const key of ['authority', 'subject', 'principal', 'scopeId', 'sessionId', 'contextId']) {
    if (typeof raw[key] !== 'string' || !raw[key].trim()) throw invalid();
  }
  for (const key of ['tenantId', 'appCode', 'employmentId']) {
    if (raw[key] !== null && (typeof raw[key] !== 'string' || !raw[key].trim())) throw invalid();
  }
  if (raw.principal !== principal || !Number.isSafeInteger(raw.contextVersion) || (raw.contextVersion as number) < 0
    || ((raw.tenantId === null) !== (raw.appCode === null))
    || ((raw.tenantId === null) !== (raw.employmentId === null))) throw invalid();
  return Object.freeze({ authority: raw.authority as string, subject: raw.subject as string,
    principal, scopeId: raw.scopeId as string, sessionId: raw.sessionId as string,
    tenantId: raw.tenantId as string | null, appCode: raw.appCode as string | null,
    employmentId: raw.employmentId as string | null, contextVersion: raw.contextVersion as number,
    contextId: raw.contextId as string });
}
function invalid(): SdkError {
  return new SdkError('CONTRACT_INVALID', 'auth.context', '服务端身份上下文无效。', '请修复已验证的身份上下文契约后重新登录。');
}
