import type { PivotDefinition, PivotResultTree } from '@react-sheets/core-model';
import type { PivotRevisionKey, PivotTaskControl } from './engine';
import type { PivotSourceIndex } from './source-index';

export const PIVOT_TASK_PROTOCOL = 'react-sheets/pivot-task' as const;
export const PIVOT_TASK_VERSION = 1 as const;

export type PivotTaskErrorCode =
  | 'PIVOT_SOURCE_INVALID'
  | 'PIVOT_SOURCE_UNAVAILABLE'
  | 'PIVOT_MEMBER_LIMIT_EXCEEDED'
  | 'PIVOT_TARGET_COLLISION'
  | 'PIVOT_TARGET_BOUNDS_EXCEEDED'
  | 'PIVOT_RESULT_LIMIT_EXCEEDED'
  | 'PIVOT_TASK_CANCELLED'
  | 'PIVOT_TASK_TIMEOUT'
  | 'PIVOT_TASK_PROTOCOL_ERROR'
  | 'PIVOT_TASK_REVISION_MISMATCH'
  | 'PIVOT_PERMISSION_DENIED'
  | 'PIVOT_TASK_FAILED';

export type PivotTaskRecovery = 'fix-source' | 'change-layout' | 'change-target' | 'retry';

export interface PivotTaskError {
  code: PivotTaskErrorCode;
  message: string;
  pivotId: string;
  sourceIdentity: string;
  sourceRevision: string;
  recovery: PivotTaskRecovery;
}

interface PivotTaskEnvelope {
  protocol: typeof PIVOT_TASK_PROTOCOL;
  version: typeof PIVOT_TASK_VERSION;
  taskId: string;
  generation: number;
}

export interface PivotSourceRegisterRequest extends PivotTaskEnvelope {
  kind: 'source-register';
  sourceIdentity: string;
  sourceRevision: string;
  source: PivotSourceIndex;
}

export interface PivotSourceReleaseRequest extends PivotTaskEnvelope {
  kind: 'source-release';
  sourceIdentity: string;
  sourceRevision: string;
}

export interface PivotCalculateRequest extends PivotTaskEnvelope {
  kind: 'calculate';
  sourceIdentity: string;
  definition: PivotDefinition;
  controls: PivotTaskControl[];
  revisions: PivotRevisionKey;
  targetBounds: { rowCount: number; columnCount: number };
}

export interface PivotTaskCancelRequest extends PivotTaskEnvelope {
  kind: 'cancel';
}

export type PivotTaskRequest = PivotSourceRegisterRequest | PivotSourceReleaseRequest | PivotCalculateRequest | PivotTaskCancelRequest;

export interface PivotTaskAcceptedResult extends PivotTaskEnvelope {
  status: 'accepted';
  sourceIdentity: string;
  sourceRevision: string;
}

export interface PivotTaskCompletedResult extends PivotTaskEnvelope {
  status: 'completed';
  sourceIdentity: string;
  sourceRevision: string;
  result: PivotResultTree;
}

export interface PivotTaskCancelledResult extends PivotTaskEnvelope {
  status: 'cancelled';
}

export interface PivotTaskFailedResult extends PivotTaskEnvelope {
  status: 'failed';
  error: PivotTaskError;
}

export type PivotTaskResult = PivotTaskAcceptedResult | PivotTaskCompletedResult | PivotTaskCancelledResult | PivotTaskFailedResult;

export function createPivotSourceRegisterRequest(
  taskId: string,
  generation: number,
  sourceIdentity: string,
  sourceRevision: string,
  source: PivotSourceIndex,
): PivotSourceRegisterRequest {
  return { protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation, kind: 'source-register', sourceIdentity, sourceRevision, source };
}

export function createPivotCalculateRequest(
  taskId: string,
  generation: number,
  sourceIdentity: string,
  definition: PivotDefinition,
  controls: PivotTaskControl[],
  revisions: PivotRevisionKey,
  targetBounds: { rowCount: number; columnCount: number },
): PivotCalculateRequest {
  return { protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation, kind: 'calculate', sourceIdentity, definition, controls, revisions, targetBounds };
}

export function createPivotTaskCancelRequest(taskId: string, generation: number): PivotTaskCancelRequest {
  return { protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation, kind: 'cancel' };
}

export function createPivotSourceReleaseRequest(
  taskId: string,
  generation: number,
  sourceIdentity: string,
  sourceRevision: string,
): PivotSourceReleaseRequest {
  return { protocol: PIVOT_TASK_PROTOCOL, version: PIVOT_TASK_VERSION, taskId, generation, kind: 'source-release', sourceIdentity, sourceRevision };
}

export function assertPivotTaskRequest(value: unknown): asserts value is PivotTaskRequest {
  if (!isRecord(value) || value.protocol !== PIVOT_TASK_PROTOCOL || value.version !== PIVOT_TASK_VERSION
    || typeof value.taskId !== 'string' || value.taskId.length === 0
    || !Number.isSafeInteger(value.generation) || Number(value.generation) < 0
    || !['source-register', 'source-release', 'calculate', 'cancel'].includes(String(value.kind))) {
    throw new Error('Pivot task request protocol is invalid');
  }
  if (value.kind !== 'cancel') {
    if (typeof value.sourceIdentity !== 'string' || value.sourceIdentity.length === 0) {
      throw new Error('Pivot task source identity is invalid');
    }
  }
  if (value.kind === 'source-register' || value.kind === 'source-release') {
    if (typeof value.sourceRevision !== 'string' || value.sourceRevision.length === 0) {
      throw new Error('Pivot source revision is invalid');
    }
  }
  if (value.kind === 'source-register' && !isRecord(value.source)) {
    throw new Error('Pivot source registration is invalid');
  }
  if (value.kind === 'calculate') {
    if (!isRecord(value.targetBounds)
      || !Number.isSafeInteger(value.targetBounds.rowCount) || Number(value.targetBounds.rowCount) <= 0
      || !Number.isSafeInteger(value.targetBounds.columnCount) || Number(value.targetBounds.columnCount) <= 0) {
      throw new Error('Pivot task target bounds are invalid');
    }
    if (!isRecord(value.definition) || typeof value.definition.id !== 'string' || value.definition.id.length === 0
      || !isRecord(value.revisions) || typeof value.revisions.sourceRevision !== 'string' || value.revisions.sourceRevision.length === 0
      || !Array.isArray(value.controls)) {
      throw new Error('Pivot calculation payload is invalid');
    }
  }
}

export function assertPivotTaskResult(value: unknown): asserts value is PivotTaskResult {
  if (!isRecord(value) || value.protocol !== PIVOT_TASK_PROTOCOL || value.version !== PIVOT_TASK_VERSION
    || typeof value.taskId !== 'string' || value.taskId.length === 0
    || !Number.isSafeInteger(value.generation) || Number(value.generation) < 0
    || !['accepted', 'completed', 'cancelled', 'failed'].includes(String(value.status))) {
    throw new Error('Pivot task result protocol is invalid');
  }
  if ((value.status === 'accepted' || value.status === 'completed')
    && (typeof value.sourceIdentity !== 'string' || value.sourceIdentity.length === 0
      || typeof value.sourceRevision !== 'string' || value.sourceRevision.length === 0)) {
    throw new Error('Pivot task result source identity is invalid');
  }
  if (value.status === 'failed') {
    if (!isRecord(value.error) || typeof value.error.code !== 'string' || typeof value.error.message !== 'string'
      || typeof value.error.pivotId !== 'string' || typeof value.error.sourceIdentity !== 'string'
      || typeof value.error.sourceRevision !== 'string' || typeof value.error.recovery !== 'string') {
      throw new Error('Pivot task failure is invalid');
    }
  }
}

/**
 * A worker reply is only useful when it is the reply to the exact request
 * which is pending.  Checking the envelope alone lets an accepted source
 * registration masquerade as a completed calculation (or vice versa), and
 * lets an unrelated source identity poison the session cache.
 */
export function assertPivotTaskResultForRequest(
  value: unknown,
  request: Exclude<PivotTaskRequest, { kind: 'cancel' }>,
): asserts value is PivotTaskResult {
  assertPivotTaskResult(value);
  if (value.taskId !== request.taskId || value.generation !== request.generation) {
    throw new Error('Pivot worker result does not match the pending task');
  }
  const expectedSourceIdentity = request.sourceIdentity;
  const expectedSourceRevision = request.kind === 'calculate' ? request.revisions.sourceRevision : request.sourceRevision;
  if (value.status === 'accepted') {
    if (request.kind === 'calculate') throw new Error('Pivot calculation cannot return an accepted result');
    if (value.sourceIdentity !== expectedSourceIdentity || value.sourceRevision !== expectedSourceRevision) {
      throw new Error('Pivot worker accepted an unexpected source revision');
    }
    return;
  }
  if (value.status === 'completed') {
    if (request.kind !== 'calculate') throw new Error('Pivot source operation cannot return a completed result');
    if (value.sourceIdentity !== expectedSourceIdentity || value.sourceRevision !== expectedSourceRevision) {
      throw new Error('Pivot worker completed an unexpected source revision');
    }
    return;
  }
  if (value.status === 'failed') {
    if (value.error.sourceIdentity !== expectedSourceIdentity || value.error.sourceRevision !== expectedSourceRevision) {
      throw new Error('Pivot worker failure does not identify the pending source revision');
    }
    if (request.kind === 'calculate' && value.error.pivotId !== request.definition.id) {
      throw new Error('Pivot worker failure does not identify the pending PivotTable');
    }
  }
}

export function pivotTaskFailure(
  request: Pick<PivotCalculateRequest, 'taskId' | 'generation' | 'sourceIdentity' | 'definition' | 'revisions'>,
  error: unknown,
  overrideCode?: PivotTaskErrorCode,
): PivotTaskFailedResult {
  const message = error instanceof Error ? error.message : 'Pivot task failed';
  const code = overrideCode ?? classifyPivotTaskError(message);
  return {
    protocol: PIVOT_TASK_PROTOCOL,
    version: PIVOT_TASK_VERSION,
    taskId: request.taskId,
    generation: request.generation,
    status: 'failed',
    error: {
      code,
      message,
      pivotId: request.definition.id,
      sourceIdentity: request.sourceIdentity,
      sourceRevision: request.revisions.sourceRevision,
      recovery: recoveryFor(code),
    },
  };
}

function classifyPivotTaskError(message: string): PivotTaskErrorCode {
  if (/member domain exceeds/i.test(message)) return 'PIVOT_MEMBER_LIMIT_EXCEEDED';
  if (/collision/i.test(message)) return 'PIVOT_TARGET_COLLISION';
  if (/worksheet boundary|worksheet-bounds|exceeds the destination/i.test(message)) return 'PIVOT_TARGET_BOUNDS_EXCEEDED';
  if (/result (?:cell|provenance) limit exceeded/i.test(message)) return 'PIVOT_RESULT_LIMIT_EXCEEDED';
  if (/source.*unavailable|unknown.*source|no worksheet range/i.test(message)) return 'PIVOT_SOURCE_UNAVAILABLE';
  if (/source|field|header|relationship|spill/i.test(message)) return 'PIVOT_SOURCE_INVALID';
  return 'PIVOT_TASK_FAILED';
}

function recoveryFor(code: PivotTaskErrorCode): PivotTaskRecovery {
  if (code === 'PIVOT_TARGET_COLLISION' || code === 'PIVOT_TARGET_BOUNDS_EXCEEDED') return 'change-target';
  if (code === 'PIVOT_MEMBER_LIMIT_EXCEEDED' || code === 'PIVOT_RESULT_LIMIT_EXCEEDED') return 'change-layout';
  if (code === 'PIVOT_SOURCE_INVALID' || code === 'PIVOT_SOURCE_UNAVAILABLE') return 'fix-source';
  return 'retry';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
