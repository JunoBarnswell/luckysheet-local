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
  if (value.kind === 'calculate') {
    if (!isRecord(value.targetBounds)
      || !Number.isSafeInteger(value.targetBounds.rowCount) || Number(value.targetBounds.rowCount) <= 0
      || !Number.isSafeInteger(value.targetBounds.columnCount) || Number(value.targetBounds.columnCount) <= 0) {
      throw new Error('Pivot task target bounds are invalid');
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
