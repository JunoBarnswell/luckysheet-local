import type {
  DrawingAnchor,
  DrawingObject,
  DrawingPayload,
  DrawingTransform,
  RangeRef,
  SheetId,
} from '@react-sheets/core-model';

export interface InsertIdentity {
  objectId: string;
  payloadId: string;
}

export interface InsertResult {
  kind: string;
  sheetId: SheetId;
  commandId: string;
  createdObjectIds: readonly string[];
  affectedRanges: readonly RangeRef[];
  commandResult: unknown;
}

export interface InsertRequest {
  kind: string;
  commandId: string;
  sheetId: SheetId;
  params: Record<string, unknown>;
  createdObjectIds: readonly string[];
}

export interface DrawingInsertRequest extends Omit<InsertRequest, 'kind' | 'params' | 'createdObjectIds'> {
  drawing: DrawingObject;
  payload: DrawingPayload;
  extraParams?: Record<string, unknown>;
}

export type InsertMutationRequest = InsertRequest;

/**
 * Owns the cross-domain insertion lifecycle. Domain reducers still own their
 * payload validation; this coordinator owns identity, default placement and
 * the post-commit activation boundary.
 */
export class InsertCoordinator {
  constructor(private readonly allocateId: (prefix: string) => string) {}

  allocateObjectId(prefix: string): string {
    return this.allocateId(prefix);
  }

  allocateIdentity(objectPrefix = 'drawing', payloadPrefix = 'drawing'): InsertIdentity {
    return {
      objectId: this.allocateId(objectPrefix),
      payloadId: this.allocateId(payloadPrefix),
    };
  }

  defaultPlacement(anchor: DrawingAnchor = { kind: 'absolute' }): { anchor: DrawingAnchor; transform: DrawingTransform; zIndex: number } {
    return {
      anchor: structuredClone(anchor),
      transform: { x: 96, y: 96, width: 480, height: 280, rotation: 0 },
      zIndex: 0,
    };
  }

  commitDrawing(
    request: DrawingInsertRequest,
    execute: (commandId: string, params: Record<string, unknown>) => unknown,
    activate: (result: InsertResult) => void,
  ): InsertResult {
    const params = {
      sheetId: request.sheetId,
      ...(request.extraParams ?? {}),
      drawing: structuredClone(request.drawing),
      payload: structuredClone(request.payload),
    };
    return this.commitMutation({
      kind: request.drawing.kind,
      commandId: request.commandId,
      sheetId: request.sheetId,
      params,
      createdObjectIds: [request.drawing.id],
    }, execute, activate);
  }

  commitMutation(
    request: InsertMutationRequest,
    execute: (commandId: string, params: Record<string, unknown>) => unknown,
    activate: (result: InsertResult) => void,
  ): InsertResult {
    const commandResult = execute(request.commandId, structuredClone(request.params));
    const affectedRanges = isAffectedRangeResult(commandResult) ? commandResult.affectedRanges : [];
    const result: InsertResult = {
      kind: request.kind,
      sheetId: request.sheetId,
      commandId: request.commandId,
      createdObjectIds: structuredClone(request.createdObjectIds),
      affectedRanges: structuredClone(affectedRanges),
      commandResult,
    };
    activate(result);
    return result;
  }
}

function isAffectedRangeResult(value: unknown): value is { affectedRanges: RangeRef[] } {
  return Boolean(value && typeof value === 'object' && Array.isArray((value as { affectedRanges?: unknown }).affectedRanges));
}
