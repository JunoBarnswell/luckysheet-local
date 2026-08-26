import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { CellHyperlink, CellNote, CommentReply, CommentThread, RangeRef } from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import {
  getCellHyperlink,
  removeCellHyperlink,
  setCellHyperlink,
  validateHyperlinkTarget,
} from './helpers';

function sheetRange(sheetId: string) {
  return [{ sheetId, startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }];
}

export interface NoteSetParams {
  sheetId: string;
  row: number;
  column: number;
  note: CellNote;
}

export interface NoteRemoveParams {
  sheetId: string;
  row: number;
  column: number;
}

export interface NoteVisibilityParams {
  sheetId: string;
  row: number;
  column: number;
  visible: boolean;
}

export interface CommentAddParams {
  sheetId: string;
  row: number;
  column: number;
  thread: CommentThread;
}

export interface CommentReplyParams {
  sheetId: string;
  threadId: string;
  reply: CommentReply;
}

export interface CommentResolveParams {
  sheetId: string;
  threadId: string;
  resolved: boolean;
  /** Stable event time supplied by the operation payload, never generated during replay. */
  resolvedAt?: string;
}

export interface CommentReplyRemoveParams {
  sheetId: string;
  threadId: string;
  replyId: string;
}

export interface CommentRemoveParams {
  sheetId: string;
  threadId: string;
}

export interface CommentUpdateParams {
  sheetId: string;
  threadId: string;
  row: number;
  column: number;
  text: string;
}

export interface HyperlinkSetParams {
  sheetId: string;
  row: number;
  column: number;
  hyperlink: CellHyperlink;
}

export interface HyperlinkRemoveParams {
  sheetId: string;
  row: number;
  column: number;
}

function cellRange(sheetId: string, row: number, column: number) {
  return [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
}

const isRecord = (value: unknown): value is Record<string, unknown> => typeof value === 'object' && value !== null;
const isNonEmptyString = (value: unknown): value is string => typeof value === 'string' && value.length > 0;
const isCellPosition = (value: unknown): value is { sheetId: string; row: number; column: number } => isRecord(value)
  && isNonEmptyString(value.sheetId) && Number.isSafeInteger(value.row) && Number.isSafeInteger(value.column)
  && Number(value.row) >= 0 && Number(value.column) >= 0;
const isCellNote = (value: unknown): value is CellNote => isRecord(value)
  && isNonEmptyString(value.id) && typeof value.author === 'string' && typeof value.text === 'string'
  && typeof value.createdAt === 'string' && typeof value.visible === 'boolean';
const isCommentReply = (value: unknown): value is CommentReply => isRecord(value)
  && isNonEmptyString(value.id) && typeof value.author === 'string'
  && typeof value.text === 'string' && typeof value.createdAt === 'string';
const isCommentThread = (value: unknown): value is CommentThread => isRecord(value)
  && isNonEmptyString(value.id) && isNonEmptyString(value.sheetId)
  && Number.isSafeInteger(value.row) && Number.isSafeInteger(value.column)
  && Number(value.row) >= 0 && Number(value.column) >= 0
  && typeof value.author === 'string' && typeof value.text === 'string'
  && typeof value.createdAt === 'string' && Array.isArray(value.replies)
  && value.replies.every(isCommentReply)
  && (value.resolved === undefined || typeof value.resolved === 'boolean')
  && (value.resolvedAt === undefined || typeof value.resolvedAt === 'string');
const isHyperlinkTarget = (value: unknown): boolean => {
  if (!isRecord(value) || typeof value.kind !== 'string') return false;
  if (value.kind === 'url') return typeof value.url === 'string' && value.url.length > 0;
  if (value.kind === 'email') return typeof value.address === 'string' && value.address.length > 0
    && (value.subject === undefined || typeof value.subject === 'string');
  if (value.kind === 'name') return isNonEmptyString(value.name);
  if (value.kind === 'sheet') return isNonEmptyString(value.sheetId)
    && (value.address === undefined || typeof value.address === 'string')
    && (value.row === undefined || (Number.isSafeInteger(value.row) && Number(value.row) >= 0))
    && (value.column === undefined || (Number.isSafeInteger(value.column) && Number(value.column) >= 0));
  return false;
};
const isCellHyperlink = (value: unknown): value is CellHyperlink => isRecord(value)
  && isNonEmptyString(value.id) && isHyperlinkTarget(value.target)
  && (value.tooltip === undefined || typeof value.tooltip === 'string');
const isNoteSet = (value: unknown): value is NoteSetParams => isRecord(value) && isCellPosition(value) && isCellNote((value as Record<string, unknown>).note);
const isNoteRemove = (value: unknown): value is NoteRemoveParams => isCellPosition(value);
const isNoteVisibility = (value: unknown): value is NoteVisibilityParams => isRecord(value) && isCellPosition(value) && typeof (value as Record<string, unknown>).visible === 'boolean';
const isCommentAdd = (value: unknown): value is CommentAddParams => {
  if (!isRecord(value) || !isCellPosition(value)) return false;
  const thread = (value as Record<string, unknown>).thread;
  return isCommentThread(thread)
    && thread.sheetId === value.sheetId
    && thread.row === value.row
    && thread.column === value.column;
};
const isCommentReplyParams = (value: unknown): value is CommentReplyParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.threadId) && isCommentReply(value.reply);
const isCommentReplyRemove = (value: unknown): value is CommentReplyRemoveParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.threadId) && isNonEmptyString(value.replyId);
const isCommentResolve = (value: unknown): value is CommentResolveParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.threadId)
  && typeof value.resolved === 'boolean'
  && (value.resolvedAt === undefined || typeof value.resolvedAt === 'string');
const isCommentRemove = (value: unknown): value is CommentRemoveParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.threadId);
const isCommentUpdate = (value: unknown): value is CommentUpdateParams => isRecord(value)
  && isNonEmptyString(value.sheetId) && isNonEmptyString(value.threadId)
  && Number.isSafeInteger(value.row) && Number(value.row) >= 0
  && Number.isSafeInteger(value.column) && Number(value.column) >= 0
  && typeof value.text === 'string';
const isHyperlinkSet = (value: unknown): value is HyperlinkSetParams => isRecord(value) && isCellPosition(value) && isCellHyperlink((value as Record<string, unknown>).hyperlink);
const isHyperlinkRemove = (value: unknown): value is HyperlinkRemoveParams => isCellPosition(value);

function reviewAffectedRanges(value: unknown): RangeRef[] {
  if (isCellPosition(value)) return cellRange(value.sheetId, value.row, value.column);
  if (isCommentAdd(value)) return cellRange(value.sheetId, value.thread.row, value.thread.column);
  if (isRecord(value) && isNonEmptyString(value.sheetId)) return sheetRange(value.sheetId);
  return [];
}

/** Apply a reviewed mutation through its literal canonical handler. */
function applyReviewMutation(mutationId: string, params: unknown, context: import('@react-sheets/command-runtime').CommandContext): void {
  switch (mutationId) {
    case 'note.set': {
      if (!isNoteSet(params)) throw new Error('Invalid note.set parameters');
      context.workbook.getSheet(params.sheetId).review.setNote(params.row, params.column, params.note);
      return;
    }
    case 'note.remove': {
      if (!isNoteRemove(params)) throw new Error('Invalid note.remove parameters');
      if (!context.workbook.getSheet(params.sheetId).review.removeNote(params.row, params.column)) throw new Error(`Note not found at ${params.sheetId}!${params.row}:${params.column}`);
      return;
    }
    case 'note.visibility': {
      if (!isNoteVisibility(params)) throw new Error('Invalid note.visibility parameters');
      context.workbook.getSheet(params.sheetId).review.updateNote(params.row, params.column, (note) => { note.visible = params.visible; });
      return;
    }
    case 'comment.add': {
      if (!isCommentAdd(params)) throw new Error('Invalid comment.add parameters');
      context.workbook.getSheet(params.sheetId).review.addThread(params.thread);
      return;
    }
    case 'comment.reply': {
      if (!isCommentReplyParams(params)) throw new Error('Invalid comment.reply parameters');
      context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
        if (thread.replies.some((entry) => entry.id === params.reply.id)) throw new Error(`Comment reply already exists: ${params.reply.id}`);
        thread.replies.push(structuredClone(params.reply));
      });
      return;
    }
    case 'comment.reply.remove': {
      if (!isCommentReplyRemove(params)) throw new Error('Invalid comment.reply.remove parameters');
      context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
        const index = thread.replies.findIndex((entry) => entry.id === params.replyId);
        if (index < 0) throw new Error(`Comment reply not found: ${params.replyId}`);
        thread.replies.splice(index, 1);
      });
      return;
    }
    case 'comment.resolve': {
      if (!isCommentResolve(params)) throw new Error('Invalid comment.resolve parameters');
      if (params.resolved && !params.resolvedAt) throw new Error('Resolving a comment requires resolvedAt from the operation payload');
      context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
        thread.resolved = params.resolved;
        thread.resolvedAt = params.resolved ? params.resolvedAt : undefined;
      });
      return;
    }
    case 'comment.remove': {
      if (!isCommentRemove(params)) throw new Error('Invalid comment.remove parameters');
      if (!context.workbook.getSheet(params.sheetId).review.removeThread(params.threadId)) throw new Error(`Comment thread not found: ${params.threadId}`);
      return;
    }
    case 'comment.update': {
      if (!isCommentUpdate(params)) throw new Error('Invalid comment.update parameters');
      context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
        if (thread.row !== params.row || thread.column !== params.column) throw new Error(`Comment thread ${params.threadId} moved before update`);
        thread.text = params.text;
      });
      return;
    }
    case 'hyperlink.set': {
      if (!isHyperlinkSet(params)) throw new Error('Invalid hyperlink.set parameters');
      validateHyperlinkTarget(params.hyperlink.target, context.workbook, params.sheetId);
      setCellHyperlink(context.workbook.getSheet(params.sheetId), params.row, params.column, params.hyperlink);
      return;
    }
    case 'hyperlink.remove': {
      if (!isHyperlinkRemove(params)) throw new Error('Invalid hyperlink.remove parameters');
      if (!removeCellHyperlink(context.workbook.getSheet(params.sheetId), params.row, params.column)) throw new Error(`Hyperlink not found at ${params.sheetId}!${params.row}:${params.column}`);
      return;
    }
    default:
      throw new Error(`Unknown review mutation: ${mutationId}`);
  }
}

export function registerReviewCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  runtime.registry.registerMutation<NoteSetParams>({
      id: 'note.set',
      handler: (item, context) => {
    const params = item.params;
    context.workbook.getSheet(params.sheetId).review.setNote(params.row, params.column, params.note);
  },
      metadata: {
    schema: { name: 'NoteSetParams', validate: isNoteSet },
    permission: { capability: 'review.note' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['note.set', 'note.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<NoteRemoveParams>({
      id: 'note.remove',
      handler: (item, context) => {
    const params = item.params;
    const removed = context.workbook.getSheet(params.sheetId).review.removeNote(params.row, params.column);
    if (!removed) throw new Error(`Note not found at ${params.sheetId}!${params.row}:${params.column}`);
  },
      metadata: {
    schema: { name: 'NoteRemoveParams', validate: isNoteRemove },
    permission: { capability: 'review.note' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['note.set'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<NoteVisibilityParams>({
      id: 'note.visibility',
      handler: (item, context) => {
    const params = item.params;
    context.workbook.getSheet(params.sheetId).review.updateNote(params.row, params.column, (note) => { note.visible = params.visible; });
  },
      metadata: {
    schema: { name: 'NoteVisibilityParams', validate: isNoteVisibility },
    permission: { capability: 'review.note' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['note.visibility'], minCount: 1, maxCount: 1 },
  },
    });

  runtime.registry.registerMutation<CommentAddParams>({
      id: 'comment.add',
      handler: (item, context) => {
    const params = item.params;
    context.workbook.getSheet(params.sheetId).review.addThread(params.thread);
  },
      metadata: {
    schema: { name: 'CommentAddParams', validate: isCommentAdd },
    permission: { capability: 'review.comment' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['comment.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<CommentReplyParams>({
      id: 'comment.reply',
      handler: (item, context) => {
    const params = item.params;
    context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
      if (thread.replies.some((entry) => entry.id === params.reply.id)) throw new Error(`Comment reply already exists: ${params.reply.id}`);
      thread.replies.push(structuredClone(params.reply));
    });
  },
      metadata: {
    schema: { name: 'CommentReplyParams', validate: isCommentReplyParams },
    permission: { capability: 'review.comment' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['comment.reply.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<CommentReplyRemoveParams>({
      id: 'comment.reply.remove',
      handler: (item, context) => {
    const params = item.params;
    context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
      const index = thread.replies.findIndex((entry) => entry.id === params.replyId);
      if (index < 0) throw new Error(`Comment reply not found: ${params.replyId}`);
      thread.replies.splice(index, 1);
    });
  },
      metadata: {
    schema: { name: 'CommentReplyRemoveParams', validate: isCommentReplyRemove },
    permission: { capability: 'review.comment' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['comment.reply'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<CommentResolveParams>({
      id: 'comment.resolve',
      handler: (item, context) => {
    const params = item.params;
    if (params.resolved && !params.resolvedAt) throw new Error('Resolving a comment requires resolvedAt from the operation payload');
    context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
      thread.resolved = params.resolved;
      thread.resolvedAt = params.resolved ? params.resolvedAt : undefined;
    });
  },
      metadata: {
    schema: { name: 'CommentResolveParams', validate: isCommentResolve },
    permission: { capability: 'review.comment' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['comment.resolve'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<CommentRemoveParams>({
      id: 'comment.remove',
      handler: (item, context) => {
    const params = item.params;
    const removed = context.workbook.getSheet(params.sheetId).review.removeThread(params.threadId);
    if (!removed) throw new Error(`Comment thread not found: ${params.threadId}`);
  },
      metadata: {
    schema: { name: 'CommentRemoveParams', validate: isCommentRemove },
    permission: { capability: 'review.comment' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'declared' },
    inversePolicy: { allowedMutationIds: ['comment.add'], minCount: 1, maxCount: 1 },
  },
    });

  runtime.registry.registerMutation<CommentUpdateParams>({
      id: 'comment.update',
      handler: (item, context) => {
    const params = item.params;
    context.workbook.getSheet(params.sheetId).review.updateThread(params.threadId, (thread) => {
      if (thread.row !== params.row || thread.column !== params.column) throw new Error(`Comment thread ${params.threadId} moved before update`);
      thread.text = params.text;
    });
  },
      metadata: {
    schema: { name: 'CommentUpdateParams', validate: isCommentUpdate },
    permission: { capability: 'review.comment' },
    affectedRanges: { resolve: (params) => isCommentUpdate(params) ? cellRange(params.sheetId, params.row, params.column) : [], mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['comment.update'], minCount: 1, maxCount: 1 },
  },
    });

  runtime.registry.registerMutation<HyperlinkSetParams>({
      id: 'hyperlink.set',
      handler: (item, context) => {
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    setCellHyperlink(sheet, params.row, params.column, params.hyperlink);
  },
      metadata: {
    schema: { name: 'HyperlinkSetParams', validate: isHyperlinkSet },
    permission: { capability: 'review.hyperlink' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['hyperlink.set', 'hyperlink.remove'], minCount: 1, maxCount: 1 },
  },
    });
  runtime.registry.registerMutation<HyperlinkRemoveParams>({
      id: 'hyperlink.remove',
      handler: (item, context) => {
    const params = item.params;
    const sheet = context.workbook.getSheet(params.sheetId);
    if (!removeCellHyperlink(sheet, params.row, params.column)) {
      throw new Error(`Hyperlink not found at ${params.sheetId}!${params.row}:${params.column}`);
    }
  },
      metadata: {
    schema: { name: 'HyperlinkRemoveParams', validate: isHyperlinkRemove },
    permission: { capability: 'review.hyperlink' },
    affectedRanges: { resolve: reviewAffectedRanges, mode: 'exact' },
    inversePolicy: { allowedMutationIds: ['hyperlink.set'], minCount: 1, maxCount: 1 },
  },
    });

  runtime.registry.registerCommand<NoteSetParams>({
    id: 'note.set',
    execute: (params, context) => {
      const note = context.workbook.getSheet(params.sheetId).review.getNoteAt(params.row, params.column);
      const affectedRanges = cellRange(params.sheetId, params.row, params.column);
      if (note) {
        context.applyMutation({ id: 'note.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'note.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, note: structuredClone(note) }, affectedRanges }], apply: () => applyReviewMutation('note.set', params, context) });
      } else {
        context.applyMutation({ id: 'note.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'note.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column }, affectedRanges }], apply: () => applyReviewMutation('note.set', params, context) });
      }
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('note.set');

  runtime.registry.registerCommand<NoteRemoveParams>({
    id: 'note.remove',
    execute: (params, context) => {
      const note = context.workbook.getSheet(params.sheetId).review.getNoteAt(params.row, params.column);
      if (!note) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, params.row, params.column);
      context.applyMutation({ id: 'note.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'note.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, note: structuredClone(note) }, affectedRanges }], apply: () => applyReviewMutation('note.remove', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('note.remove');

  runtime.registry.registerCommand<NoteVisibilityParams>({
    id: 'note.visibility',
    execute: (params, context) => {
      const note = context.workbook.getSheet(params.sheetId).review.getNoteAt(params.row, params.column);
      if (!note) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, params.row, params.column);
      context.applyMutation({ id: 'note.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'note.visibility', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, visible: note.visible }, affectedRanges }], apply: () => applyReviewMutation('note.visibility', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('note.visibility');

  runtime.registry.registerCommand<CommentAddParams>({
    id: 'comment.add',
    execute: (params, context) => {
      const affectedRanges = cellRange(params.sheetId, params.row, params.column);
      context.applyMutation({ id: 'comment.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'comment.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, threadId: params.thread.id }, affectedRanges }], apply: () => applyReviewMutation('comment.add', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('comment.add');

  runtime.registry.registerCommand<CommentReplyParams>({
    id: 'comment.reply',
    execute: (params, context) => {
      const thread = context.workbook.getSheet(params.sheetId).review.getThread(params.threadId);
      if (!thread) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, thread.row, thread.column);
      context.applyMutation({ id: 'comment.reply', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'comment.reply.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, threadId: params.threadId, replyId: params.reply.id }, affectedRanges }], apply: () => applyReviewMutation('comment.reply', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('comment.reply');

  runtime.registry.registerCommand<CommentResolveParams>({
    id: 'comment.resolve',
    execute: (params, context) => {
      const thread = context.workbook.getSheet(params.sheetId).review.getThread(params.threadId);
      if (!thread) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, thread.row, thread.column);
      const previous: CommentResolveParams = { sheetId: params.sheetId, threadId: params.threadId, resolved: thread.resolved ?? false, ...(thread.resolvedAt ? { resolvedAt: thread.resolvedAt } : {}) };
      context.applyMutation({ id: 'comment.resolve', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'comment.resolve', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }], apply: () => applyReviewMutation('comment.resolve', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('comment.resolve');

  runtime.registry.registerCommand<CommentRemoveParams>({
    id: 'comment.remove',
    execute: (params, context) => {
      const thread = context.workbook.getSheet(params.sheetId).review.getThread(params.threadId);
      if (!thread) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, thread.row, thread.column);
      context.applyMutation({ id: 'comment.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'comment.add', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: thread.row, column: thread.column, thread: structuredClone(thread) }, affectedRanges }], apply: () => applyReviewMutation('comment.remove', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('comment.remove');

  runtime.registry.registerCommand<CommentUpdateParams>({
    id: 'comment.update',
    execute: (params, context) => {
      if (!isCommentUpdate(params)) throw new Error('Invalid comment.update parameters');
      const thread = context.workbook.getSheet(params.sheetId).review.getThread(params.threadId);
      if (!thread) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, thread.row, thread.column);
      if (thread.row !== params.row || thread.column !== params.column) throw new Error(`Comment thread ${params.threadId} moved before update`);
      const previous: CommentUpdateParams = { sheetId: params.sheetId, threadId: params.threadId, row: thread.row, column: thread.column, text: thread.text };
      context.applyMutation({ id: 'comment.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'comment.update', unitId: context.workbook.unitId, sheetId: params.sheetId, params: previous, affectedRanges }], apply: () => applyReviewMutation('comment.update', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('comment.update');

  runtime.registry.registerCommand<CommentReplyRemoveParams>({
    id: 'comment.reply.remove',
    execute: (params, context) => {
      const thread = context.workbook.getSheet(params.sheetId).review.getThread(params.threadId);
      const reply = thread?.replies.find((entry) => entry.id === params.replyId);
      if (!thread || !reply) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, thread.row, thread.column);
      context.applyMutation({ id: 'comment.reply.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'comment.reply', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, threadId: params.threadId, reply: structuredClone(reply) }, affectedRanges }], apply: () => applyReviewMutation('comment.reply.remove', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('comment.reply.remove');

  runtime.registry.registerCommand<HyperlinkSetParams>({
    id: 'hyperlink.set',
    execute: (params, context) => {
      if (!isHyperlinkSet(params)) throw new Error('Invalid hyperlink.set parameters');
      validateHyperlinkTarget(params.hyperlink.target, context.workbook, params.sheetId);
      const hyperlink = getCellHyperlink(context.workbook.getSheet(params.sheetId), params.row, params.column);
      const affectedRanges = cellRange(params.sheetId, params.row, params.column);
      if (hyperlink) {
        context.applyMutation({ id: 'hyperlink.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'hyperlink.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, hyperlink }, affectedRanges }], apply: () => applyReviewMutation('hyperlink.set', params, context) });
      } else {
        context.applyMutation({ id: 'hyperlink.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'hyperlink.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column }, affectedRanges }], apply: () => applyReviewMutation('hyperlink.set', params, context) });
      }
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('hyperlink.set');

  runtime.registry.registerCommand<HyperlinkRemoveParams>({
    id: 'hyperlink.remove',
    execute: (params, context) => {
      const hyperlink = getCellHyperlink(context.workbook.getSheet(params.sheetId), params.row, params.column);
      if (!hyperlink) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
      const affectedRanges = cellRange(params.sheetId, params.row, params.column);
      context.applyMutation({ id: 'hyperlink.remove', unitId: context.workbook.unitId, sheetId: params.sheetId, params, affectedRanges, inverse: [{ id: 'hyperlink.set', unitId: context.workbook.unitId, sheetId: params.sheetId, params: { sheetId: params.sheetId, row: params.row, column: params.column, hyperlink }, affectedRanges }], apply: () => applyReviewMutation('hyperlink.remove', params, context) });
      return { operationId: context.operationId, mutationCount: 1, affectedRanges };
    },
  });
  commandIds.push('hyperlink.remove');

  return commandIds;
}

export { serializeHyperlink } from './helpers';

export const REVIEW_RIBBON_ENTRIES = [
  { id: 'review-note', tab: 'Review', group: 'Comments', label: 'New Note', commandId: 'note.set', icon: 'comment' },
  { id: 'review-comment', tab: 'Review', group: 'Comments', label: 'New Comment', commandId: 'comment.add', icon: 'comment' },
  { id: 'review-hyperlink', tab: 'Review', group: 'Links', label: 'Insert Link', commandId: 'hyperlink.set', icon: 'link' },
  { id: 'review-resolve', tab: 'Review', group: 'Comments', label: 'Resolve Thread', commandId: 'comment.resolve', icon: 'comment' },
] as const;

export function registerReviewFeature(runtime: CommandRuntime): SpreadsheetFeatureManifest {
  const commandIds = registerReviewCommands(runtime);
  return {
    id: 'review',
    version: '1.0.0',
    commandIds,
    mutationIds: ['note.set', 'note.remove', 'note.visibility', 'comment.add', 'comment.reply', 'comment.reply.remove', 'comment.resolve', 'comment.remove', 'comment.update', 'hyperlink.set', 'hyperlink.remove'],
    ribbon: [...REVIEW_RIBBON_ENTRIES],
    permissions: ['review.comment', 'review.note', 'review.hyperlink'],
  };
}
