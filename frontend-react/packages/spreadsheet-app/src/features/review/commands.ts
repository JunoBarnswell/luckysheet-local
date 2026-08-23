import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { CellHyperlink, CellNote, CommentReply, CommentThread } from '@react-sheets/core-model';
import { noteCellKey } from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { applyTrackedMutation, registerMutationHandler, removeById, sheetRange } from '../../command-helpers';
import {
  getCellHyperlink,
  removeCellHyperlink,
  setCellHyperlink,
} from './helpers';

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

interface MutationInverse {
  id: string;
  params: unknown;
  affectedRanges?: ReturnType<typeof cellRange>;
}

function cellRange(sheetId: string, row: number, column: number) {
  return [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
}

export function registerReviewCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<NoteSetParams>(runtime, 'note.set', (params, context) => {
    context.workbook.getSheet(params.sheetId).notes.set(noteCellKey(params.row, params.column), structuredClone(params.note));
  });
  registerMutationHandler<NoteRemoveParams>(runtime, 'note.remove', (params, context) => {
    const removed = context.workbook.getSheet(params.sheetId).notes.delete(noteCellKey(params.row, params.column));
    if (!removed) throw new Error(`Note not found at ${params.sheetId}!${params.row}:${params.column}`);
  });
  registerMutationHandler<NoteVisibilityParams>(runtime, 'note.visibility', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    if (!note) throw new Error(`Note not found at ${params.sheetId}!${params.row}:${params.column}`);
    note.visible = params.visible;
  });

  registerMutationHandler<CommentAddParams>(runtime, 'comment.add', (params, context) => {
    const threads = context.workbook.getSheet(params.sheetId).commentThreads;
    if (threads.some((entry) => entry.id === params.thread.id)) throw new Error(`Comment thread already exists: ${params.thread.id}`);
    threads.push(structuredClone(params.thread));
  });
  registerMutationHandler<CommentReplyParams>(runtime, 'comment.reply', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    if (!thread) throw new Error(`Comment thread not found: ${params.threadId}`);
    if (thread.replies.some((entry) => entry.id === params.reply.id)) throw new Error(`Comment reply already exists: ${params.reply.id}`);
    thread.replies.push(structuredClone(params.reply));
  });
  registerMutationHandler<CommentReplyRemoveParams>(runtime, 'comment.reply.remove', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    if (!thread) throw new Error(`Comment thread not found: ${params.threadId}`);
    const index = thread.replies.findIndex((entry) => entry.id === params.replyId);
    if (index < 0) throw new Error(`Comment reply not found: ${params.replyId}`);
    thread.replies.splice(index, 1);
  });
  registerMutationHandler<CommentResolveParams>(runtime, 'comment.resolve', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    if (!thread) throw new Error(`Comment thread not found: ${params.threadId}`);
    if (params.resolved && !params.resolvedAt) throw new Error('Resolving a comment requires resolvedAt from the operation payload');
    thread.resolved = params.resolved;
    thread.resolvedAt = params.resolved ? params.resolvedAt : undefined;
  });
  registerMutationHandler<CommentRemoveParams>(runtime, 'comment.remove', (params, context) => {
    const removed = removeById(context.workbook.getSheet(params.sheetId).commentThreads, params.threadId);
    if (!removed) throw new Error(`Comment thread not found: ${params.threadId}`);
  });

  registerMutationHandler<HyperlinkSetParams>(runtime, 'hyperlink.set', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    setCellHyperlink(sheet, params.row, params.column, params.hyperlink);
  });
  registerMutationHandler<HyperlinkRemoveParams>(runtime, 'hyperlink.remove', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    if (!removeCellHyperlink(sheet, params.row, params.column)) {
      throw new Error(`Hyperlink not found at ${params.sheetId}!${params.row}:${params.column}`);
    }
  });

  const registerSimple = <P>(
    commandId: string,
    mutationId: string,
    readInverse: (params: P, context: import('@react-sheets/command-runtime').CommandContext) => MutationInverse | undefined,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const inverse = readInverse(params, context);
        if (!inverse) return { operationId: context.operationId, mutationCount: 0, affectedRanges: [] };
        const affectedRanges = 'row' in (params as object) && 'column' in (params as object)
          ? cellRange((params as NoteSetParams).sheetId, (params as NoteSetParams).row, (params as NoteSetParams).column)
          : inverse.affectedRanges ?? sheetRange((params as CommentRemoveParams).sheetId);
        applyTrackedMutation(context, {
          id: mutationId,
          sheetId: (params as { sheetId: string }).sheetId,
          params,
          inverseId: inverse.id,
          inverseParams: inverse.params,
          affectedRanges,
          apply: () => runtime.registry.getMutation(mutationId)({ id: mutationId, unitId: context.workbook.unitId, sheetId: (params as { sheetId: string }).sheetId, params, affectedRanges }, context),
        });
        return { operationId: context.operationId, mutationCount: 1, affectedRanges };
      },
    });
    commandIds.push(commandId);
  };

  registerSimple<NoteSetParams>('note.set', 'note.set', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    return note
      ? { id: 'note.set', params: { sheetId: params.sheetId, row: params.row, column: params.column, note: structuredClone(note) } }
      : { id: 'note.remove', params: { sheetId: params.sheetId, row: params.row, column: params.column } };
  });
  registerSimple<NoteRemoveParams>('note.remove', 'note.remove', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    return note
      ? { id: 'note.set', params: { sheetId: params.sheetId, row: params.row, column: params.column, note: structuredClone(note) } }
      : undefined;
  });
  registerSimple<NoteVisibilityParams>('note.visibility', 'note.visibility', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    return note
      ? { id: 'note.visibility', params: { sheetId: params.sheetId, row: params.row, column: params.column, visible: note.visible } }
      : undefined;
  });
  registerSimple<CommentAddParams>('comment.add', 'comment.add', (params) => ({
    id: 'comment.remove',
    params: { sheetId: params.sheetId, threadId: params.thread.id },
  }));
  registerSimple<CommentReplyParams>('comment.reply', 'comment.reply', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    return thread
      ? {
          id: 'comment.reply.remove',
          params: { sheetId: params.sheetId, threadId: params.threadId, replyId: params.reply.id },
          affectedRanges: cellRange(params.sheetId, thread.row, thread.column),
        }
      : undefined;
  });
  registerSimple<CommentResolveParams>('comment.resolve', 'comment.resolve', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    return thread
      ? {
          id: 'comment.resolve',
          params: {
            sheetId: params.sheetId,
            threadId: params.threadId,
            resolved: thread.resolved ?? false,
            ...(thread.resolvedAt ? { resolvedAt: thread.resolvedAt } : {}),
          },
          affectedRanges: cellRange(params.sheetId, thread.row, thread.column),
        }
      : undefined;
  });
  registerSimple<CommentRemoveParams>('comment.remove', 'comment.remove', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    return thread
      ? {
          id: 'comment.add',
          params: { sheetId: params.sheetId, row: thread.row, column: thread.column, thread: structuredClone(thread) },
          affectedRanges: cellRange(params.sheetId, thread.row, thread.column),
        }
      : undefined;
  });
  registerSimple<CommentReplyRemoveParams>('comment.reply.remove', 'comment.reply.remove', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    const reply = thread?.replies.find((entry) => entry.id === params.replyId);
    return thread && reply
      ? {
          id: 'comment.reply',
          params: { sheetId: params.sheetId, threadId: params.threadId, reply: structuredClone(reply) },
          affectedRanges: cellRange(params.sheetId, thread.row, thread.column),
        }
      : undefined;
  });
  registerSimple<HyperlinkSetParams>('hyperlink.set', 'hyperlink.set', (params, context) => {
    const hyperlink = getCellHyperlink(context.workbook.getSheet(params.sheetId), params.row, params.column);
    return hyperlink
      ? { id: 'hyperlink.set', params: { sheetId: params.sheetId, row: params.row, column: params.column, hyperlink } }
      : { id: 'hyperlink.remove', params: { sheetId: params.sheetId, row: params.row, column: params.column } };
  });
  registerSimple<HyperlinkRemoveParams>('hyperlink.remove', 'hyperlink.remove', (params, context) => {
    const hyperlink = getCellHyperlink(context.workbook.getSheet(params.sheetId), params.row, params.column);
    return hyperlink
      ? { id: 'hyperlink.set', params: { sheetId: params.sheetId, row: params.row, column: params.column, hyperlink } }
      : undefined;
  });

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
    mutationIds: ['note.set', 'note.remove', 'note.visibility', 'comment.add', 'comment.reply', 'comment.reply.remove', 'comment.resolve', 'comment.remove', 'hyperlink.set', 'hyperlink.remove'],
    ribbon: [...REVIEW_RIBBON_ENTRIES],
    permissions: ['review.comment', 'review.note', 'review.hyperlink'],
  };
}
