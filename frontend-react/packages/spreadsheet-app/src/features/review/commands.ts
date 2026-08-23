import type { CommandRuntime } from '@react-sheets/command-runtime';
import type { CellHyperlink, CellNote, CommentReply, CommentThread } from '@react-sheets/core-model';
import { noteCellKey } from '@react-sheets/core-model';
import type { SpreadsheetFeatureManifest } from '../../feature-registry';
import { applyTrackedMutation, registerMutationHandler, removeById, sheetRange } from '../../command-helpers';

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

function cellRange(sheetId: string, row: number, column: number) {
  return [{ sheetId, startRow: row, endRow: row, startColumn: column, endColumn: column }];
}

export function registerReviewCommands(runtime: CommandRuntime): string[] {
  const commandIds: string[] = [];

  registerMutationHandler<NoteSetParams>(runtime, 'note.set', (params, context) => {
    context.workbook.getSheet(params.sheetId).notes.set(noteCellKey(params.row, params.column), structuredClone(params.note));
  });
  registerMutationHandler<NoteRemoveParams>(runtime, 'note.remove', (params, context) => {
    context.workbook.getSheet(params.sheetId).notes.delete(noteCellKey(params.row, params.column));
  });
  registerMutationHandler<NoteVisibilityParams>(runtime, 'note.visibility', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    if (note) note.visible = params.visible;
  });

  registerMutationHandler<CommentAddParams>(runtime, 'comment.add', (params, context) => {
    context.workbook.getSheet(params.sheetId).commentThreads.push(structuredClone(params.thread));
  });
  registerMutationHandler<CommentReplyParams>(runtime, 'comment.reply', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    if (thread) thread.replies.push(structuredClone(params.reply));
  });
  registerMutationHandler<CommentResolveParams>(runtime, 'comment.resolve', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    if (thread) {
      thread.resolved = params.resolved;
      thread.resolvedAt = params.resolved ? new Date().toISOString() : undefined;
    }
  });
  registerMutationHandler<CommentRemoveParams>(runtime, 'comment.remove', (params, context) => {
    removeById(context.workbook.getSheet(params.sheetId).commentThreads, params.threadId);
  });

  registerMutationHandler<HyperlinkSetParams>(runtime, 'hyperlink.set', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const current = sheet.cells.get(params.row, params.column) ?? { value: null };
    sheet.cells.set(params.row, params.column, {
      ...current,
      hyperlinkDetail: structuredClone(params.hyperlink),
      hyperlink: serializeHyperlink(params.hyperlink),
    });
  });
  registerMutationHandler<HyperlinkRemoveParams>(runtime, 'hyperlink.remove', (params, context) => {
    const sheet = context.workbook.getSheet(params.sheetId);
    const current = sheet.cells.get(params.row, params.column);
    if (!current) return;
    const next = { ...current };
    delete next.hyperlink;
    delete next.hyperlinkDetail;
    if (next.value == null && !next.formula && !next.style) sheet.cells.delete(params.row, params.column);
    else sheet.cells.set(params.row, params.column, next);
  });

  const registerSimple = <P>(
    commandId: string,
    mutationId: string,
    readPrevious: (params: P, context: import('@react-sheets/command-runtime').CommandContext) => unknown,
  ): void => {
    runtime.registry.registerCommand<P>({
      id: commandId,
      execute: (params, context) => {
        const previous = readPrevious(params, context);
        const affectedRanges = 'row' in (params as object) && 'column' in (params as object)
          ? cellRange((params as NoteSetParams).sheetId, (params as NoteSetParams).row, (params as NoteSetParams).column)
          : sheetRange((params as CommentRemoveParams).sheetId);
        applyTrackedMutation(context, {
          id: mutationId,
          sheetId: (params as { sheetId: string }).sheetId,
          params,
          inverseParams: previous ?? params,
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
    return note ? { sheetId: params.sheetId, row: params.row, column: params.column, note: structuredClone(note) } : { sheetId: params.sheetId, row: params.row, column: params.column, note: params.note };
  });
  registerSimple<NoteRemoveParams>('note.remove', 'note.remove', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    return note ? { sheetId: params.sheetId, row: params.row, column: params.column, note: structuredClone(note) } : params;
  });
  registerSimple<NoteVisibilityParams>('note.visibility', 'note.visibility', (params, context) => {
    const note = context.workbook.getSheet(params.sheetId).notes.get(noteCellKey(params.row, params.column));
    return { sheetId: params.sheetId, row: params.row, column: params.column, visible: note?.visible ?? !params.visible };
  });
  registerSimple<CommentAddParams>('comment.add', 'comment.add', () => undefined);
  registerSimple<CommentReplyParams>('comment.reply', 'comment.reply', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    return thread ? { sheetId: params.sheetId, threadId: params.threadId, reply: structuredClone(thread.replies.at(-1) ?? params.reply) } : params;
  });
  registerSimple<CommentResolveParams>('comment.resolve', 'comment.resolve', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    return { sheetId: params.sheetId, threadId: params.threadId, resolved: !(thread?.resolved ?? false) };
  });
  registerSimple<CommentRemoveParams>('comment.remove', 'comment.remove', (params, context) => {
    const thread = context.workbook.getSheet(params.sheetId).commentThreads.find((entry) => entry.id === params.threadId);
    return thread ? { sheetId: params.sheetId, row: thread.row, column: thread.column, thread: structuredClone(thread) } : params;
  });
  registerSimple<HyperlinkSetParams>('hyperlink.set', 'hyperlink.set', (params, context) => {
    const cell = context.workbook.getSheet(params.sheetId).cells.get(params.row, params.column);
    return {
      sheetId: params.sheetId,
      row: params.row,
      column: params.column,
      hyperlink: cell?.hyperlinkDetail ? structuredClone(cell.hyperlinkDetail) : params.hyperlink,
    };
  });
  registerSimple<HyperlinkRemoveParams>('hyperlink.remove', 'hyperlink.remove', (params, context) => {
    const cell = context.workbook.getSheet(params.sheetId).cells.get(params.row, params.column);
    return cell?.hyperlinkDetail ? { sheetId: params.sheetId, row: params.row, column: params.column, hyperlink: structuredClone(cell.hyperlinkDetail) } : params;
  });

  return commandIds;
}

export function serializeHyperlink(link: CellHyperlink): string {
  switch (link.target.kind) {
    case 'url':
      return link.target.url;
    case 'email':
      return link.target.subject
        ? `mailto:${link.target.address}?subject=${encodeURIComponent(link.target.subject)}`
        : `mailto:${link.target.address}`;
    case 'sheet':
      if (link.target.row != null && link.target.column != null) {
        return `#sheet:${link.target.sheetId}!${link.target.row}:${link.target.column}`;
      }
      return link.target.address ? `#sheet:${link.target.sheetId}!${link.target.address}` : `#sheet:${link.target.sheetId}`;
    case 'name':
      return `#name:${link.target.name}`;
  }
}

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
    mutationIds: ['note.set', 'note.remove', 'note.visibility', 'comment.add', 'comment.reply', 'comment.resolve', 'comment.remove', 'hyperlink.set', 'hyperlink.remove'],
    ribbon: [...REVIEW_RIBBON_ENTRIES],
    permissions: ['review.comment', 'review.note', 'review.hyperlink'],
  };
}
