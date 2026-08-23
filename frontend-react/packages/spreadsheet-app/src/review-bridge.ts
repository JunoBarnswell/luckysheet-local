import type {
  CellComment,
  CellHyperlink,
  CellNote,
  CommentReply,
  CommentThread,
  WorksheetModel,
} from '@react-sheets/core-model';
import { getCellNote } from '@react-sheets/core-model';
import { serializeHyperlink } from './features/review/commands';

export function findCommentThreadAt(sheet: WorksheetModel, row: number, column: number): CommentThread | undefined {
  return sheet.commentThreads.find((thread) => thread.row === row && thread.column === column);
}

export function threadToCellComment(thread: CommentThread): CellComment {
  return {
    id: thread.id,
    author: thread.author,
    text: thread.text,
    createdAt: thread.createdAt,
    mentions: thread.mentions,
    replies: thread.replies,
    resolved: thread.resolved,
    resolvedAt: thread.resolvedAt,
  };
}

export function buildCommentThread(
  sheetId: string,
  row: number,
  column: number,
  author: string,
  text: string,
  threadId: string,
): CommentThread {
  const mentions = [...text.matchAll(/@([\w-]+)/g)].map((match) => match[1]).filter(Boolean) as string[];
  return {
    id: threadId,
    sheetId,
    row,
    column,
    author,
    text: text.trim(),
    createdAt: new Date().toISOString(),
    mentions,
    replies: [],
    resolved: false,
  };
}

export function buildCommentReply(author: string, text: string, replyId: string): CommentReply {
  return {
    id: replyId,
    author,
    text: text.trim(),
    createdAt: new Date().toISOString(),
  };
}

export function buildCellNote(author: string, text: string, noteId: string): CellNote {
  return {
    id: noteId,
    author,
    text: text.trim(),
    createdAt: new Date().toISOString(),
    visible: true,
  };
}

export function parseUrlHyperlink(url: string, hyperlinkId: string): CellHyperlink {
  const trimmed = url.trim();
  if (trimmed.startsWith('mailto:')) {
    const body = trimmed.slice(7);
    const [address, query = ''] = body.split('?');
    const subjectMatch = query.match(/(?:^|&)subject=([^&]+)/);
    return {
      id: hyperlinkId,
      target: {
        kind: 'email',
        address: address ?? '',
        subject: subjectMatch?.[1] ? decodeURIComponent(subjectMatch[1]) : undefined,
      },
    };
  }
  if (trimmed.startsWith('#sheet:')) {
    const rest = trimmed.slice(7);
    const [sheetId, anchor] = rest.split('!');
    if (anchor?.includes(':')) {
      const [rowText, columnText] = anchor.split(':');
      return {
        id: hyperlinkId,
        target: {
          kind: 'sheet',
          sheetId: sheetId ?? '',
          row: Number(rowText),
          column: Number(columnText),
        },
      };
    }
    return {
      id: hyperlinkId,
      target: {
        kind: 'sheet',
        sheetId: sheetId ?? '',
        address: anchor,
      },
    };
  }
  if (trimmed.startsWith('#name:')) {
    return {
      id: hyperlinkId,
      target: { kind: 'name', name: trimmed.slice(6) },
    };
  }
  return {
    id: hyperlinkId,
    target: { kind: 'url', url: trimmed },
  };
}

export function resolveHyperlinkDisplay(
  hyperlink?: string,
  hyperlinkDetail?: CellHyperlink,
): string | undefined {
  if (hyperlink) return hyperlink;
  if (hyperlinkDetail) return serializeHyperlink(hyperlinkDetail);
  return undefined;
}

export function hasReviewMarkerAt(sheet: WorksheetModel, row: number, column: number): boolean {
  return Boolean(
    findCommentThreadAt(sheet, row, column)
    || getCellNote(sheet, row, column)
    || sheet.cells.get(row, column)?.comment,
  );
}
