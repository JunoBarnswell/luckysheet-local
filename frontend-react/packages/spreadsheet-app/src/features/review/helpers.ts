import type {
  CellComment,
  CellHyperlink,
  CellNote,
  CommentReply,
  CommentThread,
  WorksheetModel,
} from '@react-sheets/core-model';
import { getCellNote } from '@react-sheets/core-model';

function hyperlinkKey(row: number, column: number): string {
  return `${row}:${column}`;
}

export function getCellHyperlink(sheet: WorksheetModel, row: number, column: number): CellHyperlink | undefined {
  const link = sheet.hyperlinks.get(hyperlinkKey(row, column));
  return link ? structuredClone(link) : undefined;
}

export function setCellHyperlink(
  sheet: WorksheetModel,
  row: number,
  column: number,
  link: CellHyperlink,
): void {
  sheet.hyperlinks.set(hyperlinkKey(row, column), structuredClone(link));
}

export function removeCellHyperlink(sheet: WorksheetModel, row: number, column: number): CellHyperlink | undefined {
  const links = sheet.hyperlinks;
  const key = hyperlinkKey(row, column);
  const previous = links.get(key);
  links.delete(key);
  return previous ? structuredClone(previous) : undefined;
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

export function findCommentThreadAt(sheet: WorksheetModel, row: number, column: number): CommentThread | undefined {
  return sheet.commentThreads.find((thread) => thread.row === row && thread.column === column);
}

export function findCommentThreadsAt(sheet: WorksheetModel, row: number, column: number): CommentThread[] {
  return sheet.commentThreads
    .filter((thread) => thread.row === row && thread.column === column)
    .map((thread) => structuredClone(thread));
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
  hyperlink?: CellHyperlink,
): string | undefined {
  return hyperlink ? serializeHyperlink(hyperlink) : undefined;
}

export function hasReviewMarkerAt(sheet: WorksheetModel, row: number, column: number): boolean {
  return Boolean(
    findCommentThreadsAt(sheet, row, column).length > 0
    || getCellNote(sheet, row, column)
    || getCellHyperlink(sheet, row, column),
  );
}
