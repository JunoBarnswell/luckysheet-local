import type {
  CellComment,
  CellHyperlink,
  CellNote,
  CommentReply,
  CommentThread,
  HyperlinkTarget,
  WorkbookModel,
  WorksheetModel,
} from '@react-sheets/core-model';
import { getCellNote, parseAddress } from '@react-sheets/core-model';

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
  return sheet.review.getThreadsAt(row, column)[0];
}

export function findCommentThreadsAt(sheet: WorksheetModel, row: number, column: number): CommentThread[] {
  return sheet.review.getThreadsAt(row, column);
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

/** Validate the canonical target before it enters a command or mutation. */
export function validateHyperlinkTarget(target: HyperlinkTarget, workbook: WorkbookModel, sourceSheetId: string): void {
  if (target.kind === 'url') {
    const url = target.url.trim();
    if (!url) throw new Error('Hyperlink URL is required');
    let parsed: URL;
    try { parsed = new URL(url); } catch { throw new Error(`Invalid hyperlink URL: ${target.url}`); }
    if (!['http:', 'https:', 'ftp:'].includes(parsed.protocol)) throw new Error(`Unsupported hyperlink URL scheme: ${parsed.protocol}`);
    return;
  }
  if (target.kind === 'email') {
    const address = target.address.trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(address)) throw new Error(`Invalid email hyperlink address: ${target.address}`);
    if (target.subject !== undefined && typeof target.subject !== 'string') throw new Error('Email hyperlink subject is invalid');
    return;
  }
  if (target.kind === 'name') {
    const name = target.name.trim();
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name)) throw new Error(`Invalid defined-name hyperlink: ${target.name}`);
    const found = workbook.definedNameModels.some((entry) => entry.name.toLowerCase() === name.toLowerCase()
      && (entry.scope === 'workbook' || entry.sheetId === sourceSheetId));
    if (!found) throw new Error(`Defined name not found: ${name}`);
    return;
  }
  if (!workbook.getSheets().some((sheet) => sheet.id === target.sheetId)) throw new Error(`Hyperlink target sheet not found: ${target.sheetId}`);
  const hasAddress = target.address !== undefined;
  const hasCoordinates = target.row !== undefined || target.column !== undefined;
  if (hasAddress && hasCoordinates) throw new Error('Worksheet hyperlink must use either address or row/column coordinates');
  if (!hasAddress && !hasCoordinates) throw new Error('Worksheet hyperlink address is required');
  const targetSheet = workbook.getSheet(target.sheetId);
  const coordinate = hasAddress ? parseAddress(target.address ?? '') : { row: target.row, column: target.column };
  const row = coordinate?.row;
  const column = coordinate?.column;
  if (row === undefined || column === undefined || !Number.isSafeInteger(row) || !Number.isSafeInteger(column)
    || row < 0 || column < 0
    || row >= targetSheet.rowCount || column >= targetSheet.columnCount) {
    throw new Error(`Worksheet hyperlink address is outside ${targetSheet.name}`);
  }
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
