import type { CellNote, CommentReply, CommentThread } from './domain';

export type ReviewCellKey = string;

export interface ReviewStoreSnapshot {
  notesByCell: Record<ReviewCellKey, string>;
  notesById: Record<string, CellNote>;
  threadIdsByCell: Record<ReviewCellKey, string[]>;
  threadsById: Record<string, CommentThread>;
}

export interface ReviewNoteEntry {
  key: ReviewCellKey;
  row: number;
  column: number;
  note: CellNote;
}

export interface ReviewCellMetadataEntry {
  row: number;
  column: number;
  note?: CellNote;
  threads: CommentThread[];
}

export interface ReviewTextMetadataEntry {
  row: number;
  column: number;
  note?: Pick<CellNote, 'id' | 'text'>;
  threads: Array<Pick<CommentThread, 'id' | 'text'>>;
}

function cellKey(row: number, column: number): ReviewCellKey {
  if (!Number.isSafeInteger(row) || row < 0 || !Number.isSafeInteger(column) || column < 0) {
    throw new Error(`Review coordinate is invalid: ${row}:${column}`);
  }
  return `${row}:${column}`;
}

function parseCellKey(key: ReviewCellKey): { row: number; column: number } {
  const parts = key.split(':');
  const row = Number(parts[0]);
  const column = Number(parts[1]);
  if (parts.length !== 2 || !Number.isSafeInteger(row) || row < 0 || !Number.isSafeInteger(column) || column < 0) {
    throw new Error(`Review cell key is invalid: ${key}`);
  }
  return { row, column };
}

function cloneReply(reply: CommentReply): CommentReply {
  return structuredClone(reply);
}

/**
 * The sole runtime owner of worksheet notes and threaded comments. The maps
 * are deliberately indexed in both directions so a cell lookup never scans
 * the complete review collection and a mutation cannot leave a dangling id.
 */
export class ReviewStore {
  private readonly notesByCell = new Map<ReviewCellKey, string>();
  private readonly notesById = new Map<string, CellNote>();
  private readonly noteCellById = new Map<string, ReviewCellKey>();
  private readonly threadIdsByCell = new Map<ReviewCellKey, string[]>();
  private readonly threadsById = new Map<string, CommentThread>();
  private readonly reviewCellsByRow = new Map<number, Set<ReviewCellKey>>();
  private sortedReviewRows?: number[];
  private readonly sortedReviewCellsByRow = new Map<number, Array<{ column: number; key: ReviewCellKey }>>();

  constructor(readonly sheetId: string) {
    if (!sheetId.trim()) throw new Error('ReviewStore requires a sheet id');
  }

  get noteCount(): number { return this.notesById.size; }
  get threadCount(): number { return this.threadsById.size; }

  hasNoteAt(row: number, column: number): boolean {
    return this.notesByCell.has(cellKey(row, column));
  }

  getNoteAt(row: number, column: number): CellNote | undefined {
    const id = this.notesByCell.get(cellKey(row, column));
    if (id === undefined) return undefined;
    const note = this.notesById.get(id);
    if (!note) throw new Error(`Review note index is dangling: ${id}`);
    return structuredClone(note);
  }

  getNoteById(id: string): CellNote | undefined {
    const note = this.notesById.get(id);
    return note ? structuredClone(note) : undefined;
  }

  setNote(row: number, column: number, note: CellNote): void {
    const key = cellKey(row, column);
    if (!note.id.trim()) throw new Error('Review note requires an id');
    const previousId = this.notesByCell.get(key);
    const existingKey = this.noteCellById.get(note.id);
    if (existingKey !== undefined && existingKey !== key) throw new Error(`Review note identity already belongs to ${existingKey}: ${note.id}`);
    if (previousId !== undefined && previousId !== note.id) {
      this.notesById.delete(previousId);
      this.noteCellById.delete(previousId);
    }
    this.notesByCell.set(key, note.id);
    this.notesById.set(note.id, structuredClone(note));
    this.noteCellById.set(note.id, key);
    this.addReviewCellIndex(row, key);
  }

  removeNote(row: number, column: number): CellNote | undefined {
    const key = cellKey(row, column);
    const id = this.notesByCell.get(key);
    if (id === undefined) return undefined;
    const note = this.notesById.get(id);
    if (!note) throw new Error(`Review note index is dangling: ${id}`);
    this.notesByCell.delete(key);
    this.notesById.delete(id);
    this.noteCellById.delete(id);
    if (!this.threadIdsByCell.has(key)) this.removeReviewCellIndex(row, key);
    return structuredClone(note);
  }

  updateNote(row: number, column: number, updater: (note: CellNote) => void): CellNote {
    const current = this.getNoteAt(row, column);
    if (!current) throw new Error(`Review note not found at ${this.sheetId}!${row}:${column}`);
    updater(current);
    this.setNote(row, column, current);
    return structuredClone(current);
  }

  noteEntries(): ReviewNoteEntry[] {
    return [...this.notesByCell.entries()].map(([key, id]) => {
      const { row, column } = parseCellKey(key);
      const note = this.notesById.get(id);
      if (!note) throw new Error(`Review note index is dangling: ${id}`);
      return { key, row, column, note: structuredClone(note) };
    });
  }

  *entriesInRange(startRow: number, endRow: number, startColumn: number, endColumn: number): IterableIterator<ReviewCellMetadataEntry> {
    for (const { row, column, key } of this.reviewCellAddressesInRange(startRow, endRow, startColumn, endColumn)) {
      const noteId = this.notesByCell.get(key);
      const note = noteId === undefined ? undefined : this.notesById.get(noteId);
      if (noteId !== undefined && !note) throw new Error(`Review note index is dangling: ${noteId}`);
      const threadIds = this.threadIdsByCell.get(key) ?? [];
      const threads = threadIds.map((id) => {
        const thread = this.threadsById.get(id);
        if (!thread) throw new Error(`Review thread index is dangling: ${id}`);
        return structuredClone(thread);
      });
      if (!note && threads.length === 0) throw new Error(`Review cell index is dangling: ${key}`);
      yield { row, column, ...(note ? { note: structuredClone(note) } : {}), threads };
    }
  }

  *textEntriesInRange(startRow: number, endRow: number, startColumn: number, endColumn: number): IterableIterator<ReviewTextMetadataEntry> {
    for (const { row, column, key } of this.reviewCellAddressesInRange(startRow, endRow, startColumn, endColumn)) {
      const noteId = this.notesByCell.get(key);
      const note = noteId === undefined ? undefined : this.notesById.get(noteId);
      if (noteId !== undefined && !note) throw new Error(`Review note index is dangling: ${noteId}`);
      const threadIds = this.threadIdsByCell.get(key) ?? [];
      const threads = threadIds.map((id) => {
        const thread = this.threadsById.get(id);
        if (!thread) throw new Error(`Review thread index is dangling: ${id}`);
        return { id: thread.id, text: thread.text };
      });
      if (!note && threads.length === 0) throw new Error(`Review cell index is dangling: ${key}`);
      yield { row, column, ...(note ? { note: { id: note.id, text: note.text } } : {}), threads };
    }
  }

  getThread(id: string): CommentThread | undefined {
    const thread = this.threadsById.get(id);
    return thread ? structuredClone(thread) : undefined;
  }

  getThreadsAt(row: number, column: number): CommentThread[] {
    const key = cellKey(row, column);
    const ids = this.threadIdsByCell.get(key) ?? [];
    return ids.map((id) => {
      const thread = this.threadsById.get(id);
      if (!thread) throw new Error(`Review thread index is dangling: ${id}`);
      return structuredClone(thread);
    });
  }

  addThread(thread: CommentThread): void {
    if (!thread.id.trim()) throw new Error('Review thread requires an id');
    if (thread.sheetId !== this.sheetId) throw new Error(`Review thread targets another worksheet: ${thread.sheetId}`);
    if (this.threadsById.has(thread.id)) throw new Error(`Review thread already exists: ${thread.id}`);
    const key = cellKey(thread.row, thread.column);
    const ids = this.threadIdsByCell.get(key) ?? [];
    this.threadsById.set(thread.id, structuredClone(thread));
    ids.push(thread.id);
    this.threadIdsByCell.set(key, ids);
    this.addReviewCellIndex(thread.row, key);
  }

  updateThread(id: string, updater: (thread: CommentThread) => void): CommentThread {
    const existing = this.threadsById.get(id);
    if (!existing) throw new Error(`Review thread not found: ${id}`);
    const current = structuredClone(existing);
    updater(current);
    if (current.id !== id || current.sheetId !== this.sheetId) throw new Error(`Review thread identity cannot change: ${id}`);
    const previousKey = cellKey(existing.row, existing.column);
    const nextKey = cellKey(current.row, current.column);
    const nextIds = this.threadIdsByCell.get(nextKey) ?? [];
    if (previousKey !== nextKey && nextIds.includes(id)) throw new Error(`Review thread index already contains: ${id}`);
    if (previousKey !== nextKey) {
      const previousIds = (this.threadIdsByCell.get(previousKey) ?? []).filter((entry) => entry !== id);
      if (previousIds.length > 0) this.threadIdsByCell.set(previousKey, previousIds);
      else this.threadIdsByCell.delete(previousKey);
      if (previousIds.length === 0 && !this.notesByCell.has(previousKey)) this.removeReviewCellIndex(existing.row, previousKey);
      nextIds.push(id);
      this.threadIdsByCell.set(nextKey, nextIds);
      this.addReviewCellIndex(current.row, nextKey);
    }
    this.threadsById.set(id, structuredClone(current));
    return structuredClone(current);
  }

  removeThread(id: string): CommentThread | undefined {
    const current = this.threadsById.get(id);
    if (!current) return undefined;
    this.threadsById.delete(id);
    const key = cellKey(current.row, current.column);
    const ids = (this.threadIdsByCell.get(key) ?? []).filter((entry) => entry !== id);
    if (ids.length > 0) this.threadIdsByCell.set(key, ids);
    else this.threadIdsByCell.delete(key);
    if (ids.length === 0 && !this.notesByCell.has(key)) this.removeReviewCellIndex(current.row, key);
    return structuredClone(current);
  }

  threadEntries(): CommentThread[] {
    return [...this.threadsById.values()].map((thread) => structuredClone(thread));
  }

  replaceNotes(entries: ReadonlyArray<{ row: number; column: number; note: CellNote }>): void {
    this.notesByCell.clear();
    this.notesById.clear();
    this.noteCellById.clear();
    this.rebuildReviewCellIndex();
    for (const entry of entries) this.setNote(entry.row, entry.column, entry.note);
  }

  replaceThreads(threads: ReadonlyArray<CommentThread>): void {
    this.threadIdsByCell.clear();
    this.threadsById.clear();
    this.rebuildReviewCellIndex();
    for (const thread of threads) this.addThread(thread);
  }

  private addReviewCellIndex(row: number, key: ReviewCellKey): void {
    let keys = this.reviewCellsByRow.get(row);
    if (!keys) {
      keys = new Set();
      this.reviewCellsByRow.set(row, keys);
      this.sortedReviewRows = undefined;
    }
    if (keys.has(key)) return;
    keys.add(key);
    this.sortedReviewCellsByRow.delete(row);
  }

  private removeReviewCellIndex(row: number, key: ReviewCellKey): void {
    const keys = this.reviewCellsByRow.get(row);
    if (!keys || !keys.delete(key)) throw new Error(`Review cell index is missing: ${key}`);
    this.sortedReviewCellsByRow.delete(row);
    if (keys.size === 0) {
      this.reviewCellsByRow.delete(row);
      this.sortedReviewRows = undefined;
    }
  }

  private rebuildReviewCellIndex(): void {
    this.reviewCellsByRow.clear();
    this.sortedReviewRows = undefined;
    this.sortedReviewCellsByRow.clear();
    for (const key of this.notesByCell.keys()) this.addReviewCellIndex(parseCellKey(key).row, key);
    for (const key of this.threadIdsByCell.keys()) this.addReviewCellIndex(parseCellKey(key).row, key);
  }

  private firstRowAtLeast(rows: readonly number[], target: number): number {
    let low = 0;
    let high = rows.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (rows[middle]! < target) low = middle + 1;
      else high = middle;
    }
    return low;
  }

  private *reviewCellAddressesInRange(startRow: number, endRow: number, startColumn: number, endColumn: number): IterableIterator<{ row: number; column: number; key: ReviewCellKey }> {
    if (![startRow, endRow, startColumn, endColumn].every((coordinate) => Number.isSafeInteger(coordinate) && coordinate >= 0)
      || endRow < startRow || endColumn < startColumn) {
      throw new Error('Review range is invalid');
    }
    const rowCount = endRow - startRow + 1;
    const columnCount = endColumn - startColumn + 1;
    if (rowCount <= 16 && columnCount <= 16 && rowCount * columnCount <= 16) {
      for (let row = startRow; row <= endRow; row += 1) {
        const indexedCells = this.reviewCellsByRow.get(row);
        if (!indexedCells) continue;
        for (let column = startColumn; column <= endColumn; column += 1) {
          const key = cellKey(row, column);
          if (indexedCells.has(key)) yield { row, column, key };
        }
      }
      return;
    }
    const rows = this.sortedReviewRows ??= [...this.reviewCellsByRow.keys()].sort((left, right) => left - right);
    for (let index = this.firstRowAtLeast(rows, startRow); index < rows.length; index += 1) {
      const row = rows[index]!;
      if (row > endRow) break;
      const cellIndex = this.sortedReviewCellsByRow.get(row) ?? [...(this.reviewCellsByRow.get(row) ?? [])]
        .map((key) => ({ key, column: parseCellKey(key).column }))
        .sort((left, right) => left.column - right.column);
      this.sortedReviewCellsByRow.set(row, cellIndex);
      let low = 0;
      let high = cellIndex.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        if (cellIndex[middle]!.column < startColumn) low = middle + 1;
        else high = middle;
      }
      for (let cellIndexPosition = low; cellIndexPosition < cellIndex.length; cellIndexPosition += 1) {
        const entry = cellIndex[cellIndexPosition]!;
        if (entry.column > endColumn) break;
        yield { row, column: entry.column, key: entry.key };
      }
    }
  }

  remapCoordinates(mapper: (row: number, column: number) => { row: number; column: number } | undefined): void {
    this.validateRemapCoordinates(mapper);
    const notes = this.noteEntries().flatMap((entry) => {
      const mapped = mapper(entry.row, entry.column);
      return mapped ? [{ ...mapped, note: entry.note }] : [];
    });
    const threads = this.threadEntries().flatMap((thread) => {
      const mapped = mapper(thread.row, thread.column);
      return mapped ? [{ ...thread, row: mapped.row, column: mapped.column }] : [];
    });
    this.replaceNotes(notes);
    this.replaceThreads(threads);
  }

  validateRemapCoordinates(mapper: (row: number, column: number) => { row: number; column: number } | undefined): void {
    const noteKeys = new Set<string>();
    for (const entry of this.noteEntries()) {
      const mapped = mapper(entry.row, entry.column);
      if (!mapped) continue;
      const key = cellKey(mapped.row, mapped.column);
      if (noteKeys.has(key)) throw new Error(`Review note transform produced duplicate cell metadata at ${key}`);
      noteKeys.add(key);
    }
    const threadIds = new Set<string>();
    const threadKeys = new Map<string, Set<string>>();
    for (const thread of this.threadEntries()) {
      const mapped = mapper(thread.row, thread.column);
      if (!mapped) continue;
      const key = cellKey(mapped.row, mapped.column);
      const ids = threadKeys.get(key) ?? new Set<string>();
      if (ids.has(thread.id) || threadIds.has(thread.id)) throw new Error(`Review thread transform produced duplicate identity: ${thread.id}`);
      ids.add(thread.id);
      threadIds.add(thread.id);
      threadKeys.set(key, ids);
    }
  }

  reallocateIdentities(targetSheetId: string, allocateId: (sourceId: string) => string): void {
    if (!targetSheetId.trim()) throw new Error('Review identity target sheet is required');
    const notes = this.noteEntries().map((entry) => ({
      row: entry.row,
      column: entry.column,
      note: { ...entry.note, id: allocateId(entry.note.id) },
    }));
    const threads = this.threadEntries().map((thread) => ({
      ...thread,
      id: allocateId(thread.id),
      sheetId: targetSheetId,
      replies: thread.replies.map((reply) => ({ ...cloneReply(reply), id: allocateId(reply.id) })),
    }));
    this.replaceNotes(notes);
    this.replaceThreads(threads);
  }

  toSnapshot(): ReviewStoreSnapshot {
    return {
      notesByCell: Object.fromEntries(this.notesByCell),
      notesById: Object.fromEntries([...this.notesById.entries()].map(([id, note]) => [id, structuredClone(note)])),
      threadIdsByCell: Object.fromEntries([...this.threadIdsByCell.entries()].map(([key, ids]) => [key, [...ids]])),
      threadsById: Object.fromEntries([...this.threadsById.entries()].map(([id, thread]) => [id, structuredClone(thread)])),
    };
  }

  static fromSnapshot(sheetId: string, snapshot: ReviewStoreSnapshot): ReviewStore {
    const store = new ReviewStore(sheetId);
    if (!snapshot || typeof snapshot !== 'object' || Array.isArray(snapshot)) throw new Error('ReviewStore snapshot is invalid');
    const maps = ['notesByCell', 'notesById', 'threadIdsByCell', 'threadsById'] as const;
    for (const map of maps) if (!snapshot[map] || typeof snapshot[map] !== 'object' || Array.isArray(snapshot[map])) throw new Error(`ReviewStore snapshot map is invalid: ${map}`);
    for (const [id, note] of Object.entries(snapshot.notesById)) {
      if (!id.trim() || !note || typeof note !== 'object' || Array.isArray(note) || id !== note.id) throw new Error(`Review note id does not match its map key: ${id}`);
      store.notesById.set(id, structuredClone(note));
    }
    const indexedNoteIds = new Set<string>();
    for (const [key, id] of Object.entries(snapshot.notesByCell)) {
      parseCellKey(key);
      if (!store.notesById.has(id) || indexedNoteIds.has(id)) throw new Error(`Review note cell index is invalid: ${key}`);
      indexedNoteIds.add(id);
      store.notesByCell.set(key, id);
    }
    if (indexedNoteIds.size !== store.notesById.size) throw new Error('Review note store contains an unindexed note');
    for (const [id, thread] of Object.entries(snapshot.threadsById)) {
      if (!id.trim() || !thread || typeof thread !== 'object' || Array.isArray(thread) || id !== thread.id || thread.sheetId !== sheetId) throw new Error(`Review thread identity is invalid: ${id}`);
      cellKey(thread.row, thread.column);
      store.threadsById.set(id, structuredClone(thread));
    }
    for (const [key, ids] of Object.entries(snapshot.threadIdsByCell)) {
      const { row, column } = parseCellKey(key);
      if (!Array.isArray(ids) || new Set(ids).size !== ids.length) throw new Error(`Review thread cell index is invalid: ${key}`);
      for (const id of ids) {
        const thread = store.threadsById.get(id);
        if (!thread || thread.row !== row || thread.column !== column) throw new Error(`Review thread cell index references an incompatible thread: ${id}`);
      }
      store.threadIdsByCell.set(key, [...ids]);
    }
    for (const thread of store.threadsById.values()) {
      const key = cellKey(thread.row, thread.column);
      if (!(store.threadIdsByCell.get(key) ?? []).includes(thread.id)) throw new Error(`Review thread is missing its cell index: ${thread.id}`);
    }
    return store;
  }
}

export function reviewCellKey(row: number, column: number): ReviewCellKey {
  return cellKey(row, column);
}
