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
  private readonly noteCellById = new Map<string, ReviewCellKey>();
  private readonly notesById = new Map<string, CellNote>();
  private readonly threadIdsByCell = new Map<ReviewCellKey, string[]>();
  private readonly threadsById = new Map<string, CommentThread>();

  constructor(public sheetId: string) {
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
      nextIds.push(id);
      this.threadIdsByCell.set(nextKey, nextIds);
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
    return structuredClone(current);
  }

  threadEntries(): CommentThread[] {
    return [...this.threadsById.values()].map((thread) => structuredClone(thread));
  }

  replaceNotes(entries: ReadonlyArray<{ row: number; column: number; note: CellNote }>): void {
    this.notesByCell.clear();
    this.noteCellById.clear();
    this.notesById.clear();
    for (const entry of entries) this.setNote(entry.row, entry.column, entry.note);
  }

  replaceThreads(threads: ReadonlyArray<CommentThread>): void {
    this.threadIdsByCell.clear();
    this.threadsById.clear();
    for (const thread of threads) this.addThread(thread);
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
    const previousSheetId = this.sheetId;
    this.sheetId = targetSheetId;
    try {
      this.replaceNotes(notes);
      this.replaceThreads(threads);
    } catch (error) {
      this.sheetId = previousSheetId;
      throw error;
    }
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
      store.noteCellById.set(id, key);
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
