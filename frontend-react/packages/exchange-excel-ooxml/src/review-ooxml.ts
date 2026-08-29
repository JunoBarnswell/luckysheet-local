import type { CellNote, CommentReply, CommentThread } from '@react-sheets/core-model';
import type { NativeRelationship } from './types';
import { child, children, descendants, encodeXml, localName, parseXml, textContent, type XmlNode } from './xml';

const REL_LEGACY_COMMENTS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const REL_THREADED_COMMENTS = 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment';
const REL_PERSONS = 'http://schemas.microsoft.com/office/2017/10/relationships/person';

export interface ReviewOoxmlSnapshot {
  notesByCell: Record<string, string>;
  notesById: Record<string, CellNote>;
  threadIdsByCell: Record<string, string[]>;
  threadsById: Record<string, CommentThread>;
}

export interface ReviewOoxmlReadInput {
  sheetId: string;
  sheetPart: string;
  files: Record<string, Uint8Array>;
  relationships: readonly NativeRelationship[];
}

export interface ReviewOoxmlWriteInput {
  sheetId: string;
  notes: readonly { row: number; column: number; note: CellNote }[];
  threads: readonly CommentThread[];
}

export interface ReviewOoxmlWriteResult {
  commentsXml?: string;
  threadedCommentsXml?: string;
  personsXml?: string;
}

export function emptyReviewOoxmlSnapshot(): ReviewOoxmlSnapshot {
  return { notesByCell: {}, notesById: {}, threadIdsByCell: {}, threadsById: {} };
}

/** Read both legacy Notes and modern threaded Comments through typed ownership. */
export function readReviewOoxml(input: ReviewOoxmlReadInput): ReviewOoxmlSnapshot {
  const review = emptyReviewOoxmlSnapshot();
  const legacy = input.relationships.find((entry) => relationshipKind(entry.type) === relationshipKind(REL_LEGACY_COMMENTS));
  if (legacy) {
    const part = resolveTarget(input.sheetPart, legacy.target);
    const bytes = input.files[part];
    if (!bytes) throw new Error(`NATIVE_REVIEW_RESOURCE_MISSING: ${part}`);
    const root = firstRoot(parseXml(new TextDecoder().decode(bytes)), 'comments', part);
    const authors = children(child(root, 'authors'), 'author').map(textContent);
    for (const comment of children(child(root, 'commentList'), 'comment')) {
      const ref = parseA1(comment.attrs.ref);
      if (!ref) throw new Error(`NATIVE_REVIEW_REFERENCE_INVALID: ${comment.attrs.ref ?? ''}`);
      const key = `${ref.row}:${ref.column}`;
      if (review.notesByCell[key]) throw new Error(`NATIVE_REVIEW_DUPLICATE_NOTE: ${input.sheetId}!${key}`);
      const id = `note-${input.sheetId}-${ref.row}-${ref.column}`;
      review.notesByCell[key] = id;
      review.notesById[id] = { id, author: authors[Number(comment.attrs.author) || 0] ?? 'Unknown', text: descendants(comment, 't').map(textContent).join(''), createdAt: new Date(0).toISOString(), visible: false };
    }
  }
  const personsById = new Map<string, string>();
  const threaded = input.relationships.find((entry) => relationshipKind(entry.type) === relationshipKind(REL_THREADED_COMMENTS));
  if (threaded) {
    const threadedPart = resolveTarget(input.sheetPart, threaded.target);
    const bytes = input.files[threadedPart];
    if (!bytes) throw new Error(`NATIVE_REVIEW_RESOURCE_MISSING: ${threadedPart}`);
    const threadedRoot = firstRoot(parseXml(new TextDecoder().decode(bytes)), 'threadedComments', threadedPart);
    const personRelation = findRelated(input.files, threadedPart, REL_PERSONS);
    if (personRelation) {
      const personPart = resolveTarget(threadedPart, personRelation.target);
      const personBytes = input.files[personPart];
      if (!personBytes) throw new Error(`NATIVE_REVIEW_RESOURCE_MISSING: ${personPart}`);
      const peopleRoot = firstRoot(parseXml(new TextDecoder().decode(personBytes)), 'personList', personPart);
      for (const person of children(peopleRoot, 'person')) if (person.attrs.id) personsById.set(person.attrs.id, person.attrs.displayName ?? 'Unknown');
    }
    const roots = new Map<string, CommentThread>();
    const replies: Array<{ parentId: string; reply: CommentReply }> = [];
    for (const entry of children(threadedRoot, 'threadedComment')) {
      const id = normalizeOoxmlIdentity(entry.attrs.id, `${threadedPart} threadedComment`);
      const parentId = entry.attrs.parentId;
      const ref = parseA1(entry.attrs.ref);
      if (!ref) throw new Error(`NATIVE_REVIEW_REFERENCE_INVALID: ${entry.attrs.ref ?? ''}`);
      const author = personsById.get(entry.attrs.personId ?? '') ?? 'Unknown';
      const text = descendants(entry, 't').map(textContent).join('') || textContent(child(entry, 'text'));
      const createdAt = entry.attrs.dT && !Number.isNaN(Date.parse(entry.attrs.dT)) ? entry.attrs.dT : new Date(0).toISOString();
      if (parentId) replies.push({ parentId, reply: { id, author, text, createdAt } });
      else {
        const thread: CommentThread = { id, sheetId: input.sheetId, row: ref.row, column: ref.column, author, text, createdAt, replies: [], ...(entry.attrs.done === '1' ? { resolved: true, resolvedAt: createdAt } : {}) };
        roots.set(id, thread);
      }
    }
    for (const entry of replies) {
      const thread = roots.get(entry.parentId);
      if (!thread) throw new Error(`NATIVE_REVIEW_PARENT_MISSING: ${entry.parentId}`);
      thread.replies.push(entry.reply);
    }
    for (const thread of roots.values()) {
      const key = `${thread.row}:${thread.column}`;
      if (review.threadsById[thread.id]) throw new Error(`NATIVE_REVIEW_DUPLICATE_THREAD: ${thread.id}`);
      review.threadsById[thread.id] = thread;
      review.threadIdsByCell[key] ??= [];
      review.threadIdsByCell[key]!.push(thread.id);
    }
  }
  return review;
}

/** Serialize canonical review state. Relationship/part naming stays with the package writer. */
export function writeReviewOoxml(input: ReviewOoxmlWriteInput): ReviewOoxmlWriteResult {
  for (const entry of input.notes) {
    assertCoordinate(entry.row, entry.column, 'note');
    if (!entry.note.id.trim()) throw new Error('NATIVE_REVIEW_NOTE_ID_REQUIRED');
  }
  for (const thread of input.threads) {
    assertCoordinate(thread.row, thread.column, 'thread');
    if (thread.sheetId !== input.sheetId) throw new Error(`NATIVE_REVIEW_SHEET_MISMATCH: ${thread.id}`);
    if (!thread.id.trim()) throw new Error('NATIVE_REVIEW_THREAD_ID_REQUIRED');
    for (const reply of thread.replies) if (!reply.id.trim()) throw new Error(`NATIVE_REVIEW_REPLY_ID_REQUIRED: ${thread.id}`);
  }
  const commentsXml = input.notes.length ? serializeLegacyNotes(input) : undefined;
  const people = [...new Set(input.threads.flatMap((thread) => [thread.author, ...thread.replies.map((reply) => reply.author)]))].sort();
  const personIds = new Map(people.map((person) => [person, `person-${stableToken(person)}`]));
  const threadedCommentsXml = input.threads.length ? serializeThreadedComments(input, personIds) : undefined;
  const personsXml = input.threads.length ? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><personList xmlns="http://schemas.microsoft.com/office/2017/10/people">${people.map((person) => `<person displayName="${encodeXml(person)}" id="${encodeXml(personIds.get(person)!)}"/>`).join('')}</personList>` : undefined;
  return { commentsXml, threadedCommentsXml, personsXml };
}

function serializeLegacyNotes(input: ReviewOoxmlWriteInput): string {
  const authors = [...new Set(input.notes.map((entry) => entry.note.author))].sort();
  const authorIds = new Map(authors.map((author, index) => [author, index]));
  const comments = input.notes.map((entry) => `<comment ref="${cellA1(entry.row, entry.column)}" author="${authorIds.get(entry.note.author) ?? 0}"><text><t>${encodeXml(entry.note.text)}</t></text></comment>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authors.map((author) => `<author>${encodeXml(author)}</author>`).join('')}</authors><commentList>${comments}</commentList></comments>`;
}

function serializeThreadedComments(input: ReviewOoxmlWriteInput, personIds: ReadonlyMap<string, string>): string {
  const comments = input.threads.flatMap((thread) => {
    const root = `<threadedComment ref="${cellA1(thread.row, thread.column)}" id="${encodeXml(thread.id)}" personId="${encodeXml(personIds.get(thread.author) ?? '')}" dT="${encodeXml(thread.createdAt)}"${thread.resolved ? ' done="1"' : ''}><text>${encodeXml(thread.text)}</text></threadedComment>`;
    const replies = thread.replies.map((reply) => `<threadedComment ref="${cellA1(thread.row, thread.column)}" id="${encodeXml(reply.id)}" parentId="${encodeXml(thread.id)}" personId="${encodeXml(personIds.get(reply.author) ?? '')}" dT="${encodeXml(reply.createdAt)}"><text>${encodeXml(reply.text)}</text></threadedComment>`);
    return [root, ...replies];
  }).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><threadedComments xmlns="http://schemas.microsoft.com/office/2017/10/threadedcomments">${comments}</threadedComments>`;
}

function findRelated(files: Record<string, Uint8Array>, sourcePart: string, type: string): NativeRelationship | undefined {
  const relPart = relationshipPartName(sourcePart);
  const bytes = files[relPart];
  if (!bytes) return undefined;
  const root = firstRoot(parseXml(new TextDecoder().decode(bytes)), 'Relationships', relPart);
  return children(root, 'Relationship').map((entry) => ({ id: entry.attrs.Id ?? '', type: entry.attrs.Type ?? '', target: entry.attrs.Target ?? '', ...(entry.attrs.TargetMode ? { targetMode: entry.attrs.TargetMode } : {}) })).find((entry) => relationshipKind(entry.type) === relationshipKind(type));
}

function firstRoot(document: XmlNode, expected: string, part: string): XmlNode {
  const root = document.children[0];
  if (!root || localName(root.name) !== expected) throw new Error(`NATIVE_REVIEW_ROOT_INVALID: ${part}`);
  return root;
}

function normalizeOoxmlIdentity(value: string | undefined, label: string): string {
  if (!value?.trim()) throw new Error(`NATIVE_REVIEW_ID_INVALID: ${label}`);
  return value;
}

function parseA1(value: string | undefined): { row: number; column: number } | undefined {
  if (!value) return undefined;
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(value);
  if (!match) return undefined;
  let column = 0;
  for (const character of match[1]!) column = column * 26 + character.toUpperCase().charCodeAt(0) - 64;
  const row = Number(match[2]) - 1;
  const result = { row, column: column - 1 };
  return result.row >= 0 && result.column >= 0 ? result : undefined;
}

function cellA1(row: number, column: number): string {
  let current = column + 1;
  let letters = '';
  while (current > 0) { const remainder = (current - 1) % 26; letters = String.fromCharCode(65 + remainder) + letters; current = Math.floor((current - 1) / 26); }
  return `${letters}${row + 1}`;
}

function assertCoordinate(row: number, column: number, subject: string): void {
  if (!Number.isSafeInteger(row) || row < 0 || !Number.isSafeInteger(column) || column < 0) throw new Error(`NATIVE_REVIEW_COORDINATE_INVALID: ${subject}`);
}

function stableToken(value: string): string {
  let hash = 0x811c9dc5;
  for (const character of value) { hash ^= character.charCodeAt(0); hash = Math.imul(hash, 0x01000193); }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function relationshipKind(type: string): string {
  const normalized = type.replace(/\/+$/, '');
  return normalized.slice(normalized.lastIndexOf('/') + 1).toLowerCase();
}

function relationshipPartName(source: string): string {
  const slash = source.lastIndexOf('/');
  return slash < 0 ? `_rels/${source}.rels` : `${source.slice(0, slash)}/_rels/${source.slice(slash + 1)}.rels`;
}

function resolveTarget(source: string, target: string): string {
  if (target.startsWith('/')) return normalizePart(target.slice(1));
  const base = source.includes('/') ? source.slice(0, source.lastIndexOf('/') + 1) : '';
  return normalizePart(`${base}${target}`);
}

function normalizePart(value: string): string {
  const parts: string[] = [];
  for (const part of value.replaceAll('\\', '/').split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') { if (!parts.length) throw new Error(`Unsafe review relationship target: ${value}`); parts.pop(); }
    else parts.push(part);
  }
  return parts.join('/');
}
