import type {
  CanonicalKeyGesture,
  CellEditCaret,
  CellEditCommitRequest,
  CellEditDispatchResult,
  CellEditDraft,
  CellEditEffect,
  CellEditFailure,
  CellEditIntent,
  CellEditSession,
  CellEditSnapshot,
  CellEditorStatus,
  FormulaAutocompleteCandidate,
} from './contracts';
import type { RichTextRun, RichTextRunStyle } from '@react-sheets/core-model';
import { replaceReferenceAtCaret, rewriteAbsoluteReferenceAtCaret } from './formula-edit';
import { richTextSelectionHasFlag } from './rich-text';

const READY_SNAPSHOT: CellEditSnapshot = Object.freeze({ revision: 0, status: 'ready', session: null });

function clampOffset(value: number, length: number): number {
  return Math.max(0, Math.min(length, Number.isSafeInteger(value) ? value : length));
}

function normalizeCaret(caret: CellEditCaret, length: number): CellEditCaret {
  return {
    start: clampOffset(caret.start, length),
    end: clampOffset(caret.end, length),
  };
}

function cloneDraft(draft: CellEditDraft): CellEditDraft {
  return draft.kind === 'plain'
    ? { kind: 'plain', text: draft.text }
    : { kind: 'rich-text', text: draft.text, runs: structuredClone(draft.runs) };
}

function sameRunStyle(left: RichTextRun, right: RichTextRun): boolean {
  return JSON.stringify(left.style ?? null) === JSON.stringify(right.style ?? null)
    && JSON.stringify(left.preservedProperties ?? null) === JSON.stringify(right.preservedProperties ?? null);
}

function normalizeRuns(runs: readonly RichTextRun[]): RichTextRun[] {
  const normalized: RichTextRun[] = [];
  for (const run of runs) {
    if (!run.text) continue;
    const previous = normalized.at(-1);
    if (previous && sameRunStyle(previous, run)) previous.text += run.text;
    else normalized.push(structuredClone(run));
  }
  return normalized;
}

function sliceRuns(runs: readonly RichTextRun[], start: number, end: number): RichTextRun[] {
  if (start >= end) return [];
  const result: RichTextRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const runStart = offset;
    const runEnd = offset + run.text.length;
    offset = runEnd;
    const sliceStart = Math.max(start, runStart);
    const sliceEnd = Math.min(end, runEnd);
    if (sliceStart >= sliceEnd) continue;
    result.push({ ...structuredClone(run), text: run.text.slice(sliceStart - runStart, sliceEnd - runStart) });
  }
  return result;
}

function runStyleAt(runs: readonly RichTextRun[], offset: number): Pick<RichTextRun, 'style' | 'preservedProperties'> {
  let cursor = 0;
  for (const run of runs) {
    const end = cursor + run.text.length;
    if (offset >= cursor && offset < end) return { ...(run.style ? { style: structuredClone(run.style) } : {}), ...(run.preservedProperties ? { preservedProperties: [...run.preservedProperties] } : {}) };
    cursor = end;
  }
  return {};
}

function replaceRichTextRange(draft: Extract<CellEditDraft, { kind: 'rich-text' }>, start: number, end: number, insertion: string): Extract<CellEditDraft, { kind: 'rich-text' }> {
  const before = sliceRuns(draft.runs, 0, start);
  const after = sliceRuns(draft.runs, end, draft.text.length);
  const sourceOffset = start > 0 ? start - 1 : start;
  const inserted = insertion ? [{ text: insertion, ...runStyleAt(draft.runs, sourceOffset) }] : [];
  const text = `${draft.text.slice(0, start)}${insertion}${draft.text.slice(end)}`;
  return { kind: 'rich-text', text, runs: normalizeRuns([...before, ...inserted, ...after]) };
}

function replaceRichTextByDiff(draft: Extract<CellEditDraft, { kind: 'rich-text' }>, text: string): Extract<CellEditDraft, { kind: 'rich-text' }> {
  if (text === draft.text) return { kind: 'rich-text', text, runs: structuredClone(draft.runs) };
  let prefix = 0;
  while (prefix < draft.text.length && prefix < text.length && draft.text[prefix] === text[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < draft.text.length - prefix && suffix < text.length - prefix && draft.text[draft.text.length - 1 - suffix] === text[text.length - 1 - suffix]) suffix += 1;
  return replaceRichTextRange(draft, prefix, draft.text.length - suffix, text.slice(prefix, text.length - suffix));
}

function replaceDraftText(draft: CellEditDraft, text: string): CellEditDraft {
  if (draft.kind === 'plain') return { kind: 'plain', text };
  return replaceRichTextByDiff(draft, text);
}

function textOf(draft: CellEditDraft): string {
  return draft.text;
}

function replaceSelection(session: CellEditSession, insertion: string): CellEditSession {
  const draft = textOf(session.draft);
  const start = Math.min(session.caret.start, session.caret.end);
  let end = Math.max(session.caret.start, session.caret.end);
  if (start === end && session.overtype && insertion.length > 0) end = Math.min(draft.length, start + insertion.length);
  const text = `${draft.slice(0, start)}${insertion}${draft.slice(end)}`;
  const caret = start + insertion.length;
  return {
    ...session,
    draft: session.draft.kind === 'rich-text' ? replaceRichTextRange(session.draft, start, end, insertion) : replaceDraftText(session.draft, text),
    caret: { start: caret, end: caret },
    dirty: true,
    validation: { kind: 'idle' },
    overlay: session.overlay.kind === 'validation-confirmation' || session.overlay.kind === 'value-autocomplete' ? { kind: 'none' } : session.overlay,
  };
}

function deleteBackward(session: CellEditSession): CellEditSession {
  const start = Math.min(session.caret.start, session.caret.end);
  const end = Math.max(session.caret.start, session.caret.end);
  if (start !== end) return replaceSelection(session, '');
  if (start === 0) return session;
  return replaceSelection({ ...session, caret: { start: start - 1, end } }, '');
}

function deleteForward(session: CellEditSession): CellEditSession {
  const start = Math.min(session.caret.start, session.caret.end);
  const end = Math.max(session.caret.start, session.caret.end);
  if (start !== end) return replaceSelection(session, '');
  if (end >= textOf(session.draft).length) return session;
  return replaceSelection({ ...session, caret: { start, end: end + 1 } }, '');
}

function moveAutocomplete(candidates: readonly FormulaAutocompleteCandidate[], activeIndex: number, delta: number): number {
  if (candidates.length === 0) return 0;
  return (activeIndex + delta + candidates.length) % candidates.length;
}

function moveListIndex(length: number, activeIndex: number, delta: number): number {
  if (length === 0) return 0;
  return (activeIndex + delta + length) % length;
}

function isFormulaSession(session: CellEditSession): boolean {
  return session.adapterKind === 'formula' || textOf(session.draft).startsWith('=');
}

function commitRequest(session: CellEditSession, moveAfter: CellEditCommitRequest['moveAfter'], toSelection: boolean, validationConfirmation = false): CellEditCommitRequest {
  return {
    target: session.target,
    draft: cloneDraft(session.draft),
    adapterKind: session.adapterKind,
    moveAfter,
    toSelection,
    validationConfirmation,
    baseCellFingerprint: session.baseCellFingerprint,
    originalSelection: structuredClone(session.originalSelection),
    groupedSheetIds: [...session.groupedSheetIds],
  };
}

function keyboardCommit(session: CellEditSession, gesture: CanonicalKeyGesture): CellEditCommitRequest | null {
  if (gesture.key === 'Enter') {
    if (gesture.alt) return null;
    if (gesture.ctrl || gesture.meta) return commitRequest(session, 'none', true);
    const opposite: Record<CellEditCommitRequest['moveAfter'], CellEditCommitRequest['moveAfter']> = { down: 'up', up: 'down', left: 'right', right: 'left', none: 'none' };
    return commitRequest(session, gesture.shift ? opposite[session.enterMove] : session.enterMove, false);
  }
  if (gesture.key === 'Tab') return commitRequest(session, gesture.shift ? 'left' : 'right', false);
  return null;
}

function previousWordBoundary(text: string, offset: number): number {
  let cursor = Math.max(0, offset - 1);
  while (cursor > 0 && /\s/u.test(text[cursor]!)) cursor -= 1;
  while (cursor > 0 && !/\s/u.test(text[cursor - 1]!)) cursor -= 1;
  return cursor;
}

function nextWordBoundary(text: string, offset: number): number {
  let cursor = Math.min(text.length, offset);
  while (cursor < text.length && !/\s/u.test(text[cursor]!)) cursor += 1;
  while (cursor < text.length && /\s/u.test(text[cursor]!)) cursor += 1;
  return cursor;
}

function verticalCaret(text: string, offset: number, direction: -1 | 1): number {
  const lineStart = text.lastIndexOf('\n', Math.max(0, offset - 1)) + 1;
  const column = offset - lineStart;
  if (direction < 0) {
    if (lineStart === 0) return 0;
    const previousEnd = lineStart - 1;
    const previousStart = text.lastIndexOf('\n', Math.max(0, previousEnd - 1)) + 1;
    return Math.min(previousEnd, previousStart + column);
  }
  const lineEnd = text.indexOf('\n', offset);
  if (lineEnd < 0) return text.length;
  const nextStart = lineEnd + 1;
  const nextEnd = text.indexOf('\n', nextStart);
  return Math.min(nextEnd < 0 ? text.length : nextEnd, nextStart + column);
}

function moveCaret(session: CellEditSession, gesture: CanonicalKeyGesture): CellEditSession {
  const text = session.draft.text;
  const anchor = session.caret.start;
  const focus = session.caret.end;
  let next = focus;
  if (!gesture.shift && anchor !== focus) {
    next = gesture.key === 'ArrowLeft' || gesture.key === 'ArrowUp' || gesture.key === 'Home'
      ? Math.min(anchor, focus)
      : Math.max(anchor, focus);
  } else {
    switch (gesture.key) {
      case 'ArrowLeft':
        next = gesture.ctrl || gesture.meta ? previousWordBoundary(text, focus) : Math.max(0, focus - 1);
        break;
      case 'ArrowRight':
        next = gesture.ctrl || gesture.meta ? nextWordBoundary(text, focus) : Math.min(text.length, focus + 1);
        break;
      case 'ArrowUp':
        next = verticalCaret(text, focus, -1);
        break;
      case 'ArrowDown':
        next = verticalCaret(text, focus, 1);
        break;
      case 'Home':
        next = gesture.ctrl || gesture.meta ? 0 : text.lastIndexOf('\n', Math.max(0, focus - 1)) + 1;
        break;
      case 'End': {
        const lineEnd = text.indexOf('\n', focus);
        next = gesture.ctrl || gesture.meta || lineEnd < 0 ? text.length : lineEnd;
        break;
      }
      default:
        return session;
    }
  }
  return { ...session, caret: gesture.shift ? { start: anchor, end: next } : { start: next, end: next } };
}

function formatRichText(draft: Extract<CellEditDraft, { kind: 'rich-text' }>, caret: CellEditCaret, style: Partial<RichTextRunStyle>): Extract<CellEditDraft, { kind: 'rich-text' }> {
  const start = Math.min(caret.start, caret.end);
  const end = Math.max(caret.start, caret.end);
  if (start === end) return draft;
  const selected = sliceRuns(draft.runs, start, end).map((run) => ({ ...run, style: { ...(run.style ?? {}), ...style } }));
  return { kind: 'rich-text', text: draft.text, runs: normalizeRuns([...sliceRuns(draft.runs, 0, start), ...selected, ...sliceRuns(draft.runs, end, draft.text.length)]) };
}

const MAX_DRAFT_HISTORY_BYTES = 4 * 1024 * 1024;
type DraftHistoryEntry =
  | { kind: 'plain-replace'; start: number; deleted: string; inserted: string; beforeCaret: CellEditCaret; afterCaret: CellEditCaret; bytes: number }
  | { kind: 'snapshot'; beforeDraft: CellEditDraft; afterDraft: CellEditDraft; beforeCaret: CellEditCaret; afterCaret: CellEditCaret; beforeAdapterKind: CellEditSession['adapterKind']; afterAdapterKind: CellEditSession['adapterKind']; beforeSurface: CellEditSession['editorSurface']; afterSurface: CellEditSession['editorSurface']; bytes: number };
interface DraftHistoryState { undo: DraftHistoryEntry[]; redo: DraftHistoryEntry[]; retainedBytes: number; droppedEntries: number; }
interface DraftGestureBase { draft: CellEditDraft; caret: CellEditCaret; adapterKind: CellEditSession['adapterKind']; editorSurface: CellEditSession['editorSurface']; }

function draftBytes(draft: CellEditDraft): number {
  return draft.text.length * 2 + (draft.kind === 'rich-text' ? JSON.stringify(draft.runs).length * 2 : 0) + 64;
}

function plainHistoryEntry(previous: CellEditSession, next: CellEditSession): DraftHistoryEntry {
  const before = previous.draft.text;
  const after = next.draft.text;
  let prefix = 0;
  while (prefix < before.length && prefix < after.length && before[prefix] === after[prefix]) prefix += 1;
  let suffix = 0;
  while (suffix < before.length - prefix && suffix < after.length - prefix && before[before.length - 1 - suffix] === after[after.length - 1 - suffix]) suffix += 1;
  const deleted = before.slice(prefix, before.length - suffix);
  const inserted = after.slice(prefix, after.length - suffix);
  return { kind: 'plain-replace', start: prefix, deleted, inserted, beforeCaret: structuredClone(previous.caret), afterCaret: structuredClone(next.caret), bytes: (deleted.length + inserted.length) * 2 + 64 };
}

function snapshotHistoryEntry(previous: CellEditSession, next: CellEditSession): DraftHistoryEntry {
  const beforeDraft = cloneDraft(previous.draft);
  const afterDraft = cloneDraft(next.draft);
  return {
    kind: 'snapshot',
    beforeDraft,
    afterDraft,
    beforeCaret: structuredClone(previous.caret),
    afterCaret: structuredClone(next.caret),
    beforeAdapterKind: previous.adapterKind,
    afterAdapterKind: next.adapterKind,
    beforeSurface: structuredClone(previous.editorSurface),
    afterSurface: structuredClone(next.editorSurface),
    bytes: draftBytes(beforeDraft) + draftBytes(afterDraft) + 128,
  };
}

function sameDraft(left: CellEditDraft, right: CellEditDraft): boolean {
  return left.text === right.text && (left.kind === 'plain' && right.kind === 'plain' || left.kind === 'rich-text' && right.kind === 'rich-text' && JSON.stringify(left.runs) === JSON.stringify(right.runs));
}

function recordDraftHistory(history: DraftHistoryState, previous: CellEditSession, next: CellEditSession): CellEditSession {
  if (sameDraft(previous.draft, next.draft) && previous.adapterKind === next.adapterKind) return next;
  const entry = previous.draft.kind === 'plain' && next.draft.kind === 'plain' && previous.adapterKind === next.adapterKind
    ? plainHistoryEntry(previous, next)
    : snapshotHistoryEntry(previous, next);
  for (const redo of history.redo) history.retainedBytes -= redo.bytes;
  history.redo.length = 0;
  history.undo.push(entry);
  history.retainedBytes += entry.bytes;
  while (history.retainedBytes > MAX_DRAFT_HISTORY_BYTES && history.undo.length > 1) {
    history.retainedBytes -= history.undo.shift()!.bytes;
    history.droppedEntries += 1;
  }
  return next;
}

function replacePlainHistory(text: string, entry: Extract<DraftHistoryEntry, { kind: 'plain-replace' }>, direction: 'undo' | 'redo'): string | null {
  const expected = direction === 'undo' ? entry.inserted : entry.deleted;
  const replacement = direction === 'undo' ? entry.deleted : entry.inserted;
  if (text.slice(entry.start, entry.start + expected.length) !== expected) return null;
  return `${text.slice(0, entry.start)}${replacement}${text.slice(entry.start + expected.length)}`;
}

function applyHistoryEntry(session: CellEditSession, entry: DraftHistoryEntry, direction: 'undo' | 'redo'): CellEditSession | null {
  if (entry.kind === 'plain-replace') {
    if (session.draft.kind !== 'plain') return null;
    const text = replacePlainHistory(session.draft.text, entry, direction);
    if (text === null) return null;
    const draft = { kind: 'plain' as const, text };
    return { ...session, draft, caret: structuredClone(direction === 'undo' ? entry.beforeCaret : entry.afterCaret), overlay: { kind: 'none' }, validation: { kind: 'idle' }, dirty: !sameDraft(draft, session.baselineDraft) };
  }
  return {
    ...session,
    draft: cloneDraft(direction === 'undo' ? entry.beforeDraft : entry.afterDraft),
    caret: structuredClone(direction === 'undo' ? entry.beforeCaret : entry.afterCaret),
    adapterKind: direction === 'undo' ? entry.beforeAdapterKind : entry.afterAdapterKind,
    editorSurface: structuredClone(direction === 'undo' ? entry.beforeSurface : entry.afterSurface),
    overlay: { kind: 'none' },
    validation: { kind: 'idle' },
    dirty: !sameDraft(direction === 'undo' ? entry.beforeDraft : entry.afterDraft, session.baselineDraft),
  };
}

export class CellEditDomain {
  private snapshot: CellEditSnapshot = READY_SNAPSHOT;
  private readonly listeners = new Set<() => void>();
  private readonly history: DraftHistoryState = { undo: [], redo: [], retainedBytes: 0, droppedEntries: 0 };
  private compositionBase: DraftGestureBase | null = null;
  private referenceGestureBase: DraftGestureBase | null = null;

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  };

  getSnapshot = (): CellEditSnapshot => this.snapshot;

  getHistoryMetrics(): Readonly<{ undoCount: number; redoCount: number; retainedBytes: number; droppedEntries: number }> {
    return { undoCount: this.history.undo.length, redoCount: this.history.redo.length, retainedBytes: this.history.retainedBytes, droppedEntries: this.history.droppedEntries };
  }

  dispatch(intent: CellEditIntent): CellEditDispatchResult {
    if (intent.type === 'begin') {
      if (this.snapshot.session) return this.result(false);
      this.history.undo.length = 0;
      this.history.redo.length = 0;
      this.history.retainedBytes = 0;
      this.history.droppedEntries = 0;
      this.compositionBase = null;
      this.referenceGestureBase = null;
      const draft = cloneDraft(intent.entry.initialDraft);
      const status = intent.entry.source === 'direct-typing' ? 'enter' : 'edit';
      const session: CellEditSession = {
        target: structuredClone(intent.entry.target),
        source: intent.entry.source,
        status,
        surface: intent.entry.surface,
        adapterKind: intent.entry.adapterKind,
        editorSurface: structuredClone(intent.entry.editorSurface),
        draft,
        baselineDraft: cloneDraft(intent.entry.beforeInitialDraft ?? intent.entry.initialDraft),
        caret: normalizeCaret(intent.entry.caret, textOf(draft).length),
        composition: { active: false, text: '' },
        overtype: false,
        referenceSelections: structuredClone(intent.entry.referenceSelections ?? []),
        activeReferenceId: intent.entry.activeReferenceId ?? null,
        originalSelection: structuredClone(intent.entry.originalSelection),
        originalCell: intent.entry.originalCell ? structuredClone(intent.entry.originalCell) : null,
        validation: { kind: 'idle' },
        overlay: intent.entry.inputMessage ? { kind: 'input-message', ...intent.entry.inputMessage } : { kind: 'none' },
        baseCellFingerprint: intent.entry.baseCellFingerprint,
        dirty: intent.entry.source === 'direct-typing',
        pendingCommit: null,
        enterMove: intent.entry.enterMove,
        groupedSheetIds: [...intent.entry.groupedSheetIds],
      };
      const activeSession = intent.entry.beforeInitialDraft
        ? recordDraftHistory(this.history, { ...session, draft: cloneDraft(intent.entry.beforeInitialDraft), caret: { start: intent.entry.beforeInitialDraft.text.length, end: intent.entry.beforeInitialDraft.text.length }, dirty: false }, session)
        : session;
      this.publish(activeSession);
      return this.result(true, [
        { type: 'focus', surface: activeSession.surface },
        { type: 'lifecycle', event: { type: 'EditStarted', target: activeSession.target, status } },
      ]);
    }

    const session = this.snapshot.session;
    if (!session) return this.result(false);

    switch (intent.type) {
      case 'text.replace': {
        const next = {
          ...session,
          draft: replaceDraftText(session.draft, intent.text),
          caret: normalizeCaret(intent.caret, intent.text.length),
          dirty: true,
          validation: { kind: 'idle' } as const,
        };
        const recorded = session.composition.active ? next : recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'text.insert': {
        const next = replaceSelection(session, intent.text);
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'text.delete-backward': {
        const next = deleteBackward(session);
        if (next === session) return this.result(true);
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'text.delete-forward': {
        const next = deleteForward(session);
        if (next === session) return this.result(true);
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'caret.set': {
        const caret = normalizeCaret(intent.caret, textOf(session.draft).length);
        if (caret.start === session.caret.start && caret.end === session.caret.end) return this.result(true);
        const next = { ...session, caret };
        this.publish(next);
        return this.changed(next);
      }
      case 'composition.start':
        if (session.composition.active) return this.result(true, [], false);
        this.compositionBase = { draft: cloneDraft(session.draft), caret: structuredClone(session.caret), adapterKind: session.adapterKind, editorSurface: structuredClone(session.editorSurface) };
        this.publish({ ...session, composition: { active: true, text: '' } });
        return this.result(true);
      case 'composition.update':
        this.publish({ ...session, composition: { active: true, text: intent.text } });
        return this.result(true);
      case 'composition.end': {
        const draft = replaceDraftText(session.draft, intent.text);
        const next = {
          ...session,
          draft,
          caret: normalizeCaret(intent.caret, intent.text.length),
          composition: { active: false, text: '' },
          dirty: true,
        };
        const base = this.compositionBase
          ? { ...session, draft: cloneDraft(this.compositionBase.draft), caret: structuredClone(this.compositionBase.caret), adapterKind: this.compositionBase.adapterKind, editorSurface: structuredClone(this.compositionBase.editorSurface), composition: { active: false, text: '' } }
          : session;
        this.compositionBase = null;
        const recorded = recordDraftHistory(this.history, base, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'surface.focus':
        if (session.surface === intent.surface) return this.result(true);
        this.publish({ ...session, surface: intent.surface });
        return this.result(true, [{ type: 'focus', surface: intent.surface }, { type: 'lifecycle', event: { type: 'EditSurfaceChanged', target: session.target, surface: intent.surface } }]);
      case 'status.toggle': {
        const status: CellEditSession['status'] = session.status === 'enter'
          ? 'edit'
          : isFormulaSession(session)
            ? session.status === 'point' ? 'edit' : 'point'
            : 'edit';
        if (status === session.status) return this.result(true);
        const next = { ...session, status };
        this.publish(next);
        return this.result(true, [{ type: 'lifecycle', event: { type: 'EditorStatusChanged', target: session.target, status } }]);
      }
      case 'reference.insert': {
        const rewritten = replaceReferenceAtCaret(textOf(session.draft), session.caret, intent.referenceText);
        const insertedSelection = { ...structuredClone(intent.selection), tokenSpan: rewritten.tokenSpan };
        const next = {
          ...session,
          draft: replaceDraftText(session.draft, rewritten.text),
          caret: rewritten.caret,
          dirty: true,
          validation: { kind: 'idle' } as const,
        };
        const existingIndex = next.referenceSelections.findIndex((selection) => selection.id === insertedSelection.id);
        const referenceSelections = existingIndex < 0
          ? [...next.referenceSelections, insertedSelection]
          : next.referenceSelections.map((selection, index) => index === existingIndex ? insertedSelection : selection);
        const point = { ...next, status: 'point' as const, referenceSelections, activeReferenceId: insertedSelection.id };
        const recorded = this.referenceGestureBase ? point : recordDraftHistory(this.history, session, point);
        this.publish(recorded);
        return this.changed(recorded, [{ type: 'lifecycle', event: { type: 'EditorStatusChanged', target: point.target, status: 'point' } }]);
      }
      case 'reference.begin': {
        if (!isFormulaSession(session)) return this.result(false);
        if (session.status === 'point') return this.result(true);
        const next = { ...session, status: 'point' as const };
        this.publish(next);
        return this.result(true, [{ type: 'lifecycle', event: { type: 'EditorStatusChanged', target: session.target, status: 'point' } }]);
      }
      case 'reference.set':
        this.publish({ ...session, referenceSelections: structuredClone(intent.selections), activeReferenceId: intent.activeReferenceId });
        return this.result(true);
      case 'reference.gesture.begin':
        if (this.referenceGestureBase) return this.result(true);
        this.referenceGestureBase = { draft: cloneDraft(session.draft), caret: structuredClone(session.caret), adapterKind: session.adapterKind, editorSurface: structuredClone(session.editorSurface) };
        return this.result(true);
      case 'reference.gesture.end': {
        if (!this.referenceGestureBase) return this.result(true);
        const base = { ...session, draft: cloneDraft(this.referenceGestureBase.draft), caret: structuredClone(this.referenceGestureBase.caret), adapterKind: this.referenceGestureBase.adapterKind, editorSurface: structuredClone(this.referenceGestureBase.editorSurface) };
        this.referenceGestureBase = null;
        const next = recordDraftHistory(this.history, base, session);
        this.publish(next);
        return this.changed(next);
      }
      case 'reference.gesture.cancel':
        if (!this.referenceGestureBase) return this.result(true);
        this.publish({ ...session, draft: cloneDraft(this.referenceGestureBase.draft), caret: structuredClone(this.referenceGestureBase.caret), adapterKind: this.referenceGestureBase.adapterKind, editorSurface: structuredClone(this.referenceGestureBase.editorSurface) });
        this.referenceGestureBase = null;
        return this.result(true);
      case 'reference.toggle-absolute': {
        const rewritten = rewriteAbsoluteReferenceAtCaret(textOf(session.draft), session.caret);
        if (!rewritten) return this.result(true);
        const next = {
          ...session,
          draft: replaceDraftText(session.draft, rewritten.text),
          caret: rewritten.caret,
          status: 'point' as const,
          dirty: true,
        };
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'autocomplete.open': {
        const activeIndex = intent.candidates.length > 0 ? 0 : -1;
        this.publish({ ...session, overlay: { kind: 'autocomplete', candidates: [...intent.candidates], activeIndex, revision: intent.revision, replacementSpan: intent.replacementSpan } });
        return this.result(true);
      }
      case 'autocomplete.close':
        if (session.overlay.kind !== 'autocomplete') return this.result(false);
        this.publish({ ...session, overlay: { kind: 'none' } });
        return this.result(true);
      case 'autocomplete.move':
        if (session.overlay.kind !== 'autocomplete') return this.result(false);
        this.publish({ ...session, overlay: { ...session.overlay, activeIndex: moveAutocomplete(session.overlay.candidates, session.overlay.activeIndex, intent.delta) } });
        return this.result(true);
      case 'autocomplete.accept': {
        if (session.overlay.kind !== 'autocomplete') return this.result(false);
        const candidate = session.overlay.candidates[session.overlay.activeIndex];
        if (!candidate) return this.result(true);
        const next = replaceSelection({ ...session, caret: session.overlay.replacementSpan, overlay: { kind: 'none' } }, candidate.insertionText);
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'function-hint.open':
        this.publish({ ...session, overlay: { kind: 'function-hint', functionName: intent.functionName, argumentIndex: intent.argumentIndex } });
        return this.result(true);
      case 'function-hint.close':
        if (session.overlay.kind !== 'function-hint') return this.result(false);
        this.publish({ ...session, overlay: { kind: 'none' } });
        return this.result(true);
      case 'value-autocomplete.apply': {
        if (!session.draft.text.startsWith(intent.prefix) || intent.candidate.length <= intent.prefix.length) return this.result(false);
        const draft = replaceDraftText(session.draft, intent.candidate);
        this.publish({ ...session, draft, caret: { start: intent.prefix.length, end: intent.candidate.length }, overlay: { kind: 'value-autocomplete', prefix: intent.prefix, candidate: intent.candidate } });
        return this.result(true);
      }
      case 'value-autocomplete.close':
        if (session.overlay.kind !== 'value-autocomplete') return this.result(false);
        this.publish({ ...session, draft: replaceDraftText(session.draft, session.overlay.prefix), caret: { start: session.overlay.prefix.length, end: session.overlay.prefix.length }, overlay: { kind: 'none' } });
        return this.result(true);
      case 'editor-list.open': {
        const activeIndex = intent.items.length > 0 ? 0 : -1;
        this.publish({ ...session, overlay: { kind: 'editor-list', items: structuredClone(intent.items), activeIndex } });
        return this.result(true);
      }
      case 'editor-list.move':
        if (session.overlay.kind !== 'editor-list') return this.result(false);
        this.publish({ ...session, overlay: { ...session.overlay, activeIndex: moveListIndex(session.overlay.items.length, session.overlay.activeIndex, intent.delta) } });
        return this.result(true);
      case 'editor-list.accept': {
        if (session.overlay.kind !== 'editor-list') return this.result(false);
        const item = session.overlay.items[intent.index ?? session.overlay.activeIndex];
        if (!item) return this.result(true);
        const draft = replaceDraftText(session.draft, item.text);
        const next = { ...session, draft, caret: { start: item.text.length, end: item.text.length }, overlay: { kind: 'none' } as const, dirty: true };
        const request = commitRequest(next, 'none', false);
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish({ ...recorded, pendingCommit: request });
        return this.changed(recorded, [{ type: 'commit', request }]);
      }
      case 'editor-list.close':
        if (session.overlay.kind !== 'editor-list') return this.result(false);
        this.publish({ ...session, overlay: { kind: 'none' } });
        return this.result(true);
      case 'rich-text.format': {
        if (session.caret.start === session.caret.end) return this.result(true);
        const richDraft = session.draft.kind === 'rich-text'
          ? session.draft
          : { kind: 'rich-text' as const, text: session.draft.text, runs: session.draft.text ? [{ text: session.draft.text }] : [] };
        const draft = formatRichText(richDraft, session.caret, intent.style);
        const next = { ...session, adapterKind: 'rich-text' as const, editorSurface: { kind: 'rich-text' as const, inputMode: 'text' as const, multiline: true }, draft, dirty: true };
        const recorded = recordDraftHistory(this.history, session, next);
        this.publish(recorded);
        return this.changed(recorded);
      }
      case 'draft.undo': {
        const entry = this.history.undo.at(-1);
        if (!entry) return this.result(true);
        const restored = applyHistoryEntry(session, entry, 'undo');
        if (!restored) {
          const failure: CellEditFailure = { code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Draft undo history no longer matches the canonical draft', target: session.target.canonical, recovery: 'Cancel the edit session rather than applying an inconsistent draft history entry.' };
          return this.result(true, [], true, failure);
        }
        this.history.undo.pop();
        this.history.redo.push(entry);
        const next = restored;
        this.publish(next);
        return this.changed(next);
      }
      case 'draft.redo': {
        const entry = this.history.redo.at(-1);
        if (!entry) return this.result(true);
        const restored = applyHistoryEntry(session, entry, 'redo');
        if (!restored) {
          const failure: CellEditFailure = { code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Draft redo history no longer matches the canonical draft', target: session.target.canonical, recovery: 'Cancel the edit session rather than applying an inconsistent draft history entry.' };
          return this.result(true, [], true, failure);
        }
        this.history.redo.pop();
        this.history.undo.push(entry);
        const next = restored;
        this.publish(next);
        return this.changed(next);
      }
      case 'keyboard':
        return this.dispatchKeyboard(session, intent.gesture);
      case 'commit': {
        const request = commitRequest(session, intent.moveAfter ?? 'down', Boolean(intent.toSelection));
        this.publish({ ...session, pendingCommit: request });
        return this.result(true, [{ type: 'commit', request }]);
      }
      case 'commit.succeeded': {
        const target = session.target;
        this.clearTransientHistory();
        this.publish(null);
        return this.result(true, [{ type: 'lifecycle', event: { type: 'EditEnded', target, committed: true } }]);
      }
      case 'commit.failed':
        return this.applyCommitFailure(session, intent.error);
      case 'validation.confirm': {
        if (session.validation.kind !== 'confirmation-required' || !session.pendingCommit) return this.result(false);
        const request = { ...session.pendingCommit, validationConfirmation: true };
        this.publish({ ...session, validation: { kind: 'idle' }, overlay: { kind: 'none' }, pendingCommit: request });
        return this.result(true, [{ type: 'commit', request }]);
      }
      case 'validation.reject':
        if (session.validation.kind !== 'confirmation-required') return this.result(false);
        this.publish({ ...session, validation: { kind: 'idle' }, overlay: { kind: 'none' }, pendingCommit: null });
        return this.result(true);
      case 'cancel': {
        const target = session.target;
        const originalSelection = structuredClone(session.originalSelection);
        this.clearTransientHistory();
        this.publish(null);
        return this.result(true, [
          { type: 'cancel', originalSelection },
          { type: 'lifecycle', event: { type: 'EditEnded', target, committed: false } },
        ]);
      }
      default:
        return this.result(false);
    }
  }

  private dispatchKeyboard(session: CellEditSession, gesture: CanonicalKeyGesture): CellEditDispatchResult {
    if (gesture.composing || session.composition.active) return this.result(true, [], false);

    if (gesture.key === 'Escape') {
      if (session.overlay.kind === 'value-autocomplete') return this.dispatch({ type: 'value-autocomplete.close' });
      if (session.overlay.kind !== 'none') {
        this.publish({ ...session, overlay: { kind: 'none' } });
        return this.result(true);
      }
      return this.dispatch({ type: 'cancel' });
    }

    if ((gesture.ctrl || gesture.meta) && !gesture.alt && gesture.key.toLocaleLowerCase() === 'z') {
      if (session.overlay.kind === 'value-autocomplete') this.dispatch({ type: 'value-autocomplete.close' });
      return this.dispatch({ type: gesture.shift ? 'draft.redo' : 'draft.undo' });
    }
    if ((gesture.ctrl || gesture.meta) && !gesture.alt && gesture.key.toLocaleLowerCase() === 'y') return this.dispatch({ type: 'draft.redo' });

    if (gesture.key === 'Enter' && gesture.alt) return this.dispatch({ type: 'text.insert', text: '\n' });

    if (gesture.key === 'Tab' && session.overlay.kind === 'autocomplete') return this.dispatch({ type: 'autocomplete.accept' });
    if (session.overlay.kind === 'editor-list') {
      if (gesture.key === 'ArrowUp') return this.dispatch({ type: 'editor-list.move', delta: -1 });
      if (gesture.key === 'ArrowDown') return this.dispatch({ type: 'editor-list.move', delta: 1 });
      if (gesture.key === 'Enter' || gesture.key === 'Tab') return this.dispatch({ type: 'editor-list.accept' });
    }

    const request = keyboardCommit(session, gesture);
    if (request) {
      this.publish({ ...session, pendingCommit: request });
      return this.result(true, [{ type: 'commit', request }]);
    }

    if (gesture.key === 'F2') return this.dispatch({ type: 'status.toggle' });
    if (gesture.key === 'F4' && isFormulaSession(session)) return this.dispatch({ type: 'reference.toggle-absolute' });
    if (gesture.key === 'Insert' && session.status === 'edit') {
      this.publish({ ...session, overtype: !session.overtype });
      return this.result(true);
    }

    if (gesture.alt && gesture.key === 'ArrowDown') {
      if (session.overlay.kind === 'autocomplete') return this.dispatch({ type: 'autocomplete.close' });
      if (session.overlay.kind === 'editor-list') return this.dispatch({ type: 'editor-list.close' });
      return this.result(true, [{ type: 'overlay.toggle-request' }]);
    }

    if (session.status === 'point' && (gesture.ctrl || gesture.meta) && (gesture.key === 'PageUp' || gesture.key === 'PageDown')) {
      return this.result(true, [{ type: 'reference.switch-sheet', direction: gesture.key === 'PageUp' ? 'previous' : 'next' }]);
    }

    if (gesture.key === 'F3' && isFormulaSession(session)) return this.result(true, [{ type: 'defined-name.request' }]);

    if ((gesture.ctrl || gesture.meta) && ['a', 'A'].includes(gesture.key)) {
      const end = session.draft.text.length;
      this.publish({ ...session, caret: { start: 0, end } });
      return this.result(true);
    }

    if ((gesture.ctrl || gesture.meta) && !gesture.alt && session.caret.start !== session.caret.end) {
      const shortcut = gesture.key.toLocaleLowerCase();
      const styleKey = shortcut === 'b' ? 'bold' : shortcut === 'i' ? 'italic' : shortcut === 'u' ? 'underline' : gesture.code === 'Digit5' ? 'strikethrough' : null;
      if (styleKey) {
        const enabled = !richTextSelectionHasFlag(session.draft, session.caret, styleKey);
        const style = styleKey === 'bold' ? { bold: enabled } : styleKey === 'italic' ? { italic: enabled } : styleKey === 'underline' ? { underline: enabled } : { strikethrough: enabled };
        return this.dispatch({ type: 'rich-text.format', style });
      }
    }

    if ((gesture.ctrl || gesture.meta) && gesture.code === 'Semicolon') {
      return this.result(true, [{ type: 'insert-current', value: gesture.shift ? 'time' : 'date' }]);
    }

    if (!gesture.ctrl && !gesture.meta && !gesture.alt && gesture.key.length === 1) {
      return this.dispatch({ type: 'text.insert', text: gesture.key });
    }

    if (gesture.key === 'Backspace') return this.dispatch({ type: 'text.delete-backward' });
    if (gesture.key === 'Delete') return this.dispatch({ type: 'text.delete-forward' });

    if (session.status === 'point' && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(gesture.key)) {
      const deltas: Record<string, readonly [number, number]> = {
        ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1],
      };
      const [rowDelta, columnDelta] = deltas[gesture.key]!;
      return this.result(true, [{ type: 'reference.move', referenceId: session.activeReferenceId, rowDelta, columnDelta, extend: gesture.shift, jump: gesture.ctrl || gesture.meta }]);
    }

    if (session.status !== 'point' && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(gesture.key)) {
      const next = moveCaret(session, gesture);
      if (next !== session) this.publish(next);
      return this.result(true);
    }

    return this.result(false);
  }

  private applyCommitFailure(session: CellEditSession, failure: CellEditFailure): CellEditDispatchResult {
    if (failure.code === 'CELL_EDIT_CONFIRMATION_REQUIRED' && (failure.alertStyle === 'warning' || failure.alertStyle === 'information')) {
      const confirmation = {
        kind: 'confirmation-required' as const,
        message: failure.message,
        alertStyle: failure.alertStyle,
        ...(failure.title ? { title: failure.title } : {}),
      };
      this.publish({
        ...session,
        validation: confirmation,
        overlay: { kind: 'validation-confirmation', message: failure.message, alertStyle: failure.alertStyle, ...(failure.title ? { title: failure.title } : {}) },
      });
      return this.result(true, [{ type: 'lifecycle', event: { type: 'ValidationError', target: session.target, code: failure.code, message: failure.message } }], true, failure);
    }

    const next = {
      ...session,
      validation: { kind: 'blocking-error' as const, code: failure.code, message: failure.message },
      pendingCommit: null,
    };
    this.publish(next);
    return this.result(true, [{ type: 'lifecycle', event: { type: 'ValidationError', target: session.target, code: failure.code, message: failure.message } }], true, failure);
  }

  private changed(session: CellEditSession, effects: readonly CellEditEffect[] = []): CellEditDispatchResult {
    return this.result(true, [
      ...effects,
      { type: 'lifecycle', event: { type: 'EditChanged', target: session.target, revision: this.snapshot.revision } },
    ]);
  }

  private result(handled: boolean, effects: readonly CellEditEffect[] = [], preventDefault = handled, failure?: CellEditFailure): CellEditDispatchResult {
    return { handled, preventDefault, status: this.snapshot.status, effects, ...(failure ? { failure } : {}) };
  }

  private clearTransientHistory(): void {
    this.history.undo.length = 0;
    this.history.redo.length = 0;
    this.history.retainedBytes = 0;
    this.history.droppedEntries = 0;
    this.compositionBase = null;
    this.referenceGestureBase = null;
  }

  private publish(session: CellEditSession | null): void {
    const revision = this.snapshot.revision + 1;
    this.snapshot = {
      revision,
      status: session?.status ?? 'ready',
      session,
    };
    for (const listener of this.listeners) listener();
  }
}
