import type { CanonicalKeyGesture } from '../cell-edit/contracts';

export type ShortcutScope =
  | 'grid'
  | 'cell-editor'
  | 'formula-bar'
  | 'dialog'
  | 'pivot'
  | 'table'
  | 'drawing'
  | 'drawing-text'
  | 'comment'
  | 'ribbon'
  | 'ribbon-keytip';

export interface ShortcutEventLike {
  key: string;
  code?: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
  altKey?: boolean;
  repeat?: boolean;
  composing?: boolean;
}

export interface ShortcutContext {
  scope: ShortcutScope;
  inputMode?: string;
  ribbonTab?: string;
  formulaReferenceSelected?: boolean;
  canRepeat?: boolean;
  hasClipboard?: boolean;
  keyTipPrefix?: string;
  activeObject?: boolean;
}

export interface ShortcutBinding {
  /** Unique physical gesture binding identity. */
  id: string;
  /** Canonical command dispatched by this gesture. Defaults to id. */
  commandId?: string;
  scopes: readonly ShortcutScope[];
  key: string;
  primary?: boolean;
  shift?: boolean;
  alt?: boolean;
  when?: (context: ShortcutContext) => boolean;
}

export interface ShortcutChord {
  key: string;
  primary?: boolean;
  shift?: boolean;
  alt?: boolean;
}

export interface ShortcutSequenceBinding {
  id: string;
  commandId: string;
  scopes: readonly ShortcutScope[];
  chords: readonly ShortcutChord[];
  when?: (context: ShortcutContext) => boolean;
}

export interface ShortcutSequenceState {
  active: boolean;
  index: number;
}

export interface ResolvedShortcutSequence {
  state: ShortcutSequenceState;
  shortcut?: ResolvedShortcut;
  preventDefault: boolean;
}

export interface ResolvedShortcut {
  id: string;
  preventDefault: boolean;
  scope: ShortcutScope;
}

export type { CanonicalKeyGesture };

function normalizeKey(key: string): string {
  if (key === 'Spacebar') return ' ';
  if (key === 'Esc') return 'Escape';
  return key.length === 1 ? key.toLocaleLowerCase() : key;
}

function equalKey(left: string, right: string): boolean {
  return normalizeKey(left) === normalizeKey(right);
}

/** Convert a browser keyboard event into the canonical editor gesture shape. */
export function canonicalKeyGesture(event: ShortcutEventLike): CanonicalKeyGesture {
  return {
    key: event.key,
    code: event.code ?? event.key,
    alt: Boolean(event.altKey),
    ctrl: Boolean(event.ctrlKey),
    meta: Boolean(event.metaKey),
    shift: Boolean(event.shiftKey),
    repeat: Boolean(event.repeat),
    composing: Boolean(event.composing),
  };
}

/**
 * Canonical Excel shortcut resolver shared by Canvas, formula bar, dialogs,
 * Ribbon and contextual surfaces. Matching is the only responsibility here;
 * dispatch remains owned by the session/controller so permissions and history
 * cannot be bypassed by a keyboard path.
 */
export class ShortcutRegistry {
  private readonly bindings: ShortcutBinding[] = [];
  private readonly sequences: ShortcutSequenceBinding[] = [];

  register(binding: ShortcutBinding): void {
    if (!binding.id || !binding.key || binding.scopes.length === 0) throw new Error('Shortcut binding requires id, key, and scope');
    if (this.bindings.some((entry) => entry.id === binding.id) || this.sequences.some((entry) => entry.id === binding.id)) throw new Error(`Shortcut binding already exists: ${binding.id}`);
    this.bindings.push({ ...binding, scopes: [...binding.scopes] });
  }

  registerSequence(binding: ShortcutSequenceBinding): void {
    if (!binding.id || !binding.commandId || binding.scopes.length === 0 || binding.chords.length < 2) throw new Error('Shortcut sequence requires id, command, scopes, and at least two chords');
    if (this.bindings.some((entry) => entry.id === binding.id) || this.sequences.some((entry) => entry.id === binding.id)) throw new Error(`Shortcut binding already exists: ${binding.id}`);
    if (binding.chords.some((chord) => !chord.key)) throw new Error(`Shortcut sequence contains an empty chord: ${binding.id}`);
    this.sequences.push({ ...binding, scopes: [...binding.scopes], chords: binding.chords.map((chord) => ({ ...chord })) });
  }

  listBindings(): readonly ShortcutBinding[] {
    return this.bindings.map((binding) => ({ ...binding, scopes: [...binding.scopes] }));
  }

  listSequenceBindings(): readonly ShortcutSequenceBinding[] {
    return this.sequences.map((binding) => ({ ...binding, scopes: [...binding.scopes], chords: binding.chords.map((chord) => ({ ...chord })) }));
  }

  resolve(event: ShortcutEventLike, context: ShortcutContext): ResolvedShortcut | undefined {
    const primary = Boolean(event.ctrlKey || event.metaKey);
    for (const binding of this.bindings) {
      if (!binding.scopes.includes(context.scope)) continue;
      if (!equalKey(binding.key, event.key)) continue;
      if (Boolean(binding.primary) !== primary) continue;
      if (Boolean(binding.shift) !== Boolean(event.shiftKey)) continue;
      if (Boolean(binding.alt) !== Boolean(event.altKey)) continue;
      if (binding.when && !binding.when(context)) continue;
      return { id: binding.commandId ?? binding.id, preventDefault: true, scope: context.scope };
    }
    return undefined;
  }

  resolveSequence(event: ShortcutEventLike, context: ShortcutContext, state: ShortcutSequenceState = { active: false, index: 0 }): ResolvedShortcutSequence {
    if (event.key === 'Escape') return { state: { active: false, index: 0 }, preventDefault: state.active };
    const candidates = this.sequences.filter((binding) => binding.scopes.includes(context.scope) && (!binding.when || binding.when(context)));
    const index = state.active ? state.index : 0;
    const matching = candidates.filter((binding) => {
      const chord = binding.chords[index];
      return chord !== undefined && chordMatches(chord, event);
    });
    if (matching.length === 0) return { state: { active: false, index: 0 }, preventDefault: state.active };
    const completed = matching.find((binding) => binding.chords.length === index + 1);
    if (completed) return { state: { active: false, index: 0 }, shortcut: { id: completed.commandId, preventDefault: true, scope: context.scope }, preventDefault: true };
    return { state: { active: true, index: index + 1 }, preventDefault: true };
  }
}

function chordMatches(chord: ShortcutChord, event: ShortcutEventLike): boolean {
  const primary = Boolean(event.ctrlKey || event.metaKey);
  return equalKey(chord.key, event.key)
    && Boolean(chord.primary) === primary
    && Boolean(chord.shift) === Boolean(event.shiftKey)
    && Boolean(chord.alt) === Boolean(event.altKey);
}

export function createSpreadsheetShortcutRegistry(): ShortcutRegistry {
  const registry = new ShortcutRegistry();
  const grid = ['grid'] as const;
  const editable = ['cell-editor', 'formula-bar'] as const;
  const formulaScopes = ['cell-editor', 'formula-bar'] as const;
  registry.register({ id: 'history.undo', scopes: grid, key: 'z', primary: true });
  registry.register({ id: 'history.redo', scopes: grid, key: 'y', primary: true });
  registry.register({ id: 'history.repeat', scopes: grid, key: 'F4', when: (context) => Boolean(context.canRepeat) });
  registry.register({ id: 'clipboard.copy', scopes: grid, key: 'c', primary: true });
  registry.register({ id: 'clipboard.cut', scopes: grid, key: 'x', primary: true });
  registry.register({ id: 'clipboard.paste', scopes: grid, key: 'v', primary: true });
  registry.register({ id: 'clipboard.pasteSpecial', scopes: grid, key: 'v', primary: true, alt: true });
  registry.register({ id: 'clipboard.cancel', scopes: grid, key: 'Escape' });
  registry.register({ id: 'workbook.save', scopes: grid, key: 's', primary: true });
  registry.register({ id: 'format.bold', scopes: [...grid, ...editable], key: 'b', primary: true });
  registry.register({ id: 'format.italic', scopes: [...grid, ...editable], key: 'i', primary: true });
  registry.register({ id: 'format.underline', scopes: [...grid, ...editable], key: 'u', primary: true });
  registry.register({ id: 'find.open', scopes: grid, key: 'f', primary: true });
  registry.register({ id: 'commandPalette.open', scopes: grid, key: 'p', primary: true, shift: true });
  registry.register({ id: 'print.preview', scopes: grid, key: 'p', primary: true });
  registry.register({ id: 'replace.open', scopes: grid, key: 'h', primary: true });
  registry.register({ id: 'name.goto', scopes: grid, key: 'g', primary: true });
  registry.register({ id: 'format.cells', scopes: grid, key: '1', primary: true });
  registry.register({ id: 'format.cells.font', scopes: grid, key: 'f', primary: true, shift: true });
  registry.register({ id: 'format.number.general', scopes: grid, key: '~', primary: true, shift: true });
  registry.register({ id: 'format.number.currency', scopes: grid, key: '$', primary: true, shift: true });
  registry.register({ id: 'format.number.percent', scopes: grid, key: '%', primary: true, shift: true });
  registry.register({ id: 'format.number.scientific', scopes: grid, key: '^', primary: true, shift: true });
  registry.register({ id: 'format.number.date', scopes: grid, key: '#', primary: true, shift: true });
  registry.register({ id: 'format.number.time', scopes: grid, key: '@', primary: true, shift: true });
  registry.register({ id: 'format.number.comma', scopes: grid, key: '!', primary: true, shift: true });
  registry.register({ id: 'range.fillDown', scopes: grid, key: 'd', primary: true });
  registry.register({ id: 'range.fillRight', scopes: grid, key: 'r', primary: true });
  registry.register({ id: 'range.flashFill', scopes: grid, key: 'e', primary: true });
  registry.register({ id: 'range.clearContents', scopes: grid, key: 'Delete' });
  registry.register({ id: 'cells.insert', scopes: grid, key: '+', primary: true, shift: true });
  registry.register({ id: 'cells.delete', scopes: grid, key: '-', primary: true });
  registry.register({ id: 'filter.toggle', scopes: grid, key: 'l', primary: true, shift: true });
  registry.register({ id: 'navigation.goto', scopes: grid, key: 'F5' });
  registry.register({ id: 'ribbon.home.keyTips', scopes: grid, key: 'h', alt: true });
  registry.register({ id: 'ribbon.insert.keyTips', scopes: grid, key: 'n', alt: true });
  registry.register({ id: 'ribbon.pageLayout.keyTips', scopes: grid, key: 'p', alt: true });
  registry.register({ id: 'ribbon.formulas.keyTips', scopes: grid, key: 'm', alt: true });
  registry.register({ id: 'ribbon.data.keyTips', scopes: grid, key: 'a', alt: true });
  registry.register({ id: 'ribbon.review.keyTips', scopes: grid, key: 'r', alt: true });
  registry.register({ id: 'ribbon.view.keyTips', scopes: grid, key: 'w', alt: true });
  registry.register({ id: 'ribbon.keyTips', scopes: grid, key: 'F10' });
  registry.register({ id: 'ribbon.toggle', scopes: ['grid', 'ribbon'], key: 'F1', primary: true });
  registry.register({ id: 'hyperlink.insert', scopes: grid, key: 'k', primary: true });
  registry.register({ id: 'column.select', scopes: grid, key: ' ', primary: true });
  registry.register({ id: 'row.select', scopes: grid, key: ' ', shift: true });
  registry.register({ id: 'selection.selectAll', scopes: grid, key: 'a', primary: true });
  registry.register({ id: 'selection.extendMode', scopes: grid, key: 'F8' });
  registry.register({ id: 'selection.addMode', scopes: grid, key: 'F8', shift: true });
  registry.register({ id: 'sheet.previous', scopes: grid, key: 'PageUp', primary: true });
  registry.register({ id: 'sheet.next', scopes: grid, key: 'PageDown', primary: true });
  registry.register({ id: 'navigation.home', scopes: grid, key: 'Home', primary: true });
  registry.register({ id: 'navigation.end', scopes: grid, key: 'End', primary: true });
  registry.register({ id: 'navigation.pageDown', scopes: grid, key: 'PageDown' });
  registry.register({ id: 'navigation.pageUp', scopes: grid, key: 'PageUp' });
  registry.register({ id: 'formula.autoSum', scopes: grid, key: '=', alt: true });
  registry.register({ id: 'formula.functionWizard', scopes: grid, key: 'F3', shift: true });
  registry.register({ id: 'edit.begin', scopes: grid, key: 'F2' });
  registry.register({ id: 'formula.toggleAbsolute', scopes: formulaScopes, key: 'F4', when: (context) => Boolean(context.formulaReferenceSelected) });
  registry.register({ id: 'formula.calculate', scopes: grid, key: 'F9' });
  registry.register({ id: 'formula.calculateSheet', scopes: grid, key: 'F9', shift: true });
  registry.register({ id: 'formula.calculateFull', scopes: grid, key: 'F9', primary: true, alt: true });
  registry.register({ id: 'formula.calculateRebuild', scopes: grid, key: 'F9', primary: true, alt: true, shift: true });
  registry.register({ id: 'formula.show', scopes: grid, key: '`', primary: true });
  registry.register({ id: 'cell.insertDate', scopes: grid, key: ';', primary: true });
  registry.register({ id: 'cell.insertTime', scopes: grid, key: ':', primary: true, shift: true });
  registry.register({ id: 'formulaBar.toggle', scopes: grid, key: 'u', primary: true, shift: true });
  registry.register({ id: 'context.open', scopes: grid, key: 'F10', shift: true });
  registry.register({ id: 'quickAnalysis.open', scopes: grid, key: 'q', primary: true });
  registry.register({ id: 'table.create', scopes: grid, key: 't', primary: true });
  registry.register({ id: 'table.create.ctrl-l', commandId: 'table.create', scopes: grid, key: 'l', primary: true });
  registry.register({ id: 'chart.insert', scopes: grid, key: 'F1', alt: true });
  registry.register({ id: 'chart.sheet.insert', scopes: grid, key: 'F11' });
  registry.register({ id: 'zoom.in', scopes: grid, key: '=', primary: true, alt: true });
  registry.register({ id: 'zoom.out', scopes: grid, key: '-', primary: true, alt: true });
  registry.register({ id: 'row.hide', scopes: grid, key: '9', primary: true });
  registry.register({ id: 'column.hide', scopes: grid, key: '0', primary: true });
  registry.register({ id: 'comment.note.edit', scopes: grid, key: 'F2', shift: true });
  registry.register({ id: 'comment.thread.open', scopes: grid, key: 'F2', primary: true, shift: true });
  registry.register({ id: 'pivot.refresh', scopes: ['pivot'], key: 'F5', alt: true });
  registry.register({ id: 'drawing.remove', scopes: ['drawing'], key: 'Delete' });
  registry.register({ id: 'drawing.selectAll', scopes: ['drawing'], key: ' ', primary: true, shift: true });
  registry.registerSequence({ id: 'clipboard.pasteSpecial.alt-e-s', commandId: 'clipboard.pasteSpecial', scopes: grid, chords: [{ key: 'e', alt: true }, { key: 's' }] });
  return registry;
}
