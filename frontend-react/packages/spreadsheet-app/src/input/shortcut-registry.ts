export type ShortcutScope = 'grid' | 'cell-editor' | 'formula-bar' | 'dialog' | 'pivot' | 'drawing';

export interface ShortcutEventLike {
  key: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
  altKey?: boolean;
}

export interface ShortcutContext {
  scope: ShortcutScope;
  formulaReferenceSelected?: boolean;
  canRepeat?: boolean;
}

export interface ShortcutBinding {
  id: string;
  scopes: readonly ShortcutScope[];
  key: string;
  primary?: boolean;
  shift?: boolean;
  alt?: boolean;
  when?: (context: ShortcutContext) => boolean;
}

export interface ResolvedShortcut {
  id: string;
  preventDefault: boolean;
}

function equalKey(left: string, right: string): boolean {
  return left.length === 1 ? left.toLocaleLowerCase() === right.toLocaleLowerCase() : left === right;
}

/**
 * Declarative shortcut resolver shared by Canvas, formula bar and dialogs.
 * It owns matching only; command dispatch remains owned by the session/UI
 * catalog so a shortcut cannot bypass permissions or context validation.
 */
export class ShortcutRegistry {
  private readonly bindings: ShortcutBinding[] = [];

  register(binding: ShortcutBinding): void {
    if (!binding.id || !binding.key || binding.scopes.length === 0) throw new Error('Shortcut binding requires id, key, and scope');
    if (this.bindings.some((entry) => entry.id === binding.id)) throw new Error(`Shortcut binding already exists: ${binding.id}`);
    this.bindings.push({ ...binding, scopes: [...binding.scopes] });
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
      return { id: binding.id, preventDefault: true };
    }
    return undefined;
  }
}

export function createSpreadsheetShortcutRegistry(): ShortcutRegistry {
  const registry = new ShortcutRegistry();
  const grid = ['grid'] as const;
  const editable = ['cell-editor', 'formula-bar'] as const;
  registry.register({ id: 'history.undo', scopes: grid, key: 'z', primary: true });
  registry.register({ id: 'history.redo', scopes: grid, key: 'y', primary: true });
  registry.register({ id: 'clipboard.copy', scopes: grid, key: 'c', primary: true });
  registry.register({ id: 'clipboard.cut', scopes: grid, key: 'x', primary: true });
  registry.register({ id: 'clipboard.paste', scopes: grid, key: 'v', primary: true });
  registry.register({ id: 'workbook.save', scopes: grid, key: 's', primary: true });
  registry.register({ id: 'format.bold', scopes: grid, key: 'b', primary: true });
  registry.register({ id: 'format.italic', scopes: grid, key: 'i', primary: true });
  registry.register({ id: 'format.underline', scopes: grid, key: 'u', primary: true });
  registry.register({ id: 'find.open', scopes: grid, key: 'f', primary: true });
  registry.register({ id: 'replace.open', scopes: grid, key: 'h', primary: true });
  registry.register({ id: 'name.goto', scopes: grid, key: 'g', primary: true });
  registry.register({ id: 'format.cells', scopes: grid, key: '1', primary: true });
  registry.register({ id: 'hyperlink.insert', scopes: grid, key: 'k', primary: true });
  registry.register({ id: 'column.select', scopes: grid, key: ' ', primary: true });
  registry.register({ id: 'row.select', scopes: grid, key: ' ', shift: true });
  registry.register({ id: 'sheet.previous', scopes: grid, key: 'PageUp', primary: true });
  registry.register({ id: 'sheet.next', scopes: grid, key: 'PageDown', primary: true });
  registry.register({ id: 'navigation.home', scopes: grid, key: 'Home', primary: true });
  registry.register({ id: 'navigation.end', scopes: grid, key: 'End', primary: true });
  registry.register({ id: 'formula.autoSum', scopes: grid, key: '=', alt: true });
  registry.register({ id: 'formula.functionWizard', scopes: grid, key: 'F3', shift: true });
  registry.register({ id: 'edit.begin', scopes: grid, key: 'F2' });
  registry.register({ id: 'formula.toggleAbsolute', scopes: editable, key: 'F4', when: (context) => Boolean(context.formulaReferenceSelected) });
  registry.register({ id: 'history.repeat', scopes: grid, key: 'F4', when: (context) => Boolean(context.canRepeat) });
  registry.register({ id: 'formula.calculate', scopes: grid, key: 'F9' });
  registry.register({ id: 'context.open', scopes: grid, key: 'F10', shift: true });
  return registry;
}
