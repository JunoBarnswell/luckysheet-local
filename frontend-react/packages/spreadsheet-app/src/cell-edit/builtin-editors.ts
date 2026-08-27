import { checkboxStateFromValue, checkboxValueForState, type CellData, type RichTextRun } from '@react-sheets/core-model';
import { parseFormula } from '@react-sheets/formula-engine';
import { interpretCellInput } from '@react-sheets/sheet-features';
import type {
  CellEditCommitPayload,
  CellEditorBehavior,
  CellEditorContext,
  CellEditorValidationResult,
} from './editor-registry';
import { CellEditorRegistry } from './editor-registry';
import type { CellEditDraft, CellEditIntent, CellEditorRuntimeKind } from './contracts';

function cellText(cell: CellData | null): string {
  return cell?.formula ?? (cell?.value == null ? '' : String(cell.value));
}

function plainDraft(context: CellEditorContext): CellEditDraft {
  return { kind: 'plain', text: context.initialText ?? cellText(context.cell) };
}

function unchanged(_intent: CellEditIntent, draft: CellEditDraft): CellEditDraft {
  return draft;
}

function allowed(): { allowed: true } {
  return { allowed: true };
}

function valid(): { valid: true } {
  return { valid: true };
}

function rawText(draft: CellEditDraft): CellEditCommitPayload {
  return { kind: 'raw-text', text: draft.text };
}

function plainBehavior(kind: CellEditorRuntimeKind): CellEditorBehavior {
  return {
    kind,
    ...(kind === 'text' ? { valueAutocomplete: true } : {}),
    surface: { kind: kind === 'rich-text' ? 'rich-text' : kind === 'validation-list' || kind === 'combo-box' ? 'list' : kind === 'checkbox' ? 'checkbox' : 'text', inputMode: kind === 'number' ? 'decimal' : kind === 'datetime' ? 'numeric' : 'text', multiline: kind === 'text' || kind === 'rich-text' },
    canEnter: allowed,
    createDraft: plainDraft,
    reduce: unchanged,
    ownsKey: () => false,
    validate: valid,
    toCommitPayload: rawText,
  };
}

const numberBehavior: CellEditorBehavior = {
  ...plainBehavior('number'),
  validate: (draft, context) => {
    const interpreted = interpretCellInput(draft.text, context.inputContext);
    return interpreted.valueType === 'number' || interpreted.valueType === 'empty'
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Number editor input is not a canonical number.', recovery: 'Enter a number accepted by the workbook culture and number format.' };
  },
};

const dateTimeBehavior: CellEditorBehavior = {
  ...plainBehavior('datetime'),
  validate: (draft, context) => {
    const interpreted = interpretCellInput(draft.text, context.inputContext);
    return interpreted.valueType === 'number' || interpreted.valueType === 'empty'
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Date/time editor input is not a canonical date or time.', recovery: 'Enter a date/time accepted by the workbook culture and date system.' };
  },
};

const validationListBehavior: CellEditorBehavior = {
  ...plainBehavior('validation-list'),
  listItems: (context) => context.validationValues?.map((value) => ({ label: value, text: value })) ?? [],
  validate: (draft, context) => {
    const values = context.validationValues;
    if (!values || values.length === 0) {
      return {
        valid: false,
        code: 'CELL_EDIT_COMMIT_REJECTED',
        message: 'Validation list editor requires a resolved list rule.',
        recovery: 'Configure a DataValidationRule list source for the target cell.',
      };
    }
    return values.some((value) => value.localeCompare(draft.text, context.inputContext.cultureId === 'invariant' ? 'en-US' : context.inputContext.cultureId, { sensitivity: 'base' }) === 0)
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'The value is not present in the validation list.', recovery: 'Choose one of the canonical validation list values.' };
  },
};

const comboBoxBehavior: CellEditorBehavior = {
  ...plainBehavior('combo-box'),
  listItems: (context) => context.cell?.editor?.kind === 'combo-box'
    ? context.cell.editor.items.map((item) => ({ label: item.label ?? String(item.value ?? ''), text: String(item.value ?? '') }))
    : [],
  validate: (draft, context) => {
    if (context.cell?.editor?.kind !== 'combo-box') {
      return { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'ComboBox editor configuration is missing.', recovery: 'Apply a valid ComboBox editor configuration.' };
    }
    if (context.cell.editor.editable) return valid();
    const accepted = context.cell.editor.items.some((item) => String(item.value ?? '') === draft.text || item.label === draft.text);
    return accepted
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'The value is not present in the non-editable ComboBox.', recovery: 'Choose a configured ComboBox item.' };
  },
  toCommitPayload: (draft, context) => {
    if (context.cell?.editor?.kind !== 'combo-box') return rawText(draft);
    const item = context.cell.editor.items.find((candidate) => String(candidate.value ?? '') === draft.text || candidate.label === draft.text);
    return item ? { kind: 'typed-value', value: structuredClone(item.value) } : rawText(draft);
  },
};

function matchesMask(text: string, mask: string): boolean {
  const tokens = [...mask];
  const input = [...text];
  let sourceIndex = 0;
  let inputIndex = 0;
  while (sourceIndex < tokens.length) {
    const token = tokens[sourceIndex++]!;
    if (token === '\\') {
      const literal = tokens[sourceIndex++];
      if (literal === undefined || input[inputIndex++] !== literal) return false;
      continue;
    }
    const value = input[inputIndex++];
    if (value === undefined) return false;
    if (token === '#' && !/^[0-9]$/.test(value)) return false;
    if (token === 'A' && !/^\p{L}$/u.test(value)) return false;
    if (token === '*' && !/^[\p{L}\p{N}]$/u.test(value)) return false;
    if (token !== '#' && token !== 'A' && token !== '*' && value !== token) return false;
  }
  return inputIndex === input.length;
}

const maskBehavior: CellEditorBehavior = {
  ...plainBehavior('mask'),
  validate: (draft, context) => {
    if (context.cell?.editor?.kind !== 'mask') {
      return { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Mask editor configuration is missing.', recovery: 'Apply a valid mask editor configuration.' };
    }
    return matchesMask(draft.text, context.cell.editor.mask)
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: `The value does not match mask ${context.cell.editor.mask}.`, recovery: 'Enter a value matching the complete mask.' };
  },
};

const formulaBehavior: CellEditorBehavior = {
  ...plainBehavior('formula'),
  surface: { kind: 'text', inputMode: 'text', multiline: false, autoCapitalize: 'off' },
  validate: (draft, context): CellEditorValidationResult => {
    if (!draft.text.startsWith('=')) {
      return { valid: false, code: 'CELL_EDIT_INVALID_FORMULA', message: 'Formula editor input must start with =.', recovery: 'Enter a formula beginning with =.' };
    }
    try {
      parseFormula(draft.text);
      return valid();
    } catch (cause) {
      if (context.cell?.editor?.kind === 'formula' && context.cell.editor.allowInvalidFormula) return valid();
      return { valid: false, code: 'CELL_EDIT_INVALID_FORMULA', message: cause instanceof Error ? cause.message : 'Formula is invalid.', recovery: 'Correct the formula syntax before committing.' };
    }
  },
};

const checkboxBehavior: CellEditorBehavior = {
  kind: 'checkbox',
  surface: { kind: 'checkbox', inputMode: 'text', multiline: false },
  canEnter: (context) => context.source === 'double-click'
    ? { allowed: false, code: 'CELL_EDIT_UNSUPPORTED_TARGET', message: 'Checkbox double-click is owned by the cell control.', recovery: 'Use the checkbox glyph, Spacebar, or F2 for explicit value editing.' }
    : allowed(),
  createDraft: (context) => {
    if (context.initialText !== undefined) return { kind: 'plain', text: context.initialText };
    const config = context.cell?.editor?.kind === 'checkbox' ? context.cell.editor : undefined;
    if (!config) throw new Error('Checkbox behavior requires a checkbox editor configuration');
    const state = checkboxStateFromValue(config, context.cell?.value ?? null);
    if (!state) throw new Error('Checkbox cell value does not match its configured states');
    return { kind: 'plain', text: state === 'indeterminate' ? 'INDETERMINATE' : state === 'checked' ? 'TRUE' : 'FALSE' };
  },
  reduce: unchanged,
  ownsKey: () => true,
  validate: (draft, context) => /^(true|false)$/i.test(draft.text) || (context.cell?.editor?.kind === 'checkbox' && context.cell.editor.threeState && /^indeterminate$/i.test(draft.text))
    ? valid()
    : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Checkbox editor accepts TRUE, FALSE, or the configured indeterminate state.', recovery: 'Toggle the checkbox or enter a supported state.' },
  toCommitPayload: (draft, context) => {
    const config = context.cell?.editor?.kind === 'checkbox' ? context.cell.editor : undefined;
    if (!config) throw new Error('Checkbox behavior requires a checkbox editor configuration');
    const value = checkboxValueForState(config, /^indeterminate$/i.test(draft.text) ? 'indeterminate' : /^true$/i.test(draft.text) ? 'checked' : 'unchecked');
    return { kind: 'typed-value', value };
  },
  hitTestControl: (point, rect) => {
    const size = Math.min(14, Math.max(10, rect.height - 8));
    return point.x >= 2 && point.x <= 6 + size && point.y >= (rect.height - size) / 2 - 2 && point.y <= (rect.height + size) / 2 + 2 ? { kind: 'toggle' } : null;
  },
  controlActionForKey: (gesture) => gesture.key === ' ' && !gesture.ctrl && !gesture.meta && !gesture.alt ? { kind: 'toggle' } : null,
};

function richTextRuns(cell: CellData | null): RichTextRun[] {
  if (cell?.richText) return structuredClone(cell.richText);
  const text = cellText(cell);
  return text ? [{ text }] : [];
}

const richTextBehavior: CellEditorBehavior = {
  kind: 'rich-text',
  surface: { kind: 'rich-text', inputMode: 'text', multiline: true },
  canEnter: (context) => context.cell?.formula
    ? { allowed: false, code: 'CELL_EDIT_UNSUPPORTED_TARGET', message: 'Formula cells cannot enter rich-text editing.', recovery: 'Remove the formula before applying rich-text runs.' }
    : allowed(),
  createDraft: (context) => context.initialText === undefined
    ? { kind: 'rich-text', text: cellText(context.cell), runs: richTextRuns(context.cell) }
    : { kind: 'rich-text', text: context.initialText, runs: context.initialText ? [{ text: context.initialText }] : [] },
  reduce: unchanged,
  ownsKey: () => false,
  validate: (draft) => draft.kind === 'rich-text'
    ? valid()
    : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Rich-text editor requires a rich-text draft.', recovery: 'Re-enter the target through RichTextEditorBehavior.' },
  toCommitPayload: (draft) => draft.kind === 'rich-text'
    ? { kind: 'rich-text', text: draft.text, runs: structuredClone(draft.runs) }
    : { kind: 'rich-text', text: draft.text, runs: [{ text: draft.text }] },
};

export function createCellEditorRegistry(): CellEditorRegistry {
  const registry = new CellEditorRegistry();
  registry.register(plainBehavior('text'));
  registry.register(numberBehavior);
  registry.register(dateTimeBehavior);
  registry.register(validationListBehavior);
  registry.register(comboBoxBehavior);
  registry.register(checkboxBehavior);
  registry.register(maskBehavior);
  registry.register(formulaBehavior);
  registry.register(richTextBehavior);
  return registry;
}
