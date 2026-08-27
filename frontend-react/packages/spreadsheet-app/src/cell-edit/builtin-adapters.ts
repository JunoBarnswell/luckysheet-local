import type { CellData, RichTextRun } from '@react-sheets/core-model';
import { parseFormula } from '@react-sheets/formula-engine';
import { interpretCellInput } from '@react-sheets/sheet-features';
import type {
  CellEditCommitPayload,
  CellEditorAdapter,
  CellEditorContext,
  CellEditorValidationResult,
} from './adapter-registry';
import { CellEditorAdapterRegistry } from './adapter-registry';
import type { CellEditAdapterKind, CellEditDraft, CellEditIntent } from './contracts';

function cellText(cell: CellData | null): string {
  return cell?.formula ?? (cell?.value == null ? '' : String(cell.value));
}

function plainDraft(context: CellEditorContext): CellEditDraft {
  return { kind: 'plain', text: cellText(context.cell) };
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

function plainAdapter(kind: CellEditAdapterKind): CellEditorAdapter {
  return {
    kind,
    surface: { kind: kind === 'rich-text' ? 'rich-text' : kind === 'validation-list' || kind === 'combo-box' ? 'list' : kind === 'checkbox' ? 'checkbox' : 'text', inputMode: kind === 'number' ? 'decimal' : kind === 'datetime' ? 'numeric' : 'text', multiline: kind === 'text' || kind === 'rich-text' },
    canEnter: allowed,
    createDraft: plainDraft,
    reduce: unchanged,
    ownsKey: () => false,
    validate: valid,
    toCommitPayload: rawText,
  };
}

const numberAdapter: CellEditorAdapter = {
  ...plainAdapter('number'),
  validate: (draft, context) => {
    const interpreted = interpretCellInput(draft.text, context.inputContext);
    return interpreted.valueType === 'number' || interpreted.valueType === 'empty'
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Number editor input is not a canonical number.', recovery: 'Enter a number accepted by the workbook culture and number format.' };
  },
};

const dateTimeAdapter: CellEditorAdapter = {
  ...plainAdapter('datetime'),
  validate: (draft, context) => {
    const interpreted = interpretCellInput(draft.text, context.inputContext);
    return interpreted.valueType === 'number' || interpreted.valueType === 'empty'
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Date/time editor input is not a canonical date or time.', recovery: 'Enter a date/time accepted by the workbook culture and date system.' };
  },
};

const validationListAdapter: CellEditorAdapter = {
  ...plainAdapter('validation-list'),
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

const comboBoxAdapter: CellEditorAdapter = {
  ...plainAdapter('combo-box'),
  validate: (draft, context) => {
    if (context.config?.kind !== 'combo-box') {
      return { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'ComboBox editor configuration is missing.', recovery: 'Apply a valid ComboBox editor configuration.' };
    }
    if (context.config.editable) return valid();
    const accepted = context.config.items.some((item) => String(item.value ?? '') === draft.text || item.label === draft.text);
    return accepted
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'The value is not present in the non-editable ComboBox.', recovery: 'Choose a configured ComboBox item.' };
  },
  toCommitPayload: (draft, context) => {
    if (context.config?.kind !== 'combo-box') return rawText(draft);
    const item = context.config.items.find((candidate) => String(candidate.value ?? '') === draft.text || candidate.label === draft.text);
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

const maskAdapter: CellEditorAdapter = {
  ...plainAdapter('mask'),
  validate: (draft, context) => {
    if (context.config?.kind !== 'mask') {
      return { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Mask editor configuration is missing.', recovery: 'Apply a valid mask editor configuration.' };
    }
    return matchesMask(draft.text, context.config.mask)
      ? valid()
      : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: `The value does not match mask ${context.config.mask}.`, recovery: 'Enter a value matching the complete mask.' };
  },
};

const formulaAdapter: CellEditorAdapter = {
  ...plainAdapter('formula'),
  validate: (draft, context): CellEditorValidationResult => {
    if (!draft.text.startsWith('=')) {
      return { valid: false, code: 'CELL_EDIT_INVALID_FORMULA', message: 'Formula editor input must start with =.', recovery: 'Enter a formula beginning with =.' };
    }
    try {
      parseFormula(draft.text);
      return valid();
    } catch (cause) {
      if (context.config?.kind === 'formula' && context.config.allowInvalidFormula) return valid();
      return { valid: false, code: 'CELL_EDIT_INVALID_FORMULA', message: cause instanceof Error ? cause.message : 'Formula is invalid.', recovery: 'Correct the formula syntax before committing.' };
    }
  },
};

const checkboxAdapter: CellEditorAdapter = {
  kind: 'checkbox',
  surface: { kind: 'checkbox', inputMode: 'text', multiline: false },
  canEnter: allowed,
  createDraft: (context) => {
    const config = context.config?.kind === 'checkbox' ? context.config : undefined;
    const value = context.cell?.value ?? null;
    if (config?.threeState && config.indeterminateValue !== undefined && Object.is(value, config.indeterminateValue)) return { kind: 'plain', text: 'INDETERMINATE' };
    const checked = config?.trueValue !== undefined ? Object.is(value, config.trueValue) : value === true;
    return { kind: 'plain', text: checked ? 'TRUE' : 'FALSE' };
  },
  reduce: unchanged,
  ownsKey: () => true,
  validate: (draft, context) => /^(true|false)$/i.test(draft.text) || (context.config?.kind === 'checkbox' && context.config.threeState && /^indeterminate$/i.test(draft.text))
    ? valid()
    : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Checkbox editor accepts TRUE, FALSE, or the configured indeterminate state.', recovery: 'Toggle the checkbox or enter a supported state.' },
  toCommitPayload: (draft, context) => {
    const config = context.config?.kind === 'checkbox' ? context.config : undefined;
    const value = /^indeterminate$/i.test(draft.text)
      ? config?.indeterminateValue ?? null
      : /^true$/i.test(draft.text)
        ? config?.trueValue ?? true
        : config?.falseValue ?? false;
    return { kind: 'typed-value', value };
  },
};

function richTextRuns(cell: CellData | null): RichTextRun[] {
  if (cell?.richText) return structuredClone(cell.richText);
  const text = cellText(cell);
  return text ? [{ text }] : [];
}

const richTextAdapter: CellEditorAdapter = {
  kind: 'rich-text',
  surface: { kind: 'rich-text', inputMode: 'text', multiline: true },
  canEnter: (context) => context.cell?.formula
    ? { allowed: false, code: 'CELL_EDIT_UNSUPPORTED_TARGET', message: 'Formula cells cannot enter rich-text editing.', recovery: 'Remove the formula before applying rich-text runs.' }
    : allowed(),
  createDraft: (context) => ({ kind: 'rich-text', text: cellText(context.cell), runs: richTextRuns(context.cell) }),
  reduce: unchanged,
  ownsKey: () => false,
  validate: (draft) => draft.kind === 'rich-text'
    ? valid()
    : { valid: false, code: 'CELL_EDIT_COMMIT_REJECTED', message: 'Rich-text editor requires a rich-text draft.', recovery: 'Re-enter the target through RichTextEditorAdapter.' },
  toCommitPayload: (draft) => draft.kind === 'rich-text'
    ? { kind: 'rich-text', text: draft.text, runs: structuredClone(draft.runs) }
    : { kind: 'rich-text', text: draft.text, runs: [{ text: draft.text }] },
};

export function createCellEditorAdapterRegistry(): CellEditorAdapterRegistry {
  const registry = new CellEditorAdapterRegistry();
  registry.register(plainAdapter('text'));
  registry.register(numberAdapter);
  registry.register(dateTimeAdapter);
  registry.register(validationListAdapter);
  registry.register(comboBoxAdapter);
  registry.register(checkboxAdapter);
  registry.register(maskAdapter);
  registry.register(formulaAdapter);
  registry.register(richTextAdapter);
  return registry;
}
