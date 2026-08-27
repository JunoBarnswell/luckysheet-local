export type CellEditorScalar = string | number | boolean | null;

interface CellEditorBaseConfig {
  placeholder?: string;
  /** Mirrors the workbook-owned allowEditInCell capability. */
  allowEditInCell?: boolean;
}

export interface TextCellEditorConfig extends CellEditorBaseConfig {
  kind: 'text';
  multiline?: boolean;
}

export interface NumberCellEditorConfig extends CellEditorBaseConfig {
  kind: 'number';
  min?: number;
  max?: number;
  step?: number;
}

export interface DateTimeCellEditorConfig extends CellEditorBaseConfig {
  kind: 'datetime';
  mode: 'date' | 'time' | 'datetime';
}

/** Values and formulas are owned by the canonical DataValidationRule. */
export interface ValidationListCellEditorConfig extends CellEditorBaseConfig {
  kind: 'validation-list';
}

export interface ComboBoxCellEditorItem {
  value: CellEditorScalar;
  label?: string;
}

export interface ComboBoxCellEditorConfig extends CellEditorBaseConfig {
  kind: 'combo-box';
  items: ComboBoxCellEditorItem[];
  editable: boolean;
}

export interface CheckboxCellEditorConfig extends CellEditorBaseConfig {
  kind: 'checkbox';
  trueValue?: CellEditorScalar;
  falseValue?: CellEditorScalar;
  indeterminateValue?: CellEditorScalar;
  threeState?: boolean;
}

export type CheckboxCellState = 'checked' | 'unchecked' | 'indeterminate';

function checkboxStateValues(editor: CheckboxCellEditorConfig): Record<CheckboxCellState, CellEditorScalar> {
  return {
    checked: editor.trueValue === undefined ? true : editor.trueValue,
    unchecked: editor.falseValue === undefined ? false : editor.falseValue,
    indeterminate: editor.indeterminateValue === undefined ? null : editor.indeterminateValue,
  };
}

export function isUnambiguousCheckboxEditor(editor: CheckboxCellEditorConfig): boolean {
  const values = checkboxStateValues(editor);
  if (Object.is(values.checked, values.unchecked)) return false;
  return !editor.threeState
    || (!Object.is(values.indeterminate, values.checked) && !Object.is(values.indeterminate, values.unchecked));
}

export function checkboxValueForState(editor: CheckboxCellEditorConfig, state: CheckboxCellState): CellEditorScalar {
  if (state === 'indeterminate' && !editor.threeState) throw new Error('Two-state checkbox has no indeterminate value');
  if (!isUnambiguousCheckboxEditor(editor)) throw new Error('Checkbox state values must be distinct');
  return checkboxStateValues(editor)[state];
}

export function checkboxStateFromValue(editor: CheckboxCellEditorConfig, value: CellEditorScalar): CheckboxCellState | null {
  if (!isUnambiguousCheckboxEditor(editor)) throw new Error('Checkbox state values must be distinct');
  const values = checkboxStateValues(editor);
  if (Object.is(value, values.checked)) return 'checked';
  if (Object.is(value, values.unchecked)) return 'unchecked';
  if (editor.threeState && Object.is(value, values.indeterminate)) return 'indeterminate';
  return null;
}

export function normalizeCheckboxValue(editor: CheckboxCellEditorConfig, value: CellEditorScalar): CellEditorScalar {
  const configuredState = checkboxStateFromValue(editor, value);
  if (configuredState) return checkboxValueForState(editor, configuredState);
  if (value === null) return checkboxValueForState(editor, editor.threeState ? 'indeterminate' : 'unchecked');
  if (typeof value === 'boolean') return checkboxValueForState(editor, value ? 'checked' : 'unchecked');
  if (typeof value === 'number' && (value === 0 || value === 1)) return checkboxValueForState(editor, value === 1 ? 'checked' : 'unchecked');
  if (typeof value === 'string') {
    const normalized = value.trim().toUpperCase();
    if (normalized === 'TRUE') return checkboxValueForState(editor, 'checked');
    if (normalized === 'FALSE') return checkboxValueForState(editor, 'unchecked');
    if (normalized === 'INDETERMINATE' && editor.threeState) return checkboxValueForState(editor, 'indeterminate');
  }
  throw new Error('Checkbox source value must match a configured state, blank, Boolean, 0/1, TRUE, FALSE, or INDETERMINATE');
}

export function nextCheckboxValue(editor: CheckboxCellEditorConfig, value: CellEditorScalar): CellEditorScalar {
  const state = checkboxStateFromValue(editor, value);
  if (!state) throw new Error('Checkbox cell value does not match its configured states');
  if (state === 'unchecked') return checkboxValueForState(editor, 'checked');
  if (state === 'checked') return checkboxValueForState(editor, editor.threeState ? 'indeterminate' : 'unchecked');
  return checkboxValueForState(editor, 'unchecked');
}

export interface MaskCellEditorConfig extends CellEditorBaseConfig {
  kind: 'mask';
  /** Canonical mask grammar, interpreted only by MaskEditorBehavior. */
  mask: string;
  promptCharacter?: string;
}

export interface FormulaCellEditorConfig extends CellEditorBaseConfig {
  kind: 'formula';
  allowInvalidFormula?: boolean;
}

export interface RichTextCellEditorConfig extends CellEditorBaseConfig {
  kind: 'rich-text';
}

export type CellEditorOptionValue = null | boolean | number | string | CellEditorOptionValue[] | { [key: string]: CellEditorOptionValue };

export interface CustomCellEditorConfig extends CellEditorBaseConfig {
  kind: 'custom';
  editorId: string;
  options?: { [key: string]: CellEditorOptionValue };
}

export type CellEditorConfig =
  | TextCellEditorConfig
  | NumberCellEditorConfig
  | DateTimeCellEditorConfig
  | ValidationListCellEditorConfig
  | ComboBoxCellEditorConfig
  | CheckboxCellEditorConfig
  | MaskCellEditorConfig
  | FormulaCellEditorConfig
  | RichTextCellEditorConfig
  | CustomCellEditorConfig;

export type CellEditorKind = CellEditorConfig['kind'];

export function isCellEditorConfig(value: unknown): value is CellEditorConfig {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
  const editor = value as Record<string, unknown>;
  if (editor.allowEditInCell !== undefined && typeof editor.allowEditInCell !== 'boolean') return false;
  if (editor.placeholder !== undefined && typeof editor.placeholder !== 'string') return false;
  switch (editor.kind) {
    case 'text':
      return editor.multiline === undefined || typeof editor.multiline === 'boolean';
    case 'number':
      return ['min', 'max', 'step'].every((key) => editor[key] === undefined || (typeof editor[key] === 'number' && Number.isFinite(editor[key])));
    case 'datetime':
      return editor.mode === 'date' || editor.mode === 'time' || editor.mode === 'datetime';
    case 'validation-list':
    case 'rich-text':
      return true;
    case 'formula':
      return editor.allowInvalidFormula === undefined || typeof editor.allowInvalidFormula === 'boolean';
    case 'custom':
      return typeof editor.editorId === 'string' && editor.editorId.trim().length > 0 && (editor.options === undefined || isEditorOptionValue(editor.options));
    case 'combo-box':
      return typeof editor.editable === 'boolean' && Array.isArray(editor.items) && editor.items.every((item) => {
        if (!item || typeof item !== 'object' || Array.isArray(item)) return false;
        const candidate = item as Record<string, unknown>;
        const scalar = candidate.value === null || ['string', 'number', 'boolean'].includes(typeof candidate.value);
        return scalar && (candidate.label === undefined || typeof candidate.label === 'string');
      });
    case 'checkbox':
      if (!['trueValue', 'falseValue', 'indeterminateValue'].every((key) => editor[key] === undefined || editor[key] === null || ['string', 'number', 'boolean'].includes(typeof editor[key]))
        || (editor.threeState !== undefined && typeof editor.threeState !== 'boolean')) return false;
      return isUnambiguousCheckboxEditor(editor as unknown as CheckboxCellEditorConfig);
    case 'mask':
      return typeof editor.mask === 'string' && editor.mask.length > 0
        && (editor.promptCharacter === undefined || (typeof editor.promptCharacter === 'string' && [...editor.promptCharacter].length === 1));
    default:
      return false;
  }
}

function isEditorOptionValue(value: unknown): value is CellEditorOptionValue {
  if (value === null || ['boolean', 'number', 'string'].includes(typeof value)) return typeof value !== 'number' || Number.isFinite(value);
  if (Array.isArray(value)) return value.every(isEditorOptionValue);
  return Boolean(value && typeof value === 'object' && Object.values(value as Record<string, unknown>).every(isEditorOptionValue));
}
