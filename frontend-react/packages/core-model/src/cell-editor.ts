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

export interface MaskCellEditorConfig extends CellEditorBaseConfig {
  kind: 'mask';
  /** Canonical mask grammar, interpreted only by MaskEditorAdapter. */
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
  adapterId: string;
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
      return typeof editor.adapterId === 'string' && editor.adapterId.trim().length > 0 && (editor.options === undefined || isEditorOptionValue(editor.options));
    case 'combo-box':
      return typeof editor.editable === 'boolean' && Array.isArray(editor.items) && editor.items.every((item) => {
        if (!item || typeof item !== 'object' || Array.isArray(item)) return false;
        const candidate = item as Record<string, unknown>;
        const scalar = candidate.value === null || ['string', 'number', 'boolean'].includes(typeof candidate.value);
        return scalar && (candidate.label === undefined || typeof candidate.label === 'string');
      });
    case 'checkbox':
      return ['trueValue', 'falseValue', 'indeterminateValue'].every((key) => editor[key] === undefined || editor[key] === null || ['string', 'number', 'boolean'].includes(typeof editor[key]))
        && (editor.threeState === undefined || typeof editor.threeState === 'boolean');
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
