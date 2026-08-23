import { Button, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { useState } from 'react';
import type { PivotCalculatedField, PivotCalculatedItem, PivotFieldDefinition } from '@react-sheets/core-model';

export interface PivotCalculatedEditorProps {
  fields: readonly PivotFieldDefinition[];
  calculatedFields: readonly PivotCalculatedField[];
  calculatedItems: readonly PivotCalculatedItem[];
  disabled?: boolean;
  onFieldsChange: (fields: PivotCalculatedField[]) => void;
  onItemsChange: (items: PivotCalculatedItem[]) => void;
}

export function PivotCalculatedEditor({ calculatedFields, calculatedItems, disabled = false, fields, onFieldsChange, onItemsChange }: PivotCalculatedEditorProps) {
  const [fieldName, setFieldName] = useState('');
  const [fieldFormula, setFieldFormula] = useState('');
  const [itemField, setItemField] = useState(fields[0]?.id ?? '');
  const [itemName, setItemName] = useState('');
  const [itemFormula, setItemFormula] = useState('');
  return (
    <Stack gap="sm" className="border-t border-line/80 pt-3">
      <Text size="xs" weight="semibold" tone="muted">CALCULATED FIELDS / ITEMS</Text>
      <Stack gap="xs" className="rounded-lg border border-slate-200 p-2">
        <Text size="xs" weight="semibold">Calculated field</Text>
        <Inline gap="xs"><TextInput aria-label="Calculated field name" disabled={disabled} placeholder="Name" value={fieldName} onChange={(event) => setFieldName(event.target.value)} /><TextInput aria-label="Calculated field formula" disabled={disabled} placeholder="=Amount*1.15" value={fieldFormula} onChange={(event) => setFieldFormula(event.target.value)} /><Button disabled={disabled || !fieldName.trim() || !fieldFormula.trim()} size="xs" variant="primary" onClick={() => { onFieldsChange([...calculatedFields, { name: fieldName.trim(), formula: fieldFormula.trim() }]); setFieldName(''); setFieldFormula(''); }}>Add</Button></Inline>
        {calculatedFields.map((field) => <Inline key={field.name} gap="xs" className="items-center"><Text size="xs" className="min-w-0 flex-1 truncate">{field.name} · {field.formula}</Text><Button disabled={disabled} size="xs" variant="ghost" onClick={() => onFieldsChange(calculatedFields.filter((item) => item.name !== field.name))}>Remove</Button></Inline>)}
      </Stack>
      <Stack gap="xs" className="rounded-lg border border-slate-200 p-2">
        <Text size="xs" weight="semibold">Calculated item</Text>
        <Inline gap="xs"><Select aria-label="Calculated item field" disabled={disabled || fields.length === 0} sizeVariant="sm" value={itemField} onChange={(event) => setItemField(event.target.value)}>{fields.map((field) => <option key={field.id} value={field.id}>{field.name}</option>)}</Select><TextInput aria-label="Calculated item name" disabled={disabled} placeholder="Item name" value={itemName} onChange={(event) => setItemName(event.target.value)} /><TextInput aria-label="Calculated item formula" disabled={disabled} placeholder="=Amount*3" value={itemFormula} onChange={(event) => setItemFormula(event.target.value)} /><Button disabled={disabled || !itemField || !itemName.trim() || !itemFormula.trim()} size="xs" variant="primary" onClick={() => { onItemsChange([...calculatedItems, { field: itemField, name: itemName.trim(), formula: itemFormula.trim() }]); setItemName(''); setItemFormula(''); }}>Add</Button></Inline>
        {calculatedItems.map((item) => <Inline key={`${item.field}-${item.name}`} gap="xs" className="items-center"><Text size="xs" className="min-w-0 flex-1 truncate">{item.field}:{item.name} · {item.formula}</Text><Button disabled={disabled} size="xs" variant="ghost" onClick={() => onItemsChange(calculatedItems.filter((entry) => entry !== item))}>Remove</Button></Inline>)}
      </Stack>
    </Stack>
  );
}
