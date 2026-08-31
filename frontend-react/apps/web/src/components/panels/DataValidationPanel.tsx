import React, { useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { parseAddress, type DataValidationOperator, type DataValidationRule, type DataValidationType, type RangeRef } from '@react-sheets/core-model';

export interface DataValidationPanelProps {
  sheetId: string;
  range?: RangeRef;
  rules: DataValidationRule[];
  onAddRule: (rule: DataValidationRule) => void;
  onRemoveRule: (id: string) => void;
  onClose?: () => void;
}

export interface DataValidationRuleDraft {
  id: string;
  sheetId: string;
  range?: RangeRef;
  type: DataValidationType;
  operator: DataValidationOperator;
  formula1: string;
  formula2: string;
  listSourceKind?: 'values' | 'range' | 'formula';
  errorMessage: string;
}

/** Selection-aware canonical rule builder shared by UI tests and the panel. */
export function buildDataValidationRule(draft: DataValidationRuleDraft): DataValidationRule {
  const { id, sheetId, range, type, operator, errorMessage } = draft;
  if (!range) throw new Error('Select a worksheet range before creating a validation rule.');
  if (range.sheetId !== sheetId) throw new Error('The selected range belongs to another worksheet.');
  const first = draft.formula1.trim();
  const second = draft.formula2.trim();
  const requiresBounds = type === 'whole' || type === 'decimal' || type === 'date' || type === 'time' || type === 'textLength';
  if ((requiresBounds || type === 'custom') && !first) {
    throw new Error(type === 'custom' ? 'A custom validation formula is required.' : 'The first validation bound is required.');
  }
  if (requiresBounds && (operator === 'between' || operator === 'notBetween') && !second) {
    throw new Error('The second validation bound is required for a between operator.');
  }
  const listSourceKind = draft.listSourceKind ?? 'values';
  const values = type === 'list' && listSourceKind === 'values' ? first.split(',').map((value) => value.trim()).filter(Boolean) : [];
  if (type === 'list' && !first) throw new Error('A list source is required.');
  if (type === 'list' && listSourceKind === 'values' && values.length === 0) throw new Error('At least one list value is required.');
  let listSource: DataValidationRule['listSource'];
  if (type === 'list' && listSourceKind === 'values') listSource = { kind: 'values', values };
  if (type === 'list' && listSourceKind === 'formula') listSource = { kind: 'formula', formula: first.startsWith('=') ? first : `=${first}` };
  if (type === 'list' && listSourceKind === 'range') {
    const [startText, endText = startText] = first.replace(/^=/, '').replace(/\$/g, '').split(':');
    const start = parseAddress(startText ?? '');
    const end = parseAddress(endText ?? '');
    if (!start || !end) throw new Error('The list range must be a valid A1 range on the active worksheet.');
    listSource = { kind: 'range', range: { sheetId, startRow: Math.min(start.row, end.row), endRow: Math.max(start.row, end.row), startColumn: Math.min(start.column, end.column), endColumn: Math.max(start.column, end.column) } };
  }
  return {
    id,
    sheetId,
    ranges: [{ ...range }],
    formulaAnchor: { sheetId, row: range.startRow, column: range.startColumn },
    type,
    ...(requiresBounds ? {
      operator,
      formula1: first,
      ...((operator === 'between' || operator === 'notBetween') ? { formula2: second } : {}),
    } : {}),
    ...(type === 'custom' ? { formula1: first.startsWith('=') ? first : `=${first}` } : {}),
    ...(type === 'list' ? { listSource, showDropdown: true } : {}),
    ...(type === 'checkbox' ? { showDropdown: true } : {}),
    allowBlank: true,
    alertStyle: 'stop',
    showErrorMessage: true,
    errorMessage,
  };
}

export function DataValidationPanel({
  sheetId,
  range,
  rules,
  onAddRule,
  onRemoveRule,
  onClose,
}: DataValidationPanelProps) {
  const [type, setType] = useState<DataValidationType>('list');
  const [operator, setOperator] = useState<DataValidationOperator>('between');
  const [listSourceKind, setListSourceKind] = useState<'values' | 'range' | 'formula'>('values');
  const [formula1, setFormula1] = useState('On track, Needs review, At risk, Completed');
  const [formula2, setFormula2] = useState('');
  const [errorMessage, setErrorMessage] = useState('Please select a valid option from the list');
  const [createError, setCreateError] = useState<string | null>(null);

  const handleCreate = () => {
    setCreateError(null);
    if (!globalThis.crypto?.randomUUID) {
      setCreateError('The browser cannot allocate a canonical validation identity.');
      return;
    }
    try {
      onAddRule(buildDataValidationRule({
        id: `dv-${globalThis.crypto.randomUUID()}`,
        sheetId,
        range,
        type,
        operator,
        formula1,
        formula2,
        listSourceKind,
        errorMessage,
      }));
    } catch (error) {
      setCreateError(error instanceof Error ? error.message : 'The validation rule is invalid.');
    }
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Data Validation</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Validation Criteria
            </Text>
            <Select
              value={type}
              onChange={(e) => {
                const nextType = e.target.value as DataValidationType;
                setType(nextType);
                setFormula1(nextType === 'list' ? 'On track, Needs review, At risk, Completed' : nextType === 'custom' ? '=A1<>""' : '0');
                setFormula2(nextType === 'whole' || nextType === 'decimal' || nextType === 'date' || nextType === 'time' || nextType === 'textLength' ? '100' : '');
              }}
              sizeVariant="sm"
              options={[
                { value: 'list', label: 'Dropdown List' },
                { value: 'whole', label: 'Whole Number' },
                { value: 'decimal', label: 'Decimal Number' },
                { value: 'date', label: 'Date' },
                { value: 'time', label: 'Time' },
                { value: 'checkbox', label: 'Checkbox' },
                { value: 'textLength', label: 'Text Length' },
                { value: 'custom', label: 'Custom Formula' },
              ]}
            />
          </Box>

          {(type === 'whole' || type === 'decimal' || type === 'date' || type === 'time' || type === 'textLength') ? (
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">Operator</Text>
              <Select
                value={operator}
                onChange={(event) => setOperator(event.target.value as DataValidationOperator)}
                sizeVariant="sm"
                options={[
                  { value: 'between', label: 'Between' },
                  { value: 'notBetween', label: 'Not between' },
                  { value: 'equal', label: 'Equal to' },
                  { value: 'notEqual', label: 'Not equal to' },
                  { value: 'greaterThan', label: 'Greater than' },
                  { value: 'lessThan', label: 'Less than' },
                ]}
              />
            </Box>
          ) : null}

          {type === 'list' ? (
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">List source type</Text>
              <Select
                value={listSourceKind}
                onChange={(event) => setListSourceKind(event.target.value as 'values' | 'range' | 'formula')}
                sizeVariant="sm"
                options={[
                  { value: 'values', label: 'Values' },
                  { value: 'range', label: 'Worksheet range' },
                  { value: 'formula', label: 'Formula' },
                ]}
              />
            </Box>
          ) : null}

          {type !== 'checkbox' ? (
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              {type === 'list'
                ? listSourceKind === 'values' ? 'Allowed values (comma separated)' : listSourceKind === 'range' ? 'Worksheet range' : 'List formula'
                : type === 'custom' ? 'Validation formula' : 'First bound'}
            </Text>
            <TextInput
              value={formula1}
              onChange={(e) => setFormula1(e.target.value)}
              placeholder={type === 'list'
                ? listSourceKind === 'values' ? 'e.g. Option A, Option B' : listSourceKind === 'range' ? 'A1:A10' : '=UNIQUE(A1:A10)'
                : type === 'custom' ? '=A1<>""' : '0'}
            />
          </Box>
          ) : null}

          {(type === 'whole' || type === 'decimal' || type === 'date' || type === 'time' || type === 'textLength')
            && (operator === 'between' || operator === 'notBetween') ? (
            <Box>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">Second bound</Text>
              <TextInput value={formula2} onChange={(event) => setFormula2(event.target.value)} placeholder="100" />
            </Box>
          ) : null}

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Error Message Alert
            </Text>
            <TextInput
              value={errorMessage}
              onChange={(e) => setErrorMessage(e.target.value)}
              placeholder="e.g. Invalid input"
            />
          </Box>

          {createError ? <Text size="xs" tone="danger">{createError}</Text> : null}

          <Button variant="primary" size="sm" icon="check-circle" onClick={handleCreate} disabled={!range}>
            Add Validation Rule
          </Button>

          {rules.length > 0 ? (
            <Box className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Active Validations ({rules.length})
              </Text>
              <Stack gap="xs">
                {rules.map((r) => (
                  <Box
                    key={r.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <Stack gap="none">
                      <Text size="sm" weight="medium" className="text-slate-800">{r.type.toUpperCase()} Rule</Text>
                      <Text size="xs" tone="subtle">
                        {r.ranges.map((item) => `${item.startRow + 1}:${item.startColumn + 1}-${item.endRow + 1}:${item.endColumn + 1}`).join(', ')}
                      </Text>
                    </Stack>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onRemoveRule(r.id)}
                      className="text-rose-600 hover:bg-rose-50"
                    />
                  </Box>
                ))}
              </Stack>
            </Box>
          ) : null}
        </Stack>
      </PanelBody>

      {onClose ? (
        <PanelFooter className="border-t border-slate-200 px-4 py-2">
          <Button variant="ghost" size="sm" onClick={onClose}>
            Close Panel
          </Button>
        </PanelFooter>
      ) : null}
    </Panel>
  );
}
