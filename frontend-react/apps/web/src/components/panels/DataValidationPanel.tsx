import React, { useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { DataValidationRule, DataValidationType } from '@react-sheets/core-model';

export interface DataValidationPanelProps {
  sheetId: string;
  rules: DataValidationRule[];
  onAddRule: (rule: DataValidationRule) => void;
  onRemoveRule: (id: string) => void;
  onClose?: () => void;
}

export function DataValidationPanel({
  sheetId,
  rules,
  onAddRule,
  onRemoveRule,
  onClose,
}: DataValidationPanelProps) {
  const [type, setType] = useState<DataValidationType>('list');
  const [formula1, setFormula1] = useState('On track, Needs review, At risk, Completed');
  const [errorMessage, setErrorMessage] = useState('Please select a valid option from the list');

  const handleCreate = () => {
    const newRule: DataValidationRule = {
      id: 'dv-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      ranges: [
        {
          sheetId,
          startRow: 1,
          endRow: 20,
          startColumn: 2,
          endColumn: 2,
        },
      ],
      type,
      formula1,
      showDropdown: true,
      errorMessage,
    };
    onAddRule(newRule);
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
              onChange={(e) => setType(e.target.value as DataValidationType)}
              sizeVariant="sm"
            >
              <option value="list">Dropdown List (Comma separated)</option>
              <option value="whole">Whole Number</option>
              <option value="decimal">Decimal Number</option>
              <option value="date">Date</option>
              <option value="textLength">Text Length</option>
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Allowed Values / Source
            </Text>
            <TextInput
              value={formula1}
              onChange={(e) => setFormula1(e.target.value)}
              placeholder="e.g. Option A, Option B, Option C"
            />
          </Box>

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

          <Button variant="primary" size="sm" icon="check-circle" onClick={handleCreate}>
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
                      <Text size="xs" tone="subtle">{r.formula1}</Text>
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
