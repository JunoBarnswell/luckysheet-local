import React, { useState } from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ConditionalFormatOperator, ConditionalFormatRule, ConditionalFormatType, RangeRef } from '@react-sheets/core-model';

export interface ConditionalFormatPanelProps {
  sheetId: string;
  rules: ConditionalFormatRule[];
  onAddRule: (rule: ConditionalFormatRule) => void;
  onRemoveRule: (id: string) => void;
  onClose?: () => void;
}

export function ConditionalFormatPanel({
  sheetId,
  rules,
  onAddRule,
  onRemoveRule,
  onClose,
}: ConditionalFormatPanelProps) {
  const [type, setType] = useState<ConditionalFormatType>('highlight');
  const [operator, setOperator] = useState<ConditionalFormatOperator>('greaterThan');
  const [value1, setValue1] = useState('50');
  const [bg, setBg] = useState('#dcfce7');
  const [color, setColor] = useState('#166534');

  const handleCreate = () => {
    const newRule: ConditionalFormatRule = {
      id: 'cf-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      ranges: [
        {
          sheetId,
          startRow: 0,
          endRow: 20,
          startColumn: 0,
          endColumn: 10,
        },
      ],
      type,
      operator,
      value1,
      style: {
        background: bg,
        textColor: color,
        bold: true,
      },
    };
    onAddRule(newRule);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Conditional Formatting</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Format Type
            </Text>
            <Select
              value={type}
              onChange={(e) => setType(e.target.value as ConditionalFormatType)}
              sizeVariant="sm"
            >
              <option value="highlight">Highlight Cell Rules</option>
              <option value="colorScale">Color Scale</option>
              <option value="dataBar">Data Bars</option>
            </Select>
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Rule Condition
            </Text>
            <Select
              value={operator}
              onChange={(e) => setOperator(e.target.value as ConditionalFormatOperator)}
              sizeVariant="sm"
            >
              <option value="greaterThan">Greater Than (&gt;)</option>
              <option value="lessThan">Less Than (&lt;)</option>
              <option value="equal">Equal To (=)</option>
              <option value="between">Between Range</option>
              <option value="containsText">Text Contains</option>
              <option value="duplicate">Duplicate Values</option>
            </Select>
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Threshold Value
            </Text>
            <TextInput
              value={value1}
              onChange={(e) => setValue1(e.target.value)}
              placeholder="e.g. 100, Target, At risk"
            />
          </div>

          <div className="grid grid-cols-2 gap-2">
            <div>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                Cell Fill
              </Text>
              <div className="flex items-center gap-2">
                <input
                  type="color"
                  value={bg}
                  onChange={(e) => setBg(e.target.value)}
                  className="h-8 w-8 cursor-pointer rounded border border-slate-300 p-0"
                />
                <TextInput value={bg} onChange={(e) => setBg(e.target.value)} className="h-8 text-xs font-mono" />
              </div>
            </div>
            <div>
              <Text size="xs" weight="medium" className="mb-1 text-slate-700">
                Text Color
              </Text>
              <div className="flex items-center gap-2">
                <input
                  type="color"
                  value={color}
                  onChange={(e) => setColor(e.target.value)}
                  className="h-8 w-8 cursor-pointer rounded border border-slate-300 p-0"
                />
                <TextInput value={color} onChange={(e) => setColor(e.target.value)} className="h-8 text-xs font-mono" />
              </div>
            </div>
          </div>

          <Button variant="primary" size="sm" icon="plus" onClick={handleCreate}>
            Apply Formatting Rule
          </Button>

          {rules.length > 0 ? (
            <div className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Active Rules ({rules.length})
              </Text>
              <Stack gap="xs">
                {rules.map((r) => (
                  <div
                    key={r.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <div>
                      <div className="font-medium text-slate-800">
                        {r.operator} {r.value1}
                      </div>
                      <div className="text-[10px] text-slate-500">{r.type}</div>
                    </div>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onRemoveRule(r.id)}
                      className="text-rose-600 hover:bg-rose-50"
                    />
                  </div>
                ))}
              </Stack>
            </div>
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
