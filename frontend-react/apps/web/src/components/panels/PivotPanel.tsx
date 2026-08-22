import React, { useState } from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { PivotModel } from '@react-sheets/core-model';

export interface PivotPanelProps {
  sheetId: string;
  pivots: PivotModel[];
  onAddPivot: (pivot: PivotModel) => void;
  onRemovePivot: (id: string) => void;
  onClose?: () => void;
}

export function PivotPanel({
  sheetId,
  pivots,
  onAddPivot,
  onRemovePivot,
  onClose,
}: PivotPanelProps) {
  const [sourceRange, setSourceRange] = useState('A1:E10');
  const [rowField, setRowField] = useState('Status');
  const [valField, setValField] = useState('Target');

  const handleCreate = () => {
    const newPivot: PivotModel = {
      id: 'pivot-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      sourceRange: {
        sheetId,
        startRow: 0,
        endRow: 5,
        startColumn: 0,
        endColumn: 4,
      },
      rowFields: [rowField],
      columnFields: [],
      valueFields: [
        {
          field: valField,
          summarizeBy: 'sum',
        },
      ],
      filterFields: [],
    };
    onAddPivot(newPivot);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Pivot Table Builder</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Source Table Range
            </Text>
            <TextInput
              value={sourceRange}
              onChange={(e) => setSourceRange(e.target.value)}
              placeholder="e.g. A1:F20"
            />
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Row Group Field
            </Text>
            <TextInput
              value={rowField}
              onChange={(e) => setRowField(e.target.value)}
              placeholder="e.g. Category, Status, Region"
            />
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Values Field (SUM)
            </Text>
            <TextInput
              value={valField}
              onChange={(e) => setValField(e.target.value)}
              placeholder="e.g. Revenue, Actual, Amount"
            />
          </div>

          <Button variant="primary" size="sm" icon="table-pivot" onClick={handleCreate}>
            Generate Pivot Table
          </Button>

          {pivots.length > 0 ? (
            <div className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Active Pivots ({pivots.length})
              </Text>
              <Stack gap="xs">
                {pivots.map((p) => (
                  <div
                    key={p.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <div>
                      <div className="font-medium text-slate-800">
                        {p.rowFields.join(', ')} → {p.valueFields.map((v) => v.field).join(', ')}
                      </div>
                      <div className="text-[10px] text-slate-500">Pivot Table {p.id}</div>
                    </div>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onRemovePivot(p.id)}
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
