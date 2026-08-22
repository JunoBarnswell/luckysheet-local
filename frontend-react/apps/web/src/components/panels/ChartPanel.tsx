import React, { useState } from 'react';
import { Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ChartModel } from '@react-sheets/core-model';

export interface ChartPanelProps {
  sheetId: string;
  charts: ChartModel[];
  activeChart?: ChartModel;
  /** 当前选区的 A1 文本,作为默认数据范围 */
  defaultRange?: string;
  onAddChart: (chart: ChartModel) => void;
  onRemoveChart: (id: string) => void;
  onClose?: () => void;
}

export function ChartPanel({
  sheetId,
  charts,
  activeChart,
  defaultRange,
  onAddChart,
  onRemoveChart,
  onClose,
}: ChartPanelProps) {
  const [type, setType] = useState<ChartModel['type']>(activeChart?.type ?? 'column');
  const [title, setTitle] = useState(activeChart?.title ?? 'Sales Overview');
  const [rangeInput, setRangeInput] = useState(defaultRange ?? 'A1:C5');

  const handleCreate = () => {
    const newChart: ChartModel = {
      id: 'chart-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      type,
      title,
      sourceRanges: [
        {
          sheetId,
          startRow: 0,
          endRow: 4,
          startColumn: 0,
          endColumn: 2,
        },
      ],
      bounds: {
        x: 100,
        y: 100,
        width: 480,
        height: 280,
      },
    };
    onAddChart(newChart);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Chart Settings</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Chart Title
            </Text>
            <TextInput
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="e.g. Monthly Revenue"
            />
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Chart Type
            </Text>
            <Select
              value={type}
              onChange={(e) => setType(e.target.value as ChartModel['type'])}
              sizeVariant="sm"
            >
              <option value="column">Column Chart</option>
              <option value="bar">Bar Chart</option>
              <option value="line">Line Chart</option>
              <option value="area">Area Chart</option>
              <option value="pie">Pie Chart</option>
              <option value="doughnut">Doughnut Chart</option>
              <option value="scatter">Scatter Plot</option>
            </Select>
          </div>

          <div>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Data Source Range
            </Text>
            <TextInput
              value={rangeInput}
              onChange={(e) => setRangeInput(e.target.value)}
              placeholder="e.g. A1:F6"
            />
          </div>

          <Button variant="primary" size="sm" icon="plus" onClick={handleCreate}>
            Insert Chart to Canvas
          </Button>

          {charts.length > 0 ? (
            <div className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Worksheet Charts ({charts.length})
              </Text>
              <Stack gap="xs">
                {charts.map((c) => (
                  <div
                    key={c.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <div>
                      <div className="font-medium text-slate-800">{c.title || c.type}</div>
                      <div className="text-[10px] text-slate-500">{c.type.toUpperCase()} Chart</div>
                    </div>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onRemoveChart(c.id)}
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
