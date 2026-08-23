import React, { useState } from 'react';
import { Box, Button, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { ChartDrawingPayload, DrawingObject, DrawingPayload } from '@react-sheets/core-model';
import { parseRangeInput } from '../../domain/range-input';
import type { CommandDescriptor } from '@react-sheets/command-runtime';

export interface ChartPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  /** 当前选区的 A1 文本,作为默认数据范围 */
  defaultRange?: string;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

export function ChartPanel({
  sheetId,
  drawings,
  drawingPayloads,
  defaultRange,
  onCommand,
  onClose,
}: ChartPanelProps) {
  const [type, setType] = useState<ChartDrawingPayload['chartType']>('column');
  const [stacked, setStacked] = useState<NonNullable<ChartDrawingPayload['stacked']>>('none');
  const [title, setTitle] = useState('Sales Overview');
  const [rangeInput, setRangeInput] = useState(defaultRange ?? 'A1:C5');
  const sourceRange = parseRangeInput(rangeInput, sheetId);
  const chartEntries = drawings
    .filter((drawing) => drawing.kind === 'chart')
    .map((drawing) => ({ drawing, payload: drawingPayloads.get(drawing.payloadId) }))
    .filter((entry): entry is { drawing: DrawingObject; payload: Extract<DrawingPayload, { kind: 'chart' }> } => entry.payload?.kind === 'chart');

  const createId = (prefix: string): string => {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return `${prefix}-${crypto.randomUUID()}`;
    return `${prefix}-${Date.now().toString(36)}`;
  };

  const handleCreate = () => {
    const source = parseRangeInput(rangeInput, sheetId);
    if (!source) return;
    const chartId = createId('chart');
    onCommand({
      commandId: `chart.insert.${type}`,
      params: {
        sheetId,
        chartId,
        drawingId: createId('drawing'),
        title,
        sourceRanges: [{ sheetId, ...source }],
        bounds: { x: 100, y: 100, width: 480, height: 280 },
        stacked: stacked === 'none' ? undefined : stacked,
      },
    });
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Chart Settings</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Chart Title
            </Text>
            <TextInput
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="e.g. Monthly Revenue"
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Chart Type
            </Text>
            <Select
              value={type}
              onChange={(e) => setType(e.target.value as ChartDrawingPayload['chartType'])}
              sizeVariant="sm"
            >
              <option value="column">Column Chart</option>
              <option value="bar">Bar Chart</option>
              <option value="line">Line Chart</option>
              <option value="area">Area Chart</option>
              <option value="pie">Pie Chart</option>
              <option value="doughnut">Doughnut Chart</option>
              <option value="scatter">Scatter Plot</option>
              <option value="combo">Combo Chart</option>
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Stacking
            </Text>
            <Select value={stacked} onChange={(e) => setStacked(e.target.value as NonNullable<ChartDrawingPayload['stacked']>)} sizeVariant="sm">
              <option value="none">Grouped</option>
              <option value="stacked">Stacked</option>
              <option value="percent">100% Stacked</option>
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Data Source Range
            </Text>
            <TextInput
              value={rangeInput}
              onChange={(e) => setRangeInput(e.target.value)}
              placeholder="e.g. A1:F6"
            />
          </Box>

          <Button variant="primary" size="sm" icon="plus" disabled={!sourceRange} onClick={handleCreate}>
            Insert Chart to Canvas
          </Button>

          {chartEntries.length > 0 ? (
            <Box className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Worksheet Charts ({chartEntries.length})
              </Text>
              <Stack gap="xs">
                {chartEntries.map(({ drawing, payload }) => (
                  <Box
                    key={drawing.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <Stack gap="none">
                      <Text size="sm" weight="medium" className="text-slate-800">{payload.title || payload.chartType}</Text>
                      <Text size="xs" tone="subtle">{payload.chartType.toUpperCase()} Chart</Text>
                    </Stack>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onCommand({
                        commandId: 'chart.remove',
                        params: { sheetId, chartId: payload.chartId },
                      })}
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
