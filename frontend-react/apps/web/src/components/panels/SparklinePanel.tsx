import React, { useState } from 'react';
import { Box, Button, ColorPicker, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { SparklineModel } from '@react-sheets/core-model';
import { parseRangeInput } from '../../domain/range-input';

export interface SparklinePanelProps {
  sheetId: string;
  sparklines: SparklineModel[];
  defaultRange?: string;
  onAddSparkline: (sparkline: SparklineModel) => void;
  onRemoveSparkline: (id: string) => void;
  onClose?: () => void;
}

export function SparklinePanel({
  sheetId,
  sparklines,
  defaultRange,
  onAddSparkline,
  onRemoveSparkline,
  onClose,
}: SparklinePanelProps) {
  const [type, setType] = useState<SparklineModel['type']>('line');
  const [sourceRange, setSourceRange] = useState(defaultRange ?? 'B2:E2');
  const [targetCell, setTargetCell] = useState('F2');
  const [color, setColor] = useState('#2563eb');
  const [highlightMax, setHighlightMax] = useState(true);
  const [highlightMin, setHighlightMin] = useState(true);
  const parsedSourceRange = parseRangeInput(sourceRange, sheetId);

  const handleCreate = () => {
    if (!parsedSourceRange) return;
    // Parse target cell
    const match = /^([A-Z]+)(\d+)$/.exec(targetCell.toUpperCase());
    let r = 1;
    let c = 5;
    if (match?.[1] && match[2]) {
      let col = 0;
      for (const char of match[1]) col = col * 26 + char.charCodeAt(0) - 64;
      c = col - 1;
      r = Number(match[2]) - 1;
    }

    const newSparkline: SparklineModel = {
      id: 'spark-' + Math.random().toString(36).substring(2, 7),
      sheetId,
      anchor: { row: r, column: c },
      sourceRange: {
        sheetId,
        ...parsedSourceRange,
      },
      type,
      color,
      negativeColor: '#ef4444',
      highlightMax,
      highlightMin,
    };
    onAddSparkline(newSparkline);
  };

  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">Sparkline Mini Charts</PanelTitle>
      </PanelHeader>

      <PanelBody className="p-4">
        <Stack gap="md">
          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Sparkline Style
            </Text>
            <Select
              value={type}
              onChange={(e) => setType(e.target.value as SparklineModel['type'])}
              sizeVariant="sm"
            >
              <option value="line">Line Trend (with Min/Max dots)</option>
              <option value="column">Column Bars</option>
              <option value="win-loss">Win / Loss Indicator</option>
            </Select>
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Data Source Range
            </Text>
            <TextInput
              value={sourceRange}
              onChange={(e) => setSourceRange(e.target.value)}
              placeholder="e.g. B2:E2"
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Target Cell
            </Text>
            <TextInput
              value={targetCell}
              onChange={(e) => setTargetCell(e.target.value)}
              placeholder="e.g. F2"
            />
          </Box>

          <Box>
            <Text size="xs" weight="medium" className="mb-1 text-slate-700">
              Theme Color
            </Text>
            <Stack gap="xs">
              <ColorPicker color={color} onChange={setColor} />
              <TextInput value={color} onChange={(e) => setColor(e.target.value)} className="h-8 text-xs font-mono" />
            </Stack>
          </Box>

          <Button variant="primary" size="sm" icon="sparkline" disabled={!parsedSourceRange} onClick={handleCreate}>
            Insert In-Cell Sparkline
          </Button>

          {sparklines.length > 0 ? (
            <Box className="mt-4 border-t border-slate-200 pt-3">
              <Text size="xs" weight="semibold" className="mb-2 text-slate-700">
                Cell Sparklines ({sparklines.length})
              </Text>
              <Stack gap="xs">
                {sparklines.map((s) => (
                  <Box
                    key={s.id}
                    className="flex items-center justify-between rounded-lg border border-slate-200 bg-white p-2 text-xs"
                  >
                    <Stack gap="none">
                      <Text size="sm" weight="medium" className="text-slate-800">
                        {s.type.toUpperCase()} at R{s.anchor.row + 1}C{s.anchor.column + 1}
                      </Text>
                    </Stack>
                    <Button
                      variant="ghost"
                      size="xs"
                      icon="trash"
                      iconOnly
                      onClick={() => onRemoveSparkline(s.id)}
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
