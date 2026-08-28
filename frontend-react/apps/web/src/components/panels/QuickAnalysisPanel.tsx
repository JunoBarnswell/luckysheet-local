import { useMemo } from 'react';
import { Box, Button, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Stack, Text } from '@react-sheets/ui-system';
import type { RangeRef } from '@react-sheets/core-model';
import type { CanvasSheetSnapshot } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../../i18n';

export interface QuickAnalysisPanelProps {
  locale: Locale;
  sheet: CanvasSheetSnapshot;
  range: RangeRef;
  onPanelChange: (panel: 'chart' | 'conditionalFormat' | 'dataValidation') => void;
}

export interface QuickAnalysisSummary {
  cells: number;
  nonBlank: number;
  numeric: number;
  sum: number;
  minimum?: number;
  maximum?: number;
}

export function summarizeQuickAnalysis(sheet: CanvasSheetSnapshot, range: RangeRef): QuickAnalysisSummary {
  const summary: QuickAnalysisSummary = { cells: 0, nonBlank: 0, numeric: 0, sum: 0 };
  for (let row = range.startRow; row <= range.endRow; row += 1) {
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      summary.cells += 1;
      const cell = sheet.getCell(row, column);
      const raw = cell?.value.trim() ?? '';
      if (!raw) continue;
      summary.nonBlank += 1;
      const numeric = Number(raw.replace(/[$,%]/g, ''));
      if (!Number.isFinite(numeric)) continue;
      summary.numeric += 1;
      const value = raw.includes('%') ? numeric / 100 : numeric;
      summary.sum += value;
      summary.minimum = summary.minimum === undefined ? value : Math.min(summary.minimum, value);
      summary.maximum = summary.maximum === undefined ? value : Math.max(summary.maximum, value);
    }
  }
  return summary;
}

export function QuickAnalysisPanel({ locale, sheet, range, onPanelChange }: QuickAnalysisPanelProps) {
  const summary = useMemo(() => summarizeQuickAnalysis(sheet, range), [range, sheet]);
  const isChinese = locale === 'zh-CN';
  const label = (zh: string, en: string) => isChinese ? zh : en;
  const formatNumber = (value: number | undefined) => value === undefined ? '—' : new Intl.NumberFormat(isChinese ? 'zh-CN' : 'en-US', { maximumFractionDigits: 6 }).format(value);
  return (
    <Panel className="h-full border-0 bg-transparent shadow-none">
      <PanelHeader className="h-12 border-b border-slate-200 px-4">
        <PanelTitle size="sm">{label('快速分析', 'Quick Analysis')}</PanelTitle>
      </PanelHeader>
      <PanelBody className="p-4">
        <Stack gap="md">
          <Box className="rounded-lg border border-blue-200 bg-blue-50 p-3">
            <Text size="xs" tone="subtle">{label('当前选区', 'Current selection')}</Text>
            <Text size="sm" weight="semibold" className="mt-1 text-blue-900">{range.startRow + 1}:{range.startColumn + 1} – {range.endRow + 1}:{range.endColumn + 1}</Text>
          </Box>
          <Stack gap="none" className="divide-y divide-slate-100 rounded-lg border border-slate-200 bg-white">
            <Inline gap="sm" className="justify-between px-3 py-2"><Text size="xs" tone="muted">{label('单元格', 'Cells')}</Text><Text size="xs" weight="semibold">{summary.cells}</Text></Inline>
            <Inline gap="sm" className="justify-between px-3 py-2"><Text size="xs" tone="muted">{label('非空', 'Non-blank')}</Text><Text size="xs" weight="semibold">{summary.nonBlank}</Text></Inline>
            <Inline gap="sm" className="justify-between px-3 py-2"><Text size="xs" tone="muted">{label('数值', 'Numeric')}</Text><Text size="xs" weight="semibold">{summary.numeric}</Text></Inline>
            <Inline gap="sm" className="justify-between px-3 py-2"><Text size="xs" tone="muted">{label('合计', 'Sum')}</Text><Text size="xs" weight="semibold">{formatNumber(summary.sum)}</Text></Inline>
            <Inline gap="sm" className="justify-between px-3 py-2"><Text size="xs" tone="muted">{label('最小值', 'Minimum')}</Text><Text size="xs" weight="semibold">{formatNumber(summary.minimum)}</Text></Inline>
            <Inline gap="sm" className="justify-between px-3 py-2"><Text size="xs" tone="muted">{label('最大值', 'Maximum')}</Text><Text size="xs" weight="semibold">{formatNumber(summary.maximum)}</Text></Inline>
          </Stack>
          <Stack gap="xs">
            <Text size="xs" weight="semibold" className="text-slate-700">{label('分析操作', 'Analysis actions')}</Text>
            <Button size="sm" variant="secondary" onClick={() => onPanelChange('chart')}>{label('插入图表', 'Insert chart')}</Button>
            <Button size="sm" variant="secondary" onClick={() => onPanelChange('conditionalFormat')}>{label('突出显示数据', 'Highlight data')}</Button>
            <Button size="sm" variant="secondary" onClick={() => onPanelChange('dataValidation')}>{label('设置数据验证', 'Set validation')}</Button>
          </Stack>
        </Stack>
      </PanelBody>
    </Panel>
  );
}
