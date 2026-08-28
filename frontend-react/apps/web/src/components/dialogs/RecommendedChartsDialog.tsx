import type { ReactNode } from 'react';
import type { ChartRecommendation } from '@react-sheets/spreadsheet-app';
import { Box, Button, Dialog, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';

export interface RecommendedChartsDialogProps {
  open: boolean;
  locale: Locale;
  candidates: readonly ChartRecommendation[];
  error?: string;
  onClose: () => void;
  onSelect: (candidate: ChartRecommendation) => void;
  onOpenAllCharts?: () => void;
}

const PREVIEW_HEIGHTS = ['h-1', 'h-2', 'h-3', 'h-4', 'h-5', 'h-6', 'h-7', 'h-8'] as const;

function previewBars(candidate: ChartRecommendation): ReactNode {
  const values = candidate.preview.series[0]?.values.map((value) => typeof value === 'number' ? value : typeof value === 'string' ? Number(value.replace(/[$,%]/g, '')) : NaN).filter(Number.isFinite) ?? [];
  if (!values.length) return <Text size="xs" tone="muted">No numeric preview</Text>;
  const max = Math.max(1, ...values.map(Math.abs));
  return <Inline gap="xs" className="h-12 items-end"><>{values.slice(0, 12).map((value, index) => <Box key={`${candidate.id}-${index}`} className={`w-2 rounded-t bg-[#217346] ${PREVIEW_HEIGHTS[Math.min(PREVIEW_HEIGHTS.length - 1, Math.max(0, Math.round(Math.abs(value) / max * (PREVIEW_HEIGHTS.length - 1))))]}`} />)}</></Inline>;
}

export function RecommendedChartsDialog({ open, locale, candidates, error, onClose, onSelect, onOpenAllCharts }: RecommendedChartsDialogProps) {
  const zh = locale === 'zh-CN';
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={zh ? '推荐的图表' : 'Recommended Charts'}
      description={zh ? '根据当前选区的数据类型和结构生成候选图表。' : 'Candidates are derived from the current selection data types and structure.'}
      closeLabel={zh ? '关闭' : 'Close'}
      maxWidth="lg"
      testId="recommended-charts-dialog"
      footer={<Inline gap="sm"><Button size="sm" variant="secondary" onClick={onOpenAllCharts}>{zh ? '所有图表' : 'All Charts'}</Button><Button size="sm" variant="ghost" onClick={onClose}>{zh ? '取消' : 'Cancel'}</Button></Inline>}
    >
      {error ? <Box className="rounded border border-rose-300 bg-rose-50 p-3"><Text size="sm" tone="danger">{error}</Text></Box> : null}
      {!error && candidates.length === 0 ? <Text size="sm" tone="muted">{zh ? '当前选区没有可推荐的图表。' : 'No chart can be recommended for the current selection.'}</Text> : null}
      <Stack gap="sm">
        {candidates.map((candidate, index) => (
          <Button
            key={candidate.id}
            aria-label={candidate.title}
            data-testid={`recommended-chart-${candidate.id}`}
            variant={index === 0 ? 'secondary' : 'ghost'}
            className="h-auto w-full justify-start rounded border border-slate-200 p-3 text-left"
            onClick={() => onSelect(candidate)}
          >
            <Inline gap="md" className="w-full items-center">
              <Box className="flex h-16 w-24 shrink-0 items-center justify-center rounded border border-slate-200 bg-white px-2">
                {previewBars(candidate)}
              </Box>
              <Stack gap="xs" className="min-w-0 flex-1 items-start">
                <Text size="sm" weight="semibold">{candidate.title}</Text>
                <Text size="xs" tone="muted">{candidate.chartType} / {candidate.subtype} · {candidate.reason}</Text>
                <Text size="xs" tone="subtle">{zh ? '匹配度' : 'Confidence'} {Math.round(candidate.confidence * 100)}%</Text>
              </Stack>
            </Inline>
          </Button>
        ))}
      </Stack>
    </Dialog>
  );
}
