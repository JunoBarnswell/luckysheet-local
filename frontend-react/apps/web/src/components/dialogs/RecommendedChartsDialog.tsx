import type { ChartRecommendation } from '@react-sheets/spreadsheet-app';
import { Box, Button, Dialog, Icon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';

export interface RecommendedChartsDialogProps {
  open: boolean;
  locale: Locale;
  candidates: readonly ChartRecommendation[];
  error?: string;
  onClose: () => void;
  onSelect: (candidate: ChartRecommendation) => void;
}

const chartIcon = (type: ChartRecommendation['chartType']): 'chart-column' | 'chart-line' | 'chart-pie' | 'chart-scatter' => {
  if (type === 'line') return 'chart-line';
  if (type === 'pie' || type === 'doughnut') return 'chart-pie';
  if (type === 'scatter' || type === 'bubble') return 'chart-scatter';
  return 'chart-column';
};

export function RecommendedChartsDialog({ open, locale, candidates, error, onClose, onSelect }: RecommendedChartsDialogProps) {
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
      footer={<Button size="sm" variant="ghost" onClick={onClose}>{zh ? '取消' : 'Cancel'}</Button>}
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
              <Box className="flex h-16 w-24 shrink-0 items-center justify-center rounded border border-slate-200 bg-white">
                <Icon name={chartIcon(candidate.chartType)} size="xl" className="text-[#217346]" />
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
