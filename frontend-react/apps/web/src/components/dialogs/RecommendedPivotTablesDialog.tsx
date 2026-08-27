import type { PivotTableRecommendation } from '@react-sheets/spreadsheet-app';
import { Box, Button, Dialog, Icon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';

export interface RecommendedPivotTablesDialogProps {
  open: boolean;
  locale: Locale;
  candidates: readonly PivotTableRecommendation[];
  error?: string;
  onClose: () => void;
  onSelect: (candidate: PivotTableRecommendation) => Promise<void>;
}

export function RecommendedPivotTablesDialog({ open, locale, candidates, error, onClose, onSelect }: RecommendedPivotTablesDialogProps) {
  const zh = locale === 'zh-CN';
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={zh ? '推荐的数据透视表' : 'Recommended PivotTables'}
      description={zh ? '根据当前数据区域的字段类型生成真实数据透视表布局。' : 'Layouts are derived from field types in the current data region.'}
      closeLabel={zh ? '关闭' : 'Close'}
      maxWidth="lg"
      testId="recommended-pivot-tables-dialog"
      footer={<Button size="sm" variant="ghost" onClick={onClose}>{zh ? '取消' : 'Cancel'}</Button>}
    >
      {error ? <Box className="rounded border border-rose-300 bg-rose-50 p-3"><Text size="sm" tone="danger">{error}</Text></Box> : null}
      <Stack gap="sm">
        {candidates.map((candidate, index) => (
          <Button
            key={candidate.id}
            aria-label={candidate.title}
            data-testid={`recommended-pivot-${candidate.id}`}
            variant={index === 0 ? 'secondary' : 'ghost'}
            className="h-auto w-full justify-start rounded border border-slate-200 p-3 text-left"
            onClick={() => { void onSelect(candidate); }}
          >
            <Inline gap="md" className="w-full items-center">
              <Box className="flex h-16 w-24 shrink-0 items-center justify-center rounded border border-slate-200 bg-white">
                <Icon name="table-pivot" size="xl" className="text-[#217346]" />
              </Box>
              <Stack gap="xs" className="min-w-0 flex-1 items-start">
                <Text size="sm" weight="semibold">{candidate.title}</Text>
                <Text size="xs" tone="muted">{candidate.summary}</Text>
                <Text size="xs" tone="subtle">{zh ? '匹配度' : 'Confidence'} {Math.round(candidate.confidence * 100)}%</Text>
              </Stack>
            </Inline>
          </Button>
        ))}
      </Stack>
    </Dialog>
  );
}
