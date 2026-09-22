import { Box, Button, Icon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { WorkbookBackstageShellProps } from './types';
import { WorkbookTopBar } from './WorkbookTopBar';
import { WorkbookStatusBadge } from './WorkbookStatusBadge';

const statusKind = (status: WorkbookBackstageShellProps['syncStatus']) => status;

export function WorkbookBackstageShell({ workbookName, syncStatus, readOnly = false, onBack, onHelp, onSettings, actions, activeActionId, children }: WorkbookBackstageShellProps) {
  return (
    <Box as="main" className="flex min-h-screen flex-col bg-white text-ink" data-testid="workbook-backstage">
      <WorkbookTopBar onBrandClick={onBack} onHelp={onHelp} onSettings={onSettings} saveLabel="工作簿信息" workbookName={workbookName} />
      <Box className="flex min-h-0 flex-1">
        <Stack as="nav" aria-label="工作簿文件操作" gap="xs" className="w-[clamp(208px,22vw,260px)] shrink-0 overflow-y-auto border-r border-slate-200 bg-slate-50/90 p-4">
          <Button className="mb-3 w-full !justify-start rounded-xl border-b border-slate-200 pb-4 text-[14px] font-semibold text-slate-800" icon="arrow-left" onClick={onBack} size="md" variant="ghost">返回工作簿</Button>
          {actions.map((action) => (
            <Button key={action.id} aria-current={activeActionId === action.id ? 'page' : undefined} aria-disabled={action.disabled || undefined} className="h-[68px] min-h-[68px] w-full !justify-start gap-3 overflow-hidden rounded-xl px-3 py-2.5 text-left" disabled={action.disabled} icon={action.icon} onClick={action.onSelect} size="md" variant={activeActionId === action.id ? 'soft' : 'ghost'}>
              <Stack gap="none" className="min-w-0 max-w-full items-start">
                <Text className="max-w-full truncate text-[13px] font-semibold text-slate-800">{action.label}</Text>
                {action.description ? <Text title={action.description} className="w-full max-w-full truncate text-[10px] leading-4 text-slate-500">{action.description}</Text> : null}
              </Stack>
            </Button>
          ))}
        </Stack>
        <Box className="min-w-0 flex-1 overflow-y-auto">
          <Stack gap="lg" className="mx-auto max-w-[1040px] px-6 py-8 min-[1100px]:px-12">
            <Inline gap="md" className="items-start justify-between">
              <Stack gap="xs">
                <Text className="text-[12px] font-medium uppercase tracking-[0.12em] text-slate-400">文件</Text>
                <Text className="text-[26px] font-semibold tracking-[-0.03em] text-slate-900" weight="semibold">{workbookName}</Text>
                <Inline gap="sm"><WorkbookStatusBadge item={{ unitId: 'backstage', name: workbookName, updatedAt: new Date().toISOString(), locationLabel: '当前工作簿', storageLocation: 'remote', syncStatus: statusKind(syncStatus), lifecycle: 'active', role: readOnly ? 'viewer' : 'editor', sourceKind: 'native', favorite: false }} /><Text size="xs" tone="muted">{readOnly ? '只读模式' : '可编辑'}</Text></Inline>
              </Stack>
              <Button icon="arrow-left" onClick={onBack} size="sm" variant="outline">返回编辑器</Button>
            </Inline>
            {children ?? <Box className="rounded-xl border border-dashed border-brand-line bg-brand-pale p-8 text-center"><Icon name="file-text" size="xl" className="mb-3 text-brand" /><Text as="p" className="text-sm text-slate-600">选择左侧操作查看工作簿信息。</Text></Box>}
          </Stack>
        </Box>
      </Box>
    </Box>
  );
}
