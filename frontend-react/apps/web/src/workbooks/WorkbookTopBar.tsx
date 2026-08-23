import { Box, Button, Divider, Icon, Inline, Stack, Text } from '@react-sheets/ui-system';

export interface WorkbookTopBarProps {
  workbookName?: string;
  saveLabel?: string;
  onHelp: () => void;
  onSettings: () => void;
  onBrandClick: () => void;
}

export function WorkbookTopBar({ workbookName, saveLabel = '已保存到服务器', onHelp, onSettings, onBrandClick }: WorkbookTopBarProps) {
  return (
    <Box as="header" className="flex h-[58px] shrink-0 items-center justify-between border-b border-slate-200 bg-white px-5 text-ink">
      <Inline gap="md" className="min-w-0">
        <Button
          aria-label="返回工作簿中心"
          className="h-7 w-7 rounded-md bg-brand p-0 text-white shadow-none hover:bg-brand-dark"
          icon="grid"
          iconOnly
          onClick={onBrandClick}
          size="sm"
          variant="brand"
        />
        <Text className="text-[21px] font-bold tracking-[-0.04em] text-slate-900" weight="bold">云表格</Text>
        <Divider orientation="vertical" className="mx-1 h-6 bg-slate-200" />
        <Inline gap="xs" className="min-w-0">
          <Icon name="check-circle" size="sm" className="text-green-600" />
          <Text className="truncate text-[14px] text-slate-500" size="md">{saveLabel}</Text>
          {workbookName ? <Text className="max-w-[360px] truncate text-[14px] text-slate-500" size="md">· {workbookName}</Text> : null}
        </Inline>
      </Inline>

      <Inline gap="xs" className="shrink-0" aria-hidden="false">
        <Button aria-label="帮助" icon="help" iconOnly onClick={onHelp} size="sm" variant="ghost" className="h-8 w-8 text-slate-700 hover:bg-slate-100" />
        <Button aria-label="设置" icon="settings" iconOnly onClick={onSettings} size="sm" variant="ghost" className="h-8 w-8 text-slate-700 hover:bg-slate-100" />
        <Divider orientation="vertical" className="mx-2 h-6 bg-slate-200" />
        <Stack gap="none" aria-hidden="true" className="flex-row items-center gap-5 px-1 text-slate-700">
          <Icon name="minimize" size="sm" />
          <Icon name="maximize" size="sm" />
          <Icon name="x" size="sm" />
        </Stack>
      </Inline>
    </Box>
  );
}
