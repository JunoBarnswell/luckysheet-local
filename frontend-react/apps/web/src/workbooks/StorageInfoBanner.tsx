import { Box, Button, Icon, Inline, Text } from '@react-sheets/ui-system';

export interface StorageInfoBannerProps {
  onLearnMore: () => void;
}

export function StorageInfoBanner({ onLearnMore }: StorageInfoBannerProps) {
  return (
    <Box className="flex min-h-[68px] w-full items-center justify-between gap-4 rounded-lg border border-brand-line bg-gradient-to-r from-brand-pale via-[#f4faf6] to-[#edf6f1] px-5 py-3">
      <Inline gap="md" className="min-w-0">
        <Box className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full border-2 border-brand text-brand"><Icon name="info" size="sm" /></Box>
        <Text className="truncate text-[14px] text-slate-700">当前本地工作簿仅保存在此页面的内存会话中，刷新或关闭页面后会清空；云端工作簿仍由服务端保存。</Text>
      </Inline>
      <Button onClick={onLearnMore} size="sm" variant="outline" className="shrink-0 border-brand/40 text-brand-dark hover:bg-white">了解更多 <Icon name="arrow-right" size="sm" /></Button>
    </Box>
  );
}
