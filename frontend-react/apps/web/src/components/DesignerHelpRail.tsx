import { Box, Button, Icon, Stack, Text } from '@react-sheets/ui-system';

const DESIGNER_HELP_LINKS = [
  { id: 'consult', label: '在线咨询', icon: 'comment' as const, color: 'text-[#c13e9c]', href: 'https://www.grapecity.com.cn/contact' },
  { id: 'demos', label: '更多示例', icon: 'external-link' as const, color: 'text-[#d3479f]', href: 'https://www.grapecity.com.cn/developer/spreadjs/demo' },
  { id: 'download', label: '立即下载', icon: 'download' as const, color: 'text-[#a43cbd]', href: 'https://www.grapecity.com.cn/developer/spreadjs/download' },
];

/** The reference Designer exposes these real product links as a floating rail. */
export function DesignerHelpRail() {
  return (
    <Box aria-label="Designer help links" className="pointer-events-auto absolute right-0 top-1/2 z-30 -translate-y-1/2 overflow-hidden rounded-l-md border border-slate-200 bg-white shadow-[0_2px_10px_rgba(61,60,65,0.18)]">
      <Stack gap="none">
        {DESIGNER_HELP_LINKS.map((link) => (
          <Button
            key={link.id}
            aria-label={link.label}
            onClick={() => window.open(link.href, '_blank', 'noopener,noreferrer')}
            size="xs"
            variant="ghost"
            className="h-[68px] w-[70px] flex-col gap-1 rounded-none border-b border-slate-100 px-1 text-slate-600 last:border-b-0 hover:bg-[#f5f4f5] hover:text-[#217345]"
          >
            <Icon name={link.icon} size="md" className={link.color} />
            <Text as="span" size="xs" className="text-[10px] leading-3">{link.label}</Text>
          </Button>
        ))}
      </Stack>
    </Box>
  );
}
