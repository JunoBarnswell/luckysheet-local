import { useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { Button } from './Button';
import { DropdownMenu } from './DropdownMenu';
import { Box, Inline, Text } from './layout';
import { RIBBON_DENSITY, RIBBON_TAB_ORDER, type RibbonLayoutMode, type RibbonLayoutState, type RibbonTabId } from './shell-types';
import { Tab, TabList, Tabs } from './Tabs';

export type { RibbonLayoutMode, RibbonLayoutState, RibbonTabId } from './shell-types';
export { RIBBON_TAB_ORDER } from './shell-types';

export function ribbonLayoutModeForWidth(width: number): RibbonLayoutMode {
  if (width >= 1920) return 'wide';
  if (width >= 1024) return 'compact';
  return 'narrow';
}

const DENSE_COMPACT_MIN_WIDTH = 1440;
const RIBBON_TAB_WIDTHS: Partial<Record<RibbonTabId, string>> = {
  file: 'w-[52px]',
  home: 'w-[56px]',
  insert: 'w-[64px]',
  pageLayout: 'w-[82px]',
  formulas: 'w-[60px]',
  data: 'w-[52px]',
  view: 'w-[58px]',
  review: 'w-[58px]',
  settings: 'w-[58px]',
};

function isDenseCompact(width: number, mode: RibbonLayoutMode): boolean {
  return mode === 'compact' && width >= DENSE_COMPACT_MIN_WIDTH;
}

export interface RibbonShellProps {
  activeTab: RibbonTabId;
  children: ReactNode | ((layout: RibbonLayoutState) => ReactNode);
  /** Context tabs are session state supplied by the host, never workbook data. */
  contextualTabs?: readonly RibbonTabId[];
  disabled?: boolean;
  onFileEntry?: () => void;
  onTabChange: (tab: RibbonTabId) => void;
  status?: ReactNode;
  tabLabel: (tab: RibbonTabId) => string;
}

export function RibbonShell({
  activeTab,
  children,
  contextualTabs = [],
  disabled = false,
  onFileEntry,
  onTabChange,
  status,
  tabLabel,
}: RibbonShellProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<RibbonLayoutState>(() => {
    const width = typeof window === 'undefined' ? 1920 : Math.round(window.innerWidth);
    return { mode: ribbonLayoutModeForWidth(width), width };
  });
  const denseCompact = isDenseCompact(layout.width, layout.mode);
  const tabs = [
    ...RIBBON_TAB_ORDER.filter((tab) => !onFileEntry || tab !== 'file'),
    ...contextualTabs.filter((tab, index, values) => !RIBBON_TAB_ORDER.includes(tab) && values.indexOf(tab) === index),
  ];

  useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root) return undefined;
    const update = (width: number) => {
      const nextWidth = Math.max(0, Math.round(width));
      setLayout((previous) => {
        const mode = ribbonLayoutModeForWidth(nextWidth);
        return previous.width === nextWidth && previous.mode === mode ? previous : { mode, width: nextWidth };
      });
    };
    update(root.getBoundingClientRect().width);
    if (typeof ResizeObserver === 'undefined') return undefined;
    const observer = new ResizeObserver((entries) => {
      const width = entries[0]?.contentRect.width;
      if (width !== undefined) update(width);
    });
    observer.observe(root);
    return () => observer.disconnect();
  }, []);

  return (
    <Tabs ref={rootRef} className="h-[133px] overflow-hidden border-b border-[#e7e7e7] bg-[#f5f5f3]" data-ribbon-layout={layout.mode} data-ribbon-height={RIBBON_DENSITY.shellHeight} data-testid="ribbon-shell">
      <Inline gap="none" className="h-[32px] min-w-0 flex-nowrap px-2">
        {onFileEntry ? (
          <DropdownMenu
            disabled={disabled}
            trigger={(
              <Button
                aria-label="Open workbook menu"
                className={`h-full shrink-0 rounded-none border-b-2 border-transparent px-0 font-semibold text-slate-700 hover:border-[#217345] hover:bg-[#f3f8f4] hover:text-[#217345] ${RIBBON_TAB_WIDTHS.file} ${denseCompact ? 'text-[14px]' : layout.mode === 'wide' ? 'text-xs' : 'text-[11px]'}`}
                size="sm"
                variant="ghost"
              >
                文件
              </Button>
            )}
          >
            {({ close }) => (
              <Button className="min-w-[9rem] justify-start" onClick={() => { close(); onFileEntry(); }} size="sm" variant="ghost">
                File / 工作簿
              </Button>
            )}
          </DropdownMenu>
        ) : null}
        <TabList label="Workbook ribbon tabs" className="h-full min-w-0 flex-1 flex-nowrap gap-0 overflow-hidden">
          {tabs.map((tab) => (
            <Tab
              key={tab}
              active={activeTab === tab}
              data-testid={`ribbon-tab-${tab}`}
              disabled={disabled}
              onClick={() => onTabChange(tab)}
              className={`!h-full !min-h-0 min-w-0 !shrink-0 rounded-none border-b-2 border-transparent font-semibold text-[#3d3c41] aria-selected:!bg-white aria-selected:!text-[#217345] aria-selected:border-[#217345] ${RIBBON_TAB_WIDTHS[tab] ?? 'w-[58px]'} !px-0 ${denseCompact ? 'text-[14px]' : layout.mode === 'wide' ? 'text-xs' : 'text-[11px]'}`}
            >
              {tabLabel(tab)}
            </Tab>
          ))}
        </TabList>
        {status ? (
          <Inline gap="xs" className={`${layout.mode === 'wide' ? 'flex' : 'hidden'} ml-auto shrink-0 border-l border-slate-100 pl-3`}>
            {status}
          </Inline>
        ) : null}
      </Inline>
      <Box className="h-[101px] overflow-hidden border-t-0 bg-white px-0 py-0">
        {typeof children === 'function' ? children(layout) : children}
      </Box>
    </Tabs>
  );
}

export function RibbonEmptyState({ message }: { message: string }) {
  return (
    <Inline gap="sm" className="min-h-[88px] items-center">
      <Text size="sm" tone="muted">{message}</Text>
    </Inline>
  );
}
