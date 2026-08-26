import { useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { Button } from './Button';
import { DropdownMenu } from './DropdownMenu';
import { Box, Inline, Text } from './layout';
import { RIBBON_TAB_ORDER, type RibbonLayoutMode, type RibbonLayoutState, type RibbonTabId } from './shell-types';
import { Tab, TabList, Tabs } from './Tabs';

export type { RibbonLayoutMode, RibbonLayoutState, RibbonTabId } from './shell-types';
export { RIBBON_TAB_ORDER } from './shell-types';

export function ribbonLayoutModeForWidth(width: number): RibbonLayoutMode {
  if (width >= 1920) return 'wide';
  if (width >= 1024) return 'compact';
  return 'narrow';
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
  const [layout, setLayout] = useState<RibbonLayoutState>({ mode: 'wide', width: 1920 });
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
    <Tabs ref={rootRef} className="h-[157px] overflow-hidden border-b border-[#e7e7e7] bg-[#f5f5f3]" data-ribbon-layout={layout.mode} data-testid="ribbon-shell">
      <Inline gap="none" className="h-[42px] min-w-0 flex-nowrap px-2">
        {onFileEntry ? (
          <DropdownMenu
            disabled={disabled}
            trigger={(
              <Button
                aria-label="Open workbook menu"
                className={`h-full shrink-0 rounded-none border-b-2 border-transparent px-0 text-xs font-semibold text-slate-700 hover:border-[#217345] hover:bg-[#f3f8f4] hover:text-[#217345] ${layout.mode === 'wide' ? 'w-[52px]' : 'w-[42px]'}`}
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
              className={`!h-full !min-h-0 min-w-0 shrink rounded-none border-b-2 border-transparent font-semibold text-[#3d3c41] aria-selected:!bg-white aria-selected:!text-[#217345] aria-selected:border-[#217345] ${layout.mode === 'wide' ? 'px-3 text-xs' : 'px-1.5 text-[11px]'}`}
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
      <Box className="h-[115px] overflow-hidden border-t-0 bg-white px-0 py-0">
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
