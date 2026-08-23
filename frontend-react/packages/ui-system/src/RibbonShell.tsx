import { useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { Box, Inline, Text } from './layout';
import { RIBBON_TAB_ORDER, type RibbonLayoutMode, type RibbonLayoutState, type RibbonTabId } from './shell-types';
import { Tab, TabList, Tabs } from './Tabs';

export type { RibbonLayoutMode, RibbonLayoutState, RibbonTabId } from './shell-types';
export { RIBBON_TAB_ORDER } from './shell-types';

export function ribbonLayoutModeForWidth(width: number): RibbonLayoutMode {
  if (width >= 1120) return 'wide';
  if (width >= 760) return 'compact';
  return 'narrow';
}

export interface RibbonShellProps {
  activeTab: RibbonTabId;
  children: ReactNode | ((layout: RibbonLayoutState) => ReactNode);
  disabled?: boolean;
  onTabChange: (tab: RibbonTabId) => void;
  status?: ReactNode;
  tabLabel: (tab: RibbonTabId) => string;
}

export function RibbonShell({
  activeTab,
  children,
  disabled = false,
  onTabChange,
  status,
  tabLabel,
}: RibbonShellProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<RibbonLayoutState>({ mode: 'wide', width: 1120 });

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
    <Tabs ref={rootRef} className="border-b border-slate-200 bg-white" data-ribbon-layout={layout.mode} data-testid="ribbon-shell">
      <Inline gap="lg" className="min-h-10 flex-wrap px-4 py-1">
        <TabList label="Workbook ribbon tabs" className="h-full flex-wrap gap-1">
          {RIBBON_TAB_ORDER.map((tab) => (
            <Tab
              key={tab}
              active={activeTab === tab}
              data-testid={`ribbon-tab-${tab}`}
              disabled={disabled}
              onClick={() => onTabChange(tab)}
              className="h-full border-b-2 border-transparent px-3 text-xs font-semibold data-active:border-blue-600 data-active:text-blue-600"
            >
              {tabLabel(tab)}
            </Tab>
          ))}
        </TabList>
        {status ? (
          <Inline gap="xs" className="ml-auto shrink-0 border-l border-slate-100 pl-3">
            {status}
          </Inline>
        ) : null}
      </Inline>
      <Box className="min-h-[104px] overflow-hidden border-t border-slate-100 bg-slate-50/80 px-4 py-2">
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
