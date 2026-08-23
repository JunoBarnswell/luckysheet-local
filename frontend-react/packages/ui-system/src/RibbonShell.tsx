import type { ReactNode } from 'react';
import { Box, Inline, Text } from './layout';
import { RIBBON_TAB_ORDER, type RibbonTabId } from './shell-types';
import { Tab, TabList, Tabs } from './Tabs';

export type { RibbonTabId } from './shell-types';
export { RIBBON_TAB_ORDER } from './shell-types';

export interface RibbonShellProps {
  activeTab: RibbonTabId;
  children: ReactNode;
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
  return (
    <Tabs className="border-b border-slate-200 bg-white" data-testid="ribbon-shell">
      <Inline gap="lg" className="h-10 overflow-x-auto px-4">
        <TabList label="Workbook ribbon tabs" className="h-full gap-1">
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
      <Box className="min-h-[104px] overflow-x-auto border-t border-slate-100 bg-slate-50/80 px-4 py-2">
        {children}
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
