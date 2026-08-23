import { Tab, TabList, Tabs } from '@react-sheets/ui-system';
import type { WorkbookCategoryTab } from './types';

export interface WorkbookCategoryTabsProps {
  activeTab: WorkbookCategoryTab;
  onChange: (tab: WorkbookCategoryTab) => void;
}

const tabs: readonly { id: WorkbookCategoryTab; label: string }[] = [
  { id: 'recent', label: '最近' },
  { id: 'cloud', label: '我的云文档' },
  { id: 'local', label: '本地文件' },
  { id: 'shared', label: '与我共享' },
];

export function WorkbookCategoryTabs({ activeTab, onChange }: WorkbookCategoryTabsProps) {
  return (
    <Tabs className="border-b border-slate-200">
      <TabList label="工作簿分类" className="gap-1">
        {tabs.map((tab) => <Tab key={tab.id} active={tab.id === activeTab} onClick={() => onChange(tab.id)}>{tab.label}</Tab>)}
      </TabList>
    </Tabs>
  );
}
