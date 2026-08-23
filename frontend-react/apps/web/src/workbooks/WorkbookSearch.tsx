import { Button, Inline, TextInput } from '@react-sheets/ui-system';

export interface WorkbookSearchProps {
  value: string;
  onChange: (value: string) => void;
  onFilter: () => void;
  onToggleView: () => void;
  viewMode: 'list' | 'grid';
}

export function WorkbookSearch({ value, onChange, onFilter, onToggleView, viewMode }: WorkbookSearchProps) {
  return (
    <Inline gap="sm" className="w-full justify-end">
      <TextInput aria-label="搜索文件、位置或人员" className="h-10 w-full max-w-[330px] rounded-lg border-slate-300 bg-white text-[13px]" leadingIcon="search" onChange={(event) => onChange(event.target.value)} placeholder="搜索文件、位置或人员" value={value} />
      <Button aria-label="筛选" icon="filter" iconOnly onClick={onFilter} size="md" variant="outline" className="h-10 w-10 border-slate-300 text-slate-600" />
      <Button aria-label={viewMode === 'list' ? '切换卡片视图' : '切换列表视图'} icon={viewMode === 'list' ? 'grid' : 'menu'} iconOnly onClick={onToggleView} size="md" variant="outline" className="h-10 w-10 border-slate-300 text-slate-600" />
    </Inline>
  );
}
