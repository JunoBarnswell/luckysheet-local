import { useEffect, useState } from 'react';
import { Button, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import type { RangeRef, SheetTableModel } from '@react-sheets/core-model';
import type { Locale } from '../../i18n';

export type TableDesignOption = 'hasHeaderRow' | 'showFirstColumn' | 'showLastColumn' | 'showBandedRows' | 'showBandedColumns' | 'showFilterButton';

export interface TableDesignPanelProps {
  table: SheetTableModel;
  locale: Locale;
  selectedRange?: RangeRef;
  onNameChange: (name: string) => void;
  onToggle: (option: TableDesignOption) => void;
  onResize: (range: RangeRef) => void;
  onStyleChange: (styleName: string) => void;
  onConvert: () => void;
}

export function TableDesignPanel({ table, locale, selectedRange, onNameChange, onToggle, onResize, onStyleChange, onConvert }: TableDesignPanelProps) {
  const isZh = locale === 'zh-CN';
  const [name, setName] = useState(table.name);
  useEffect(() => setName(table.name), [table.id, table.name]);
  const toggle = (option: TableDesignOption, label: string, enabled: boolean) => (
    <Button size="sm" variant={enabled ? 'secondary' : 'ghost'} aria-pressed={enabled} className="justify-start" onClick={() => onToggle(option)}>{label}: {enabled ? (isZh ? '开' : 'On') : (isZh ? '关' : 'Off')}</Button>
  );
  return (
    <Stack gap="sm">
      <Panel className="shadow-none"><PanelHeader><PanelTitle as="h3" size="sm">{isZh ? '表格设计' : 'Table Design'}</PanelTitle></PanelHeader><PanelBody><Stack gap="sm">
        <Inline gap="sm"><TextInput aria-label={isZh ? '表名称' : 'Table name'} value={name} onChange={(event) => setName(event.target.value)} onKeyDown={(event) => { if (event.key === 'Enter') onNameChange(name); }} /><Button size="sm" variant="outline" disabled={!name.trim() || name.trim() === table.name} onClick={() => onNameChange(name)}>{isZh ? '应用' : 'Apply'}</Button></Inline>
        <Select aria-label={isZh ? '表格样式' : 'Table style'} sizeVariant="sm" value={table.styleName ?? 'TableStyleMedium2'} onChange={(event) => onStyleChange(event.currentTarget.value)} options={[
          { value: 'TableStyleMedium2', label: 'Medium 2' },
          { value: 'TableStyleMedium4', label: 'Medium 4' },
          { value: 'TableStyleMedium9', label: 'Medium 9' },
          { value: 'TableStyleLight1', label: 'Light 1' },
          { value: 'TableStyleLight2', label: 'Light 2' },
        ]} />
        {toggle('hasHeaderRow', isZh ? '标题行' : 'Header Row', table.hasHeaderRow)}
        {toggle('showFirstColumn', isZh ? '第一列' : 'First Column', table.showFirstColumn)}
        {toggle('showLastColumn', isZh ? '最后一列' : 'Last Column', table.showLastColumn)}
        {toggle('showBandedRows', isZh ? '带状行' : 'Banded Rows', table.showBandedRows)}
        {toggle('showBandedColumns', isZh ? '带状列' : 'Banded Columns', table.showBandedColumns)}
        {toggle('showFilterButton', isZh ? '筛选按钮' : 'Filter Button', table.showFilterButton)}
        <Inline gap="sm"><Button size="sm" variant="outline" disabled={!selectedRange} onClick={() => { if (selectedRange) onResize(selectedRange); }}>{isZh ? '调整表格大小' : 'Resize Table'}</Button><Button size="sm" variant="danger" onClick={onConvert}>{isZh ? '转换为区域' : 'Convert to Range'}</Button></Inline>
      </Stack></PanelBody></Panel>
      <Panel className="shadow-none"><PanelBody><Text size="xs" tone="muted">{isZh ? `范围：${table.range.startRow + 1}:${table.range.endRow + 1} · ${table.styleName ?? '默认样式'}` : `Range ${table.range.startRow + 1}:${table.range.endRow + 1} · ${table.styleName ?? 'Default style'}`}</Text></PanelBody></Panel>
    </Stack>
  );
}
