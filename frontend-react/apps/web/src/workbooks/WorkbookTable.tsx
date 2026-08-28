import { Button, DataTable, FileIcon, Icon, Inline, Text } from '@react-sheets/ui-system';
import type { DataTableColumn } from '@react-sheets/ui-system';
import type { WorkbookCatalogItem } from './types';
import { WorkbookRowMenu, type WorkbookRowMenuProps } from './WorkbookRowMenu';
import { WorkbookStatusBadge } from './WorkbookStatusBadge';

export interface WorkbookTableProps extends Omit<WorkbookRowMenuProps, 'item'> {
  items: readonly WorkbookCatalogItem[];
  selectedKeys: readonly string[];
  onSelectionChange: (keys: readonly string[]) => void;
  onOpen: (unitId: string) => void;
  empty: React.ReactNode;
}

function formatDate(value: string) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat('zh-CN', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }).format(date);
}

export function WorkbookTable({ items, selectedKeys, onSelectionChange, onOpen, empty, ...menuProps }: WorkbookTableProps) {
  const columns: readonly DataTableColumn<WorkbookCatalogItem>[] = [
    {
      key: 'name',
      header: '名称',
      width: '38%',
      render: (item) => (
        <Inline gap="sm" className="min-w-0">
          <FileIcon kind="native-document" size="sm" />
          <Text className="min-w-0 truncate text-[14px] text-slate-800" weight="medium">{item.name}</Text>
          <Button aria-label={item.favorite ? `取消 ${item.name} 的星标` : `为 ${item.name} 添加星标`} icon="star" iconOnly onClick={(event) => { event.stopPropagation(); menuProps.onFavorite(item.unitId, !item.favorite); }} size="xs" variant="ghost" className={item.favorite ? 'h-7 w-7 text-amber-500' : 'h-7 w-7 text-slate-300 opacity-0 group-hover:opacity-100'} />
        </Inline>
      ),
    },
    {
      key: 'location',
      header: '位置',
      width: '29%',
      render: (item) => <Text className="truncate text-[13px] text-slate-500">{item.locationLabel}</Text>,
    },
    {
      key: 'updated',
      header: <Inline gap="xs" className="whitespace-nowrap">修改时间 <Icon name="arrow-down" size="xs" className="text-slate-700" /></Inline>,
      width: '17%',
      render: (item) => <Text className="text-[13px] text-slate-500">{formatDate(item.updatedAt)}</Text>,
    },
    {
      key: 'status',
      header: '状态',
      width: '13%',
      render: (item) => <WorkbookStatusBadge item={item} />,
    },
    {
      key: 'menu',
      header: '',
      width: '56px',
      className: 'text-right',
      render: (item) => <WorkbookRowMenu item={item} {...menuProps} onOpen={onOpen} />,
    },
  ];

  return (
    <DataTable
      columns={columns}
      empty={empty}
      onRowDoubleClick={(item) => onOpen(item.unitId)}
      onSelectionChange={onSelectionChange}
      rowKey={(item) => item.unitId}
      rows={items}
      selectable
      selectedKeys={selectedKeys}
      testId="workbook-table"
    />
  );
}
