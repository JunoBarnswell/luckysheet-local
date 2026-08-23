import type { ReactNode } from 'react';
import { cn } from './cn';

export interface DataTableColumn<Row> {
  key: string;
  header: ReactNode;
  className?: string;
  headerClassName?: string;
  width?: string;
  render: (row: Row, index: number) => ReactNode;
}

export interface DataTableProps<Row> {
  columns: readonly DataTableColumn<Row>[];
  rows: readonly Row[];
  rowKey: (row: Row) => string;
  className?: string;
  empty?: ReactNode;
  selectable?: boolean;
  selectedKeys?: readonly string[];
  onSelectionChange?: (keys: readonly string[]) => void;
  onRowClick?: (row: Row) => void;
  onRowDoubleClick?: (row: Row) => void;
  testId?: string;
}

/**
 * The shared table primitive owns table markup and selection semantics. Business
 * pages only provide typed rows and renderers, so their layout never needs to
 * scatter native table elements or recreate checkbox behavior.
 */
export function DataTable<Row>({
  columns,
  rows,
  rowKey,
  className,
  empty,
  selectable = false,
  selectedKeys = [],
  onSelectionChange,
  onRowClick,
  onRowDoubleClick,
  testId,
}: DataTableProps<Row>) {
  const selected = new Set(selectedKeys);
  const allSelected = selectable && rows.length > 0 && rows.every((row) => selected.has(rowKey(row)));

  const setRowSelected = (row: Row, checked: boolean) => {
    if (!onSelectionChange) return;
    const key = rowKey(row);
    const next = new Set(selected);
    if (checked) next.add(key);
    else next.delete(key);
    onSelectionChange([...next]);
  };

  const setAllSelected = (checked: boolean) => {
    if (!onSelectionChange) return;
    onSelectionChange(checked ? rows.map(rowKey) : []);
  };

  return (
    <div className={cn('min-w-0 overflow-x-auto', className)} data-testid={testId}>
      <table className="w-full min-w-[760px] border-separate border-spacing-0 text-left">
        <thead>
          <tr className="h-11 border-b border-brand-line bg-white text-[12px] font-medium text-slate-500">
            {selectable ? (
              <th className="w-12 border-b border-brand-line px-4" scope="col">
                <input
                  aria-label="全选"
                  checked={allSelected}
                  className="h-4 w-4 accent-brand"
                  disabled={rows.length === 0 || !onSelectionChange}
                  onChange={(event) => setAllSelected(event.target.checked)}
                  type="checkbox"
                />
              </th>
            ) : null}
            {columns.map((column) => (
              <th
                key={column.key}
                className={cn('border-b border-brand-line px-4 py-2.5', column.headerClassName)}
                scope="col"
                style={{ width: column.width }}
              >
                {column.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 ? (
            <tr>
              <td className="px-4 py-10" colSpan={columns.length + (selectable ? 1 : 0)}>
                {empty}
              </td>
            </tr>
          ) : (
            rows.map((row, index) => {
              const key = rowKey(row);
              const isSelected = selected.has(key);
              return (
                <tr
                  key={key}
                  aria-selected={isSelected}
                  className={cn(
                    'group h-[45px] transition-colors hover:bg-brand-pale',
                    isSelected && 'bg-brand-soft',
                    onRowClick && 'cursor-pointer',
                  )}
                  onClick={() => onRowClick?.(row)}
                  onDoubleClick={() => onRowDoubleClick?.(row)}
                >
                  {selectable ? (
                    <td className="border-b border-slate-100 px-4" onClick={(event) => event.stopPropagation()}>
                      <input
                        aria-label={`选择 ${key}`}
                        checked={isSelected}
                        className="h-4 w-4 accent-brand"
                        disabled={!onSelectionChange}
                        onChange={(event) => setRowSelected(row, event.target.checked)}
                        type="checkbox"
                      />
                    </td>
                  ) : null}
                  {columns.map((column) => (
                    <td key={column.key} className={cn('border-b border-slate-100 px-4 py-2.5', column.className)}>
                      {column.render(row, index)}
                    </td>
                  ))}
                </tr>
              );
            })
          )}
        </tbody>
      </table>
    </div>
  );
}
