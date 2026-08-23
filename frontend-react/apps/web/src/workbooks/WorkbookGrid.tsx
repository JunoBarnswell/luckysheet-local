import { Box, FileIcon, Inline, Stack, Text } from '@react-sheets/ui-system';
import type { WorkbookCatalogItem } from './types';
import { WorkbookRowMenu, type WorkbookRowMenuProps } from './WorkbookRowMenu';
import { WorkbookStatusBadge } from './WorkbookStatusBadge';

export interface WorkbookGridProps extends Omit<WorkbookRowMenuProps, 'item'> {
  items: readonly WorkbookCatalogItem[];
  selectedKeys: readonly string[];
  onSelectionChange: (keys: readonly string[]) => void;
  onOpen: (unitId: string) => void;
}

export function WorkbookGrid({ items, selectedKeys, onSelectionChange, onOpen, ...menuProps }: WorkbookGridProps) {
  const selected = new Set(selectedKeys);
  return (
    <Stack gap="md" className="grid grid-cols-1 min-[620px]:grid-cols-2 min-[1020px]:grid-cols-3 min-[1400px]:grid-cols-4">
      {items.map((item) => {
        const isSelected = selected.has(item.unitId);
        return (
          <Box
            key={item.unitId}
            aria-pressed={isSelected}
            aria-label={item.name}
            className={isSelected ? 'h-[170px] items-stretch justify-start rounded-lg border-2 border-brand bg-brand-pale p-4 text-left' : 'h-[170px] items-stretch justify-start rounded-lg border border-slate-200 bg-white p-4 text-left shadow-hub-card hover:border-brand hover:bg-brand-pale'}
            onClick={() => onSelectionChange(isSelected ? selectedKeys.filter((key) => key !== item.unitId) : [...selectedKeys, item.unitId])}
            onDoubleClick={() => onOpen(item.unitId)}
            onKeyDown={(event) => { if (event.key === 'Enter') onOpen(item.unitId); }}
            role="button"
            tabIndex={0}
          >
            <Inline gap="sm" className="w-full justify-between">
              <Inline gap="sm" className="min-w-0"><FileIcon size="md" /><Stack gap="none" className="min-w-0"><Text className="truncate text-[14px] text-slate-800" weight="semibold">{item.name}</Text><Text className="truncate text-[11px] text-slate-400">{item.locationLabel}</Text></Stack></Inline>
              <WorkbookRowMenu item={item} {...menuProps} onOpen={onOpen} />
            </Inline>
            <Stack gap="sm" className="mt-auto w-full items-start">
              <WorkbookStatusBadge item={item} />
              <Text className="text-[11px] text-slate-400">双击打开 · {new Date(item.updatedAt).toLocaleDateString('zh-CN')}</Text>
            </Stack>
          </Box>
        );
      })}
    </Stack>
  );
}
