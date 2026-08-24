import React, { useMemo, useState } from 'react';
import { Box } from './layout';

export interface VirtualListProps<Item> {
  items: readonly Item[];
  itemHeight: number;
  height: number;
  itemKey: (item: Item, index: number) => string;
  renderItem: (item: Item, index: number) => React.ReactNode;
  overscan?: number;
  className?: string;
}

/** Shared fixed-row virtual list; business components provide data and row content only. */
export function VirtualList<Item>({ items, itemHeight, height, itemKey, renderItem, overscan = 6, className }: VirtualListProps<Item>): React.ReactElement {
  const [scrollTop, setScrollTop] = useState(0);
  const safeItemHeight = Math.max(1, itemHeight);
  const safeHeight = Math.max(safeItemHeight, height);
  const first = Math.max(0, Math.floor(scrollTop / safeItemHeight) - overscan);
  const last = Math.min(items.length, Math.ceil((scrollTop + safeHeight) / safeItemHeight) + overscan);
  const visible = useMemo(() => items.slice(first, last), [first, items, last]);
  return (
    <Box
      className={className}
      onScroll={(event) => setScrollTop(event.currentTarget.scrollTop)}
      style={{ height: safeHeight, overflowY: 'auto' }}
    >
      <Box className="relative" style={{ height: items.length * safeItemHeight }}>
        {visible.map((item, offset) => {
          const index = first + offset;
          return (
            <Box key={itemKey(item, index)} className="absolute inset-x-0" style={{ height: safeItemHeight, top: index * safeItemHeight }}>
              {renderItem(item, index)}
            </Box>
          );
        })}
      </Box>
    </Box>
  );
}
