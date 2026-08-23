import type { DragEvent, ReactNode } from 'react';
import { cn } from './cn';
import { Box, Text, type BoxProps } from './layout';

export interface FieldDropZoneProps<T extends { id: string }> extends Omit<BoxProps, 'children'> {
  emptyLabel: string;
  items: readonly T[];
  renderItem: (item: T, index: number) => ReactNode;
  onDropItem: (event: DragEvent<HTMLElement>, index?: number) => void;
  onDragOverItem?: (event: DragEvent<HTMLElement>) => void;
  disabled?: boolean;
}

export function FieldDropZone<T extends { id: string }>({
  className,
  disabled = false,
  emptyLabel,
  items,
  onDragOverItem,
  onDropItem,
  renderItem,
  ...props
}: FieldDropZoneProps<T>) {
  return (
    <Box
      {...props}
      aria-disabled={disabled || undefined}
      className={cn(
        'min-h-12 rounded-lg border border-dashed border-slate-300 bg-slate-50/70 p-2 transition-colors',
        !disabled && 'focus-within:border-accent/60 has-[:hover]:border-accent/60 has-[:hover]:bg-blue-50/40',
        disabled && 'cursor-not-allowed opacity-55',
        className,
      )}
      onDragOver={(event) => {
        if (!disabled) {
          event.preventDefault();
          onDragOverItem?.(event);
        }
      }}
      onDrop={(event) => {
        if (!disabled) onDropItem(event);
      }}
      role="list"
    >
      {items.length === 0 ? <Text size="xs" tone="subtle" className="block px-1 py-1.5">{emptyLabel}</Text> : null}
      {items.map((item, index) => (
        <Box
          key={item.id}
          className="mb-1 last:mb-0"
          onDragOver={(event) => {
            if (!disabled) {
              event.preventDefault();
              event.stopPropagation();
            }
          }}
          onDrop={(event) => {
            event.stopPropagation();
            if (!disabled) onDropItem(event, index);
          }}
          role="listitem"
        >
          {renderItem(item, index)}
        </Box>
      ))}
    </Box>
  );
}
