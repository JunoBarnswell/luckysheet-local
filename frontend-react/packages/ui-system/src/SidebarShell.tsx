import { useCallback, useEffect, useRef, useState, type ReactNode } from 'react';
import { Button } from './Button';
import { Box, Heading, Inline, Stack, Text } from './layout';

const DEFAULT_WIDTH = 340;
const MIN_WIDTH = 320;
const MAX_WIDTH = 360;

export interface SidebarShellProps {
  children?: ReactNode;
  minWidth?: number;
  maxWidth?: number;
  onOpenChange?: (open: boolean) => void;
  onWidthChange?: (width: number) => void;
  open: boolean;
  title?: string;
  width?: number;
  showHeader?: boolean;
  contentOverflow?: 'auto' | 'hidden';
}

export function SidebarShell({
  children,
  minWidth = MIN_WIDTH,
  maxWidth = MAX_WIDTH,
  onOpenChange,
  onWidthChange,
  open,
  title,
  width = DEFAULT_WIDTH,
  showHeader = true,
  contentOverflow = 'auto',
}: SidebarShellProps) {
  const [localWidth, setLocalWidth] = useState(width);
  const dragging = useRef(false);

  useEffect(() => {
    setLocalWidth(width);
  }, [width]);

  const clampWidth = useCallback((next: number) => Math.min(maxWidth, Math.max(minWidth, next)), [maxWidth, minWidth]);

  const handleResizePointerDown = (event: React.PointerEvent<HTMLButtonElement>) => {
    event.preventDefault();
    dragging.current = true;
    event.currentTarget.setPointerCapture(event.pointerId);
  };

  const handleResizePointerMove = (event: React.PointerEvent<HTMLButtonElement>) => {
    if (!dragging.current) return;
    const shell = event.currentTarget.closest('[data-sidebar-shell]') as HTMLElement | null;
    if (!shell) return;
    const rect = shell.getBoundingClientRect();
    const next = clampWidth(rect.right - event.clientX);
    setLocalWidth(next);
    onWidthChange?.(next);
  };

  const handleResizePointerUp = (event: React.PointerEvent<HTMLButtonElement>) => {
    dragging.current = false;
    event.currentTarget.releasePointerCapture(event.pointerId);
  };

  if (!open) return null;

  return (
    <Box
      as="aside"
      aria-label={title ?? 'Context sidebar'}
      className="relative flex h-full min-h-0 shrink-0 flex-col overflow-hidden border-l border-slate-200 bg-slate-50/70"
      data-sidebar-shell
      data-testid="sidebar-shell"
      style={{ width: localWidth }}
    >
      <Button
        aria-label="Resize sidebar"
        className="absolute inset-y-0 -left-1 z-10 w-2 cursor-col-resize rounded-none border-0 bg-transparent p-0 hover:bg-blue-500/20"
        onPointerDown={handleResizePointerDown}
        onPointerMove={handleResizePointerMove}
        onPointerUp={handleResizePointerUp}
        type="button"
        variant="ghost"
      />
      <Stack gap="none" className="min-h-0 flex-1">
        {showHeader ? <Inline gap="sm" className="shrink-0 items-center justify-between border-b border-slate-200 bg-white px-3 py-2">
          <Heading as="h2" size="sm">{title ?? 'Properties'}</Heading>
          <Button
            aria-label="Close sidebar"
            data-testid="sidebar-close"
            icon="x"
            iconOnly
            onClick={() => onOpenChange?.(false)}
            size="sm"
            variant="ghost"
          />
        </Inline> : null}
        <Box className={`min-h-0 flex-1 overflow-${contentOverflow}`}>{children}</Box>
      </Stack>
    </Box>
  );
}

export function SidebarPlaceholder({ description }: { description: string }) {
  return (
    <Box className="p-4">
      <Text size="sm" tone="muted">{description}</Text>
    </Box>
  );
}
