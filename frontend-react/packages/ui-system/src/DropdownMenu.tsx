import React, { useEffect, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { cn } from './cn';

export interface DropdownMenuProps {
  trigger: ReactNode;
  children: ReactNode | ((helpers: { close: () => void }) => ReactNode);
  align?: 'left' | 'right';
  className?: string;
  disabled?: boolean;
}

export function DropdownMenu({ trigger, children, align = 'left', className, disabled = false }: DropdownMenuProps) {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [coords, setCoords] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  const updatePosition = () => {
    if (!triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const menuRect = menuRef.current?.getBoundingClientRect();
    const menuWidth = menuRect?.width ?? 0;
    const menuHeight = menuRect?.height ?? 0;
    const viewportMargin = 8;
    const preferredX = align === 'right' ? rect.right - menuWidth : rect.left;
    const maxX = Math.max(viewportMargin, window.innerWidth - menuWidth - viewportMargin);
    const belowY = rect.bottom + 4;
    const preferredY = menuHeight > 0 && belowY + menuHeight > window.innerHeight - viewportMargin
      ? rect.top - menuHeight - 4
      : belowY;
    const maxY = Math.max(viewportMargin, window.innerHeight - menuHeight - viewportMargin);
    setCoords({
      x: Math.min(Math.max(viewportMargin, preferredX), maxX),
      y: Math.min(Math.max(viewportMargin, preferredY), maxY),
    });
  };

  const toggle = () => {
    if (disabled) return;
    updatePosition();
    setOpen((prev) => !prev);
  };

  useEffect(() => {
    if (disabled) setOpen(false);
  }, [disabled]);

  useEffect(() => {
    if (!open) return;
    const handleClickOutside = (e: MouseEvent) => {
      if (
        triggerRef.current &&
        !triggerRef.current.contains(e.target as Node) &&
        menuRef.current &&
        !menuRef.current.contains(e.target as Node)
      ) {
        setOpen(false);
      }
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };

    document.addEventListener('mousedown', handleClickOutside);
    document.addEventListener('keydown', handleKeyDown);
    window.addEventListener('resize', updatePosition);
    window.addEventListener('scroll', updatePosition, true);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('resize', updatePosition);
      window.removeEventListener('scroll', updatePosition, true);
    };
  }, [open]);

  useLayoutEffect(() => {
    if (open) updatePosition();
  }, [open, align]);

  return (
    <div className="relative inline-flex" ref={triggerRef}>
      <div onClick={toggle} aria-disabled={disabled || undefined} className={disabled ? 'cursor-not-allowed' : 'cursor-pointer'}>
        {trigger}
      </div>
      {open && typeof document !== 'undefined'
        ? createPortal(
            <div
              ref={menuRef}
              className={cn(
                'fixed z-50 max-h-[calc(100vh-1rem)] overflow-y-auto rounded-lg border border-slate-200 bg-white p-1 shadow-xl animate-in fade-in zoom-in-95 duration-100',
                className,
              )}
              style={{ left: coords.x, top: coords.y }}
            >
              {typeof children === 'function' ? children({ close: () => setOpen(false) }) : children}
            </div>,
            document.body,
          )
        : null}
    </div>
  );
}
