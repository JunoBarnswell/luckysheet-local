import React, { useEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { cn } from './cn';

export interface DropdownMenuProps {
  trigger: ReactNode;
  children: ReactNode | ((helpers: { close: () => void }) => ReactNode);
  align?: 'left' | 'right';
  className?: string;
}

export function DropdownMenu({ trigger, children, align = 'left', className }: DropdownMenuProps) {
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [coords, setCoords] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  const updatePosition = () => {
    if (!triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    setCoords({
      x: align === 'right' ? rect.right : rect.left,
      y: rect.bottom + 4,
    });
  };

  const toggle = () => {
    updatePosition();
    setOpen((prev) => !prev);
  };

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
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
      document.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('resize', updatePosition);
    };
  }, [open]);

  return (
    <div className="relative inline-flex" ref={triggerRef}>
      <div onClick={toggle} className="cursor-pointer">
        {trigger}
      </div>
      {open && typeof document !== 'undefined'
        ? createPortal(
            <div
              ref={menuRef}
              className={cn(
                'fixed z-50 rounded-lg border border-slate-200 bg-white p-1 shadow-xl animate-in fade-in zoom-in-95 duration-100',
                align === 'right' && '-translate-x-full',
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
