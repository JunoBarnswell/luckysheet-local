import React, { createContext, useContext, useEffect, useId, useLayoutEffect, useRef, useState, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { cn } from './cn';

export interface DropdownMenuProps {
  trigger: ReactNode;
  children: ReactNode | ((helpers: { close: () => void }) => ReactNode);
  align?: 'left' | 'right';
  className?: string;
  disabled?: boolean;
}

const MenuOwnersContext = createContext('');

export function DropdownMenu({ trigger, children, align = 'left', className, disabled = false }: DropdownMenuProps) {
  const menuId = useId();
  const owners = `${useContext(MenuOwnersContext)} ${menuId}`.trim();
  const [open, setOpen] = useState(false);
  const triggerRef = useRef<HTMLDivElement>(null);
  const menuRef = useRef<HTMLDivElement>(null);
  const [coords, setCoords] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  const updatePosition = () => {
    if (!triggerRef.current) return;
    const rect = triggerRef.current.getBoundingClientRect();
    const menu = menuRef.current?.getBoundingClientRect();
    const width = menu?.width ?? 0;
    const height = menu?.height ?? 0;
    const availableHeight = window.innerHeight - 16;
    setCoords({
      x: Math.max(8, Math.min(align === 'right' ? rect.right - width : rect.left, window.innerWidth - width - 8)),
      y: Math.max(8, Math.min(rect.bottom + 4, window.innerHeight - Math.min(height, availableHeight) - 8)),
    });
  };

  useLayoutEffect(() => { if (open) updatePosition(); }, [open, align]);

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
      const target = e.target instanceof Element ? e.target : null;
      if (target?.closest('[data-menu-owners]')?.getAttribute('data-menu-owners')?.split(' ').includes(menuId)) return;
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
    <MenuOwnersContext.Provider value={owners}>
    <div className="relative inline-flex" ref={triggerRef}>
      <div onClick={toggle} aria-disabled={disabled || undefined} className={disabled ? 'cursor-not-allowed' : 'cursor-pointer'}>
        {trigger}
      </div>
      {open && typeof document !== 'undefined'
        ? createPortal(
            <div
              ref={menuRef}
              data-menu-owners={owners}
              className={cn(
                'fixed z-50 rounded-lg border border-slate-200 bg-white p-1 shadow-xl animate-in fade-in zoom-in-95 duration-100',
                className,
              )}
              style={{ left: coords.x, top: coords.y, maxWidth: 'calc(100vw - 16px)', maxHeight: 'calc(100vh - 16px)', overflowY: 'auto' }}
            >
              {typeof children === 'function' ? children({ close: () => setOpen(false) }) : children}
            </div>,
            document.body,
          )
        : null}
    </div>
    </MenuOwnersContext.Provider>
  );
}
