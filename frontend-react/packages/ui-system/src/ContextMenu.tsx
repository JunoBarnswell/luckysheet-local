import React, { useEffect, useRef } from 'react';
import { createPortal } from 'react-dom';
import { cn } from './cn';
import { Icon, type IconName } from './Icon';

export interface ContextMenuActionItem {
  id: string;
  label: string;
  icon?: IconName;
  shortcut?: string;
  disabled?: boolean;
  danger?: boolean;
  separator?: false;
  onSelect: () => void;
}

export interface ContextMenuSeparatorItem {
  id: string;
  label?: string;
  separator: true;
}

export type ContextMenuItem = ContextMenuActionItem | ContextMenuSeparatorItem;

export interface ContextMenuProps {
  x: number;
  y: number;
  open: boolean;
  items: ContextMenuItem[];
  onClose: () => void;
}

export function ContextMenu({ x, y, open, items, onClose }: ContextMenuProps) {
  const menuRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    const handleOutsideClick = (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose();
      }
    };
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };

    document.addEventListener('mousedown', handleOutsideClick);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handleOutsideClick);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [open, onClose]);

  if (!open || typeof document === 'undefined') return null;

  // Adjust coordinates if overflowing screen
  const screenW = window.innerWidth;
  const screenH = window.innerHeight;
  const menuW = 200;
  const menuH = items.length * 32;
  const posX = x + menuW > screenW ? Math.max(10, screenW - menuW - 10) : x;
  const posY = y + menuH > screenH ? Math.max(10, screenH - menuH - 10) : y;

  return createPortal(
    <div
      ref={menuRef}
      role="menu"
      aria-orientation="vertical"
      className="fixed z-50 min-w-[190px] overflow-hidden rounded-lg border border-slate-200 bg-white p-1 text-slate-700 shadow-xl animate-in fade-in zoom-in-95 duration-100"
      style={{ left: posX, top: posY }}
    >
      {items.map((item, index) => {
        if (item.separator) {
          return <div key={`sep-${index}`} className="my-1 h-px bg-slate-200" role="separator" />;
        }

        return (
          <button
            key={item.id}
            role="menuitem"
            disabled={item.disabled}
            onClick={() => {
              item.onSelect();
              onClose();
            }}
            className={cn(
              'flex w-full items-center justify-between gap-3 rounded-md px-2.5 py-1.5 text-left text-xs font-medium transition-colors',
              item.danger
                ? 'text-rose-600 hover:bg-rose-50'
                : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900',
              item.disabled && 'pointer-events-none opacity-40',
            )}
          >
            <span className="flex items-center gap-2">
              {item.icon ? <Icon name={item.icon} size="sm" className="opacity-75" /> : <span className="w-3.5" />}
              {item.label}
            </span>
            {item.shortcut ? <span className="text-[10px] text-slate-400">{item.shortcut}</span> : null}
          </button>
        );
      })}
    </div>,
    document.body,
  );
}
