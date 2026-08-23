import React, { useEffect, useId, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { cn } from './cn';
import { Button } from './Button';
import { Icon } from './Icon';

export interface DialogProps {
  open: boolean;
  title: string;
  description?: string;
  children: ReactNode;
  footer?: ReactNode;
  onClose: () => void;
  maxWidth?: 'sm' | 'md' | 'lg' | 'xl';
  /** Localized accessible label for the close affordance. */
  closeLabel?: string;
  /** Stable hook for browser-level interaction tests. */
  testId?: string;
  bodyClassName?: string;
}

const maxWidths = {
  sm: 'max-w-sm',
  md: 'max-w-md',
  lg: 'max-w-lg',
  xl: 'max-w-xl',
};

export function Dialog({
  open,
  title,
  description,
  children,
  footer,
  onClose,
  maxWidth = 'md',
  closeLabel = 'Close dialog',
  testId,
  bodyClassName,
}: DialogProps) {
  const dialogRef = useRef<HTMLDivElement>(null);
  const titleId = useId();

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [open, onClose]);

  if (!open || typeof document === 'undefined') return null;

  return createPortal(
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
      {/* Backdrop */}
      <div
        className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs transition-opacity animate-in fade-in duration-150"
        onClick={onClose}
        aria-hidden="true"
      />

      {/* Dialog Window */}
      <div
        ref={dialogRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        data-testid={testId}
        className={cn(
          'relative z-10 w-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-2xl animate-in zoom-in-95 duration-150',
          maxWidths[maxWidth],
        )}
      >
        {/* Header */}
        <div className="flex items-center justify-between border-b border-slate-100 px-5 py-4">
          <div>
            <h3 id={titleId} className="text-base font-semibold text-slate-900">{title}</h3>
            {description ? <p className="mt-0.5 text-xs text-slate-500">{description}</p> : null}
          </div>
          <Button
            variant="ghost"
            size="sm"
            icon="x"
            iconOnly
            onClick={onClose}
            aria-label={closeLabel}
            className="text-slate-400 hover:text-slate-700"
          />
        </div>

        {/* Body */}
        <div className={cn('max-h-[calc(85vh-130px)] overflow-y-auto px-5 py-4 text-sm text-slate-600', bodyClassName)}>
          {children}
        </div>

        {/* Footer */}
        {footer ? (
          <div className="flex items-center justify-end gap-2 border-t border-slate-100 bg-slate-50/70 px-5 py-3">
            {footer}
          </div>
        ) : null}
      </div>
    </div>,
    document.body,
  );
}
