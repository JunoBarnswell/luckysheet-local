import { forwardRef, type InputHTMLAttributes } from 'react';
import { cn } from './cn';
import { Icon } from './Icon';

export interface CheckToggleProps extends Omit<InputHTMLAttributes<HTMLInputElement>, 'type'> {
  label?: string;
}

export const CheckToggle = forwardRef<HTMLInputElement, CheckToggleProps>(function CheckToggle(
  { checked, className, disabled, label, ...props },
  ref,
) {
  return (
    <label className={cn('inline-flex min-w-0 items-center gap-2 text-xs text-ink', disabled && 'cursor-not-allowed opacity-50', className)}>
      <span className="relative inline-flex h-4 w-4 shrink-0 items-center justify-center rounded border border-slate-300 bg-white transition-colors has-[:focus-visible]:ring-2 has-[:focus-visible]:ring-accent/30 has-[:checked]:border-accent has-[:checked]:bg-accent">
        <input ref={ref} {...props} checked={checked} className="peer absolute inset-0 cursor-pointer opacity-0 disabled:cursor-not-allowed" disabled={disabled} type="checkbox" />
        {checked ? <Icon name="check" size="xs" className="text-white" /> : null}
      </span>
      {label ? <span className="truncate">{label}</span> : null}
    </label>
  );
});
