import { forwardRef, type InputHTMLAttributes, type ReactNode } from 'react';
import { cn } from './cn';
import { Icon, type IconName } from './Icon';

export interface TextInputProps extends InputHTMLAttributes<HTMLInputElement> {
  /** Layout classes for the shared input wrapper. */
  containerClassName?: string;
  error?: boolean;
  leadingIcon?: IconName;
  trailing?: ReactNode;
}

export const TextInput = forwardRef<HTMLInputElement, TextInputProps>(function TextInput(
  { className, containerClassName, error = false, leadingIcon, trailing, ...props },
  ref,
) {
  return (
    <span className={cn('relative flex min-w-0 items-center', containerClassName)}>
      {leadingIcon ? <Icon name={leadingIcon} size="sm" className="pointer-events-none absolute left-3 text-slate-400" /> : null}
      <input
        ref={ref}
        aria-invalid={error || undefined}
        className={cn(
          'min-h-9 w-full min-w-0 rounded-lg border bg-white px-3 text-sm text-ink outline-none transition placeholder:text-slate-400 focus:border-accent focus:ring-2 focus:ring-accent/15 disabled:cursor-not-allowed disabled:bg-slate-50 disabled:text-slate-400 read-only:bg-slate-50/70',
          error ? 'border-rose-300 focus:border-rose-500 focus:ring-rose-100' : 'border-line',
          leadingIcon && 'pl-9',
          trailing ? 'pr-9' : undefined,
          className,
        )}
        {...props}
      />
      {trailing ? <span className="absolute right-2.5 inline-flex items-center">{trailing}</span> : null}
    </span>
  );
});
