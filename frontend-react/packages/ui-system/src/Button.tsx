import { forwardRef, type ButtonHTMLAttributes, type ReactNode } from 'react';
import { cn } from './cn';
import { Icon, type IconName } from './Icon';

export type ButtonVariant = 'danger' | 'ghost' | 'outline' | 'primary' | 'secondary' | 'soft';
export type ButtonSize = 'xs' | 'sm' | 'md' | 'lg';

export interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  children?: ReactNode;
  icon?: IconName;
  iconOnly?: boolean;
  loading?: boolean;
  size?: ButtonSize;
  variant?: ButtonVariant;
}

const variants: Record<ButtonVariant, string> = {
  primary: 'bg-accent text-white shadow-sm shadow-accent/20 hover:bg-blue-700',
  secondary: 'border border-slate-200 bg-white text-ink shadow-sm hover:border-accent/40 hover:bg-blue-50/50',
  outline: 'border border-line bg-transparent text-ink hover:border-accent/50 hover:bg-blue-50/40',
  ghost: 'bg-transparent text-muted hover:bg-slate-100 hover:text-ink',
  soft: 'bg-blue-50 text-accent hover:bg-blue-100',
  danger: 'bg-rose-50 text-rose-600 hover:bg-rose-100',
};

const sizes: Record<ButtonSize, string> = {
  xs: 'min-h-7 rounded-md px-2 text-[11px]',
  sm: 'min-h-8 rounded-md px-2.5 text-xs',
  md: 'min-h-9 rounded-lg px-3 text-sm',
  lg: 'min-h-11 rounded-xl px-4 text-sm',
};

export const Button = forwardRef<HTMLButtonElement, ButtonProps>(function Button(
  { children, className, disabled, icon, iconOnly = false, loading = false, size = 'md', type = 'button', variant = 'secondary', ...props },
  ref,
) {
  return (
    <button
      ref={ref}
      aria-busy={loading || undefined}
      className={cn(
        'inline-flex shrink-0 items-center justify-center gap-2 whitespace-nowrap font-medium transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-accent/50 focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-45',
        sizes[size],
        variants[variant],
        iconOnly && 'w-9 px-0',
        className,
      )}
      disabled={disabled || loading}
      type={type}
      {...props}
    >
      {loading ? <Icon name="loader" size="sm" className="animate-spin" /> : icon ? <Icon name={icon} size="sm" /> : null}
      {children}
    </button>
  );
});
