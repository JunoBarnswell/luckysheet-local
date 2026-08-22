import React, { forwardRef, type SelectHTMLAttributes } from 'react';
import { cn } from './cn';
import { Icon } from './Icon';

export interface SelectProps extends SelectHTMLAttributes<HTMLSelectElement> {
  error?: boolean;
  sizeVariant?: 'sm' | 'md' | 'lg';
}

const selectSizes = {
  sm: 'h-8 text-xs px-2.5 pr-7',
  md: 'h-9 text-sm px-3 pr-8',
  lg: 'h-10 text-base px-3.5 pr-9',
};

export const Select = forwardRef<HTMLSelectElement, SelectProps>(function Select(
  { className, error, sizeVariant = 'md', children, ...props },
  ref,
) {
  return (
    <div className="relative inline-flex w-full items-center">
      <select
        ref={ref}
        className={cn(
          'w-full appearance-none rounded-lg border bg-white font-medium transition-colors focus:outline-hidden',
          error
            ? 'border-rose-400 text-rose-900 focus:border-rose-500 focus:ring-2 focus:ring-rose-200'
            : 'border-slate-200 text-slate-800 hover:border-slate-300 focus:border-blue-500 focus:ring-2 focus:ring-blue-100',
          selectSizes[sizeVariant],
          className,
        )}
        {...props}
      >
        {children}
      </select>
      <Icon
        name="chevron-down"
        size="sm"
        className="pointer-events-none absolute right-2.5 text-slate-400"
      />
    </div>
  );
});
