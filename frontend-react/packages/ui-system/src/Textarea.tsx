import React, { forwardRef } from 'react';
import { cn } from './cn';

export interface TextareaProps extends Omit<React.TextareaHTMLAttributes<HTMLTextAreaElement>, 'className'> {
  className?: string;
}

export const Textarea = forwardRef<HTMLTextAreaElement, TextareaProps>(function Textarea(
  { className, ...props },
  ref,
): React.ReactElement {
  return (
    <textarea
      {...props}
      ref={ref}
      className={cn(
        'w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-700',
        'outline-none transition-colors focus:border-blue-500 focus:ring-2 focus:ring-blue-100',
        className,
      )}
    />
  );
});
