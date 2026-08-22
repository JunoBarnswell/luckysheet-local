import React from 'react';
import { cn } from './cn';

export interface TextareaProps extends Omit<React.TextareaHTMLAttributes<HTMLTextAreaElement>, 'className'> {
  className?: string;
}

export function Textarea({ className, ...props }: TextareaProps): React.ReactElement {
  return (
    <textarea
      {...props}
      className={cn(
        'w-full rounded-md border border-slate-200 bg-white px-2.5 py-1.5 text-xs text-slate-700',
        'outline-none transition-colors focus:border-blue-500 focus:ring-2 focus:ring-blue-100',
        className,
      )}
    />
  );
}
