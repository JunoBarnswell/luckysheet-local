import type { HTMLAttributes, ReactNode } from 'react';
import { cn } from './cn';
import { Box, Heading, Text, type BoxProps, type HeadingProps, type TextProps } from './layout';

export type PanelTone = 'accent' | 'default' | 'dashed' | 'dark' | 'subtle';

export interface PanelProps extends BoxProps {
  children?: ReactNode;
  tone?: PanelTone;
}

const panelTones: Record<PanelTone, string> = {
  default: 'border border-line bg-white shadow-panel',
  subtle: 'border border-line/80 bg-slate-50/80',
  accent: 'border border-blue-100 bg-blue-50/70',
  dark: 'border border-slate-800 bg-slate-900 text-white',
  dashed: 'border border-dashed border-slate-300 bg-slate-50/60',
};

export function Panel({ className, tone = 'default', ...props }: PanelProps) {
  return <Box as="section" className={cn('min-w-0 rounded-xl', panelTones[tone], className)} {...props} />;
}

export function PanelHeader({ className, ...props }: BoxProps) {
  return <Box className={cn('flex min-w-0 items-center justify-between gap-3 border-b border-line/80 px-4 py-3', className)} {...props} />;
}

export function PanelBody({ className, ...props }: BoxProps) {
  return <Box className={cn('min-w-0 p-4', className)} {...props} />;
}

export function PanelTitle({ className, size = 'md', ...props }: HeadingProps) {
  return <Heading className={cn('truncate', className)} size={size} {...props} />;
}

export function PanelDescription({ className, size = 'sm', tone = 'muted', ...props }: TextProps) {
  return <Text className={cn('leading-5', className)} size={size} tone={tone} {...props} />;
}

export interface PanelFooterProps extends HTMLAttributes<HTMLElement> {
  children?: ReactNode;
}

export function PanelFooter({ className, children, ...props }: PanelFooterProps) {
  return (
    <footer className={cn('flex items-center justify-between gap-3 border-t border-line/80 px-4 py-3', className)} {...props}>
      {children}
    </footer>
  );
}
