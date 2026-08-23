import type { HTMLAttributes, ReactNode, Ref } from 'react';
import { cn } from './cn';

type BoxElement = 'article' | 'aside' | 'div' | 'footer' | 'form' | 'header' | 'li' | 'main' | 'nav' | 'section' | 'ul';

export interface BoxProps extends HTMLAttributes<HTMLElement> {
  as?: BoxElement;
  children?: ReactNode;
  /** React 19 ref-as-prop:允许业务层持有 DOM 引用 */
  ref?: Ref<HTMLElement>;
}

export function Box({ as = 'div', className, children, ref, ...props }: BoxProps) {
  const Component = as as 'div';
  return (
    <Component ref={ref as Ref<HTMLDivElement>} className={className} {...props}>
      {children}
    </Component>
  );
}

export interface StackProps extends BoxProps {
  gap?: 'none' | 'xs' | 'sm' | 'md' | 'lg' | 'xl';
}

const stackGaps: Record<NonNullable<StackProps['gap']>, string> = {
  none: 'gap-0',
  xs: 'gap-1',
  sm: 'gap-2',
  md: 'gap-4',
  lg: 'gap-6',
  xl: 'gap-8',
};

export function Stack({ gap = 'md', className, ...props }: StackProps) {
  return <Box className={cn('flex min-w-0 flex-col', stackGaps[gap], className)} {...props} />;
}

export interface InlineProps extends BoxProps {
  gap?: 'none' | 'xs' | 'sm' | 'md' | 'lg' | 'xl';
}

const inlineGaps: Record<NonNullable<InlineProps['gap']>, string> = stackGaps;

export function Inline({ gap = 'md', className, ...props }: InlineProps) {
  return <Box className={cn('flex min-w-0 flex-row items-center', inlineGaps[gap], className)} {...props} />;
}

type TextElement = 'label' | 'p' | 'small' | 'span' | 'strong';

export interface TextProps extends HTMLAttributes<HTMLElement> {
  as?: TextElement;
  children?: ReactNode;
  htmlFor?: string;
  size?: 'xs' | 'sm' | 'md' | 'lg' | 'xl';
  tone?: 'default' | 'muted' | 'subtle' | 'inverse' | 'accent' | 'success' | 'danger';
  weight?: 'normal' | 'medium' | 'semibold' | 'bold';
}

const textSizes: Record<NonNullable<TextProps['size']>, string> = {
  xs: 'text-[10px] leading-4',
  sm: 'text-xs leading-5',
  md: 'text-sm leading-6',
  lg: 'text-base leading-6',
  xl: 'text-lg leading-7',
};

const textTones: Record<NonNullable<TextProps['tone']>, string> = {
  default: 'text-ink',
  muted: 'text-muted',
  subtle: 'text-slate-400',
  inverse: 'text-white',
  accent: 'text-accent',
  success: 'text-emerald-600',
  danger: 'text-rose-600',
};

const textWeights: Record<NonNullable<TextProps['weight']>, string> = {
  normal: 'font-normal',
  medium: 'font-medium',
  semibold: 'font-semibold',
  bold: 'font-bold',
};

export function Text({ as = 'span', size = 'md', tone = 'default', weight = 'normal', className, children, ...props }: TextProps) {
  const Component = as;
  return (
    <Component className={cn('min-w-0', textSizes[size], textTones[tone], textWeights[weight], className)} {...props}>
      {children}
    </Component>
  );
}

type HeadingElement = 'h1' | 'h2' | 'h3';

export interface HeadingProps extends HTMLAttributes<HTMLHeadingElement> {
  as?: HeadingElement;
  children?: ReactNode;
  size?: 'sm' | 'md' | 'lg' | 'xl';
}

const headingSizes: Record<NonNullable<HeadingProps['size']>, string> = {
  sm: 'text-sm leading-5',
  md: 'text-base leading-6',
  lg: 'text-lg leading-7',
  xl: 'text-xl leading-8',
};

export function Heading({ as = 'h2', size = 'md', className, children, ...props }: HeadingProps) {
  const Component = as;
  return (
    <Component className={cn('font-semibold tracking-[-0.01em] text-ink', headingSizes[size], className)} {...props}>
      {children}
    </Component>
  );
}

export interface DividerProps extends HTMLAttributes<HTMLHRElement> {
  orientation?: 'horizontal' | 'vertical';
}

export function Divider({ orientation = 'horizontal', className, ...props }: DividerProps) {
  return (
    <hr
      aria-orientation={orientation}
      className={cn(orientation === 'horizontal' ? 'h-px w-full border-0 bg-line' : 'h-5 w-px border-0 bg-line', className)}
      {...props}
    />
  );
}

export function Kbd({ children, className, ...props }: HTMLAttributes<HTMLElement> & { children?: ReactNode }) {
  return (
    <kbd
      className={cn('inline-flex min-h-5 items-center rounded border border-line bg-slate-50 px-1.5 font-mono text-[10px] font-medium text-muted', className)}
      {...props}
    >
      {children}
    </kbd>
  );
}

export function ScrollArea({ className, ...props }: BoxProps) {
  return <Box className={cn('min-h-0 overflow-auto', className)} {...props} />;
}
