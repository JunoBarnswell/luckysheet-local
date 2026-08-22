import { forwardRef, type ButtonHTMLAttributes, type ReactNode } from 'react';
import { Button } from './Button';
import { cn } from './cn';
import { Box, type BoxProps } from './layout';
import { Icon, type IconName } from './Icon';

export function Tabs({ className, ...props }: BoxProps) {
  return <Box className={cn('min-w-0', className)} {...props} />;
}

export interface TabListProps extends BoxProps {
  label: string;
}

export function TabList({ className, label, ...props }: TabListProps) {
  return <Box aria-label={label} className={cn('flex min-w-0 items-center gap-1', className)} role="tablist" {...props} />;
}

export interface TabProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  active?: boolean;
  children?: ReactNode;
  icon?: IconName;
}

export const Tab = forwardRef<HTMLButtonElement, TabProps>(function Tab({ active = false, children, className, disabled, icon, ...props }, ref) {
  return (
    <Button
      ref={ref}
      aria-selected={active}
      className={cn(
        'relative rounded-lg px-3 py-1.5 text-xs focus-visible:ring-offset-0',
        active ? 'bg-blue-50 font-semibold text-accent' : 'text-muted hover:bg-slate-100 hover:text-ink',
        className,
      )}
      disabled={disabled}
      icon={icon}
      role="tab"
      size="sm"
      tabIndex={0}
      variant="ghost"
      {...props}
    >
      {children}
    </Button>
  );
});

export interface TabPanelProps extends BoxProps {
  active?: boolean;
  labelledBy?: string;
}

export function TabPanel({ active = true, labelledBy, className, ...props }: TabPanelProps) {
  return (
    <Box
      aria-labelledby={labelledBy}
      className={cn('min-w-0', className)}
      hidden={!active}
      role="tabpanel"
      tabIndex={0}
      {...props}
    />
  );
}

export function TabDivider() {
  return <Icon name="chevron-right" size="xs" className="mx-0.5 text-slate-300" />;
}
