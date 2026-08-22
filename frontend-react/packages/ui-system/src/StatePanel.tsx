import type { ReactNode } from 'react';
import { Button } from './Button';
import { cn } from './cn';
import { Icon, type IconName } from './Icon';
import { Panel, PanelDescription, PanelTitle } from './Panel';
import { Stack } from './layout';

export type StatePanelKind = 'disabled' | 'empty' | 'error' | 'loading';

export interface StatePanelProps {
  actionDisabled?: boolean;
  actionLabel?: string;
  className?: string;
  description?: string;
  icon?: IconName;
  kind: StatePanelKind;
  onAction?: () => void;
  title?: string;
}

const stateDefaults: Record<StatePanelKind, { icon: IconName; title: string; tone: 'accent' | 'danger' | 'muted' }> = {
  loading: { icon: 'loader', title: 'Loading workspace', tone: 'accent' },
  error: { icon: 'alert-circle', title: 'Something went wrong', tone: 'danger' },
  empty: { icon: 'file-plus', title: 'Nothing here yet', tone: 'muted' },
  disabled: { icon: 'lock', title: 'Feature unavailable', tone: 'muted' },
};

const stateIconTones: Record<NonNullable<(typeof stateDefaults)[StatePanelKind]['tone']>, string> = {
  accent: 'bg-blue-100 text-accent',
  danger: 'bg-rose-100 text-rose-600',
  muted: 'bg-slate-200 text-muted',
};

export function StatePanel({ actionDisabled = false, actionLabel, className, description, icon, kind, onAction, title }: StatePanelProps) {
  const defaults = stateDefaults[kind];
  const resolvedIcon = icon ?? defaults.icon;
  return (
    <Panel tone="subtle" className={cn('flex min-h-40 items-center justify-center px-6 py-8 text-center shadow-none', className)}>
      <Stack gap="sm" className="max-w-sm items-center">
        <span className={cn('inline-flex h-10 w-10 items-center justify-center rounded-xl', stateIconTones[defaults.tone])}>
          <Icon name={resolvedIcon} size="lg" className={kind === 'loading' ? 'animate-spin' : undefined} />
        </span>
        <PanelTitle as="h3" size="sm">{title ?? defaults.title}</PanelTitle>
        {description ? <PanelDescription className="max-w-xs">{description}</PanelDescription> : null}
        {onAction && actionLabel ? (
          <Button disabled={actionDisabled} onClick={onAction} size="sm" variant={kind === 'error' ? 'soft' : 'secondary'}>
            {actionLabel}
          </Button>
        ) : null}
      </Stack>
    </Panel>
  );
}
