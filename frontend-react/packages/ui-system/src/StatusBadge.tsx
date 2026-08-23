import { cn } from './cn';
import { Icon, type IconName } from './Icon';

export type StatusBadgeKind = 'synced' | 'local' | 'shared' | 'syncing' | 'pending' | 'offline' | 'conflict' | 'error' | 'trashed' | 'viewer' | 'commenter' | 'editor' | 'owner';

export interface StatusBadgeProps {
  kind: StatusBadgeKind;
  label?: string;
  compact?: boolean;
  className?: string;
}

const statusConfig: Record<StatusBadgeKind, { label: string; icon: IconName; className: string }> = {
  synced: { label: '已同步', icon: 'check-circle', className: 'bg-emerald-50 text-emerald-700' },
  local: { label: '本地文件', icon: 'users', className: 'bg-blue-50 text-blue-700' },
  shared: { label: '共享', icon: 'users', className: 'bg-violet-50 text-violet-700' },
  syncing: { label: '正在同步', icon: 'refresh', className: 'bg-amber-50 text-amber-700' },
  pending: { label: '待同步', icon: 'clock', className: 'bg-amber-50 text-amber-700' },
  offline: { label: '离线', icon: 'cloud-check', className: 'bg-slate-100 text-slate-600' },
  conflict: { label: '存在冲突', icon: 'alert-circle', className: 'bg-rose-50 text-rose-700' },
  error: { label: '同步失败', icon: 'alert-circle', className: 'bg-rose-50 text-rose-700' },
  trashed: { label: '回收站', icon: 'trash', className: 'bg-slate-100 text-slate-600' },
  owner: { label: '所有者', icon: 'users', className: 'bg-brand-soft text-brand-dark' },
  editor: { label: '可编辑', icon: 'pencil', className: 'bg-blue-50 text-blue-700' },
  commenter: { label: '可评论', icon: 'comment', className: 'bg-amber-50 text-amber-700' },
  viewer: { label: '只读', icon: 'eye', className: 'bg-slate-100 text-slate-600' },
};

export function StatusBadge({ kind, label, compact = false, className }: StatusBadgeProps) {
  const config = statusConfig[kind];
  return (
    <span className={cn('inline-flex max-w-full items-center gap-1.5 rounded-full px-2.5 py-1 text-[11px] font-medium leading-4', compact && 'px-2 py-0.5', config.className, className)}>
      <Icon name={config.icon} size="xs" />
      <span className="truncate">{label ?? config.label}</span>
    </span>
  );
}
