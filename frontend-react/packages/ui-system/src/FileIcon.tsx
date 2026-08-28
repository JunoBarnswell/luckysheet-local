import { cn } from './cn';
import { Icon, type IconSize } from './Icon';

export type FileIconKind = 'native-document' | 'workbook' | 'folder' | 'import';

export interface FileIconProps {
  kind?: FileIconKind;
  size?: 'sm' | 'md' | 'lg';
  className?: string;
  label?: string;
}

const iconSizes: Record<NonNullable<FileIconProps['size']>, string> = {
  sm: 'h-7 w-7 rounded-md',
  md: 'h-9 w-9 rounded-lg',
  lg: 'h-11 w-11 rounded-xl',
};

const glyphSizes: Record<NonNullable<FileIconProps['size']>, IconSize> = {
  sm: 'sm',
  md: 'md',
  lg: 'lg',
};

export function FileIcon({ kind = 'native-document', size = 'md', className, label }: FileIconProps) {
  const isFolder = kind === 'folder';
  return (
    <span
      aria-label={label}
      className={cn(
        'inline-flex shrink-0 items-center justify-center border',
        iconSizes[size],
        isFolder ? 'border-amber-200 bg-amber-50 text-amber-600' : 'border-brand-line bg-brand-soft text-brand',
        className,
      )}
      role={label ? 'img' : undefined}
    >
      <Icon name={isFolder ? 'folder' : kind === 'import' ? 'upload' : 'file-spreadsheet'} size={glyphSizes[size]} />
    </span>
  );
}
