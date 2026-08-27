import { forwardRef, type ImgHTMLAttributes } from 'react';
import { cn } from './cn';

export type AssetIconSize = 'xs' | 'sm' | 'md' | 'lg' | 'xl';

export interface AssetIconProps extends Omit<ImgHTMLAttributes<HTMLImageElement>, 'alt' | 'height' | 'width'> {
  alt?: string;
  size?: AssetIconSize;
}

const sizes: Record<AssetIconSize, string> = {
  xs: 'size-2',
  sm: 'size-3.5',
  md: 'size-4',
  lg: 'size-5',
  xl: 'size-8',
};

/** Renders committed design assets without recreating their vector geometry. */
export const AssetIcon = forwardRef<HTMLImageElement, AssetIconProps>(function AssetIcon(
  { alt = '', className, draggable = false, size = 'md', ...props },
  ref,
) {
  return (
    <img
      ref={ref}
      alt={alt}
      aria-hidden={alt ? undefined : true}
      className={cn('block shrink-0 object-contain', sizes[size], className)}
      draggable={draggable}
      {...props}
    />
  );
});
