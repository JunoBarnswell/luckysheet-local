import { AssetIcon, type AssetIconSize } from '@react-sheets/ui-system';

export const HOME_RIBBON_ICON_NAMES = [
  'align-center-vertical',
  'align-center',
  'align-left',
  'align-right',
  'align-vertical-justify-end',
  'align-vertical-space-around',
  'arrow-down-01',
  'arrow-down-square',
  'arrow-down',
  'arrow-up-left',
  'bold',
  'chevron-down',
  'circle-x',
  'clipboard-paste',
  'file-spreadsheet',
  'grid-2x2',
  'italic',
  'layers-2',
  'paint-bucket',
  'paintbrush-2',
  'scissors',
  'search',
  'sigma-square',
  'sort-asc',
  'strikethrough',
  'table-cells-merge',
  'table-properties',
  'table',
  'text-align-center',
  'text-wrap',
  'underline',
  'wand-sparkles',
  'wrap-text',
] as const;

export type HomeRibbonIconName = typeof HOME_RIBBON_ICON_NAMES[number];

export interface HomeRibbonIconProps {
  name: HomeRibbonIconName;
  size?: AssetIconSize;
  className?: string;
}

/** Exact SVG exported from Figma node 3:28. */
export function HomeRibbonIcon({ className, name, size = 'md' }: HomeRibbonIconProps) {
  return <AssetIcon className={className} size={size} src={`/figma/home-ribbon/${name}.svg`} />;
}
