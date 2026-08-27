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

const OFFICIAL_FLUENT_HOME_ASSETS: Partial<Record<HomeRibbonIconName, string>> = {
  'align-center-vertical': '/icons/fluent/ic_fluent_align_center_vertical_24_regular.svg', 'align-center': '/icons/fluent/ic_fluent_align_center_horizontal_24_regular.svg', 'align-left': '/icons/fluent/ic_fluent_align_left_24_regular.svg', 'align-right': '/icons/fluent/ic_fluent_align_right_24_regular.svg', 'align-vertical-justify-end': '/icons/fluent/ic_fluent_align_bottom_24_regular.svg',
  bold: '/icons/fluent/ic_fluent_text_bold_24_regular.svg', italic: '/icons/fluent/ic_fluent_text_italic_24_regular.svg', underline: '/icons/fluent/ic_fluent_text_underline_24_regular.svg', 'clipboard-paste': '/icons/fluent/ic_fluent_clipboard_paste_24_regular.svg', 'paint-bucket': '/icons/fluent/ic_fluent_paint_bucket_24_regular.svg', scissors: '/icons/fluent/ic_fluent_cut_24_regular.svg', search: '/icons/fluent/ic_fluent_search_24_regular.svg', 'table-cells-merge': '/icons/fluent/ic_fluent_table_cells_merge_24_regular.svg', table: '/icons/fluent/ic_fluent_table_24_regular.svg', 'text-wrap': '/icons/fluent/ic_fluent_text_wrap_24_regular.svg', 'wrap-text': '/icons/fluent/ic_fluent_text_wrap_24_regular.svg',
};

export function HomeRibbonIcon({ className, name, size = 'md' }: HomeRibbonIconProps) {
  const explicitSize = size === 'xs' ? '!h-3 !w-3' : size === 'sm' ? '!h-3.5 !w-3.5' : size === 'md' ? '!h-4 !w-4' : size === 'lg' ? '!h-5 !w-5' : '!h-8 !w-8';
  return <AssetIcon className={`${explicitSize} ${className ?? ''}`} size={size} src={OFFICIAL_FLUENT_HOME_ASSETS[name] ?? `/figma/home-ribbon/${name}.svg`} />;
}
