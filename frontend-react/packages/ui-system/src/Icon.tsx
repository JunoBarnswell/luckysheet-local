import type { SVGAttributes } from 'react';
import { cn } from './cn';

export type IconName =
  | 'alert-circle'
  | 'align-center'
  | 'align-left'
  | 'align-right'
  | 'align-top'
  | 'align-middle'
  | 'align-bottom'
  | 'arrow-down'
  | 'arrow-left'
  | 'arrow-right'
  | 'arrow-up'
  | 'bold'
  | 'borders'
  | 'calculator'
  | 'chart'
  | 'chart-column'
  | 'chart-bar'
  | 'chart-line'
  | 'chart-area'
  | 'chart-pie'
  | 'chart-scatter'
  | 'data-chart'
  | 'barcode'
  | 'camera'
  | 'checkbox'
  | 'check'
  | 'check-circle'
  | 'clock'
  | 'chevron-down'
  | 'chevron-left'
  | 'chevron-right'
  | 'chevron-up'
  | 'clipboard'
  | 'comma'
  | 'cloud-check'
  | 'comment'
  | 'columns'
  | 'copy'
  | 'dollar-sign'
  | 'decimal-decrease'
  | 'decimal-increase'
  | 'download'
  | 'external-link'
  | 'eye'
  | 'fill-down'
  | 'fill-up'
  | 'fill-right'
  | 'fill-left'
  | 'file-plus'
  | 'file-spreadsheet'
  | 'file-text'
  | 'form-control'
  | 'filter'
  | 'freeze'
  | 'function'
  | 'folder'
  | 'folder-open'
  | 'grid'
  | 'help'
  | 'home'
  | 'history'
  | 'info'
  | 'italic'
  | 'indent-decrease'
  | 'indent-increase'
  | 'keyboard'
  | 'layout'
  | 'link'
  | 'loader'
  | 'lock'
  | 'maximize'
  | 'menu'
  | 'minimize'
  | 'merge-cells'
  | 'more-horizontal'
  | 'more-vertical'
  | 'paint-bucket'
  | 'picture'
  | 'palette'
  | 'pencil'
  | 'percent'
  | 'plus'
  | 'printer'
  | 'redo'
  | 'refresh'
  | 'rows'
  | 'save'
  | 'scissors'
  | 'search'
  | 'settings'
  | 'shape-circle'
  | 'shape-square'
  | 'share'
  | 'sliders'
  | 'sort'
  | 'star'
  | 'sparkles'
  | 'sparkline'
  | 'strikethrough'
  | 'table'
  | 'table-sheet'
  | 'gantt-sheet'
  | 'report-sheet'
  | 'table-pivot'
  | 'trash'
  | 'type'
  | 'textbox'
  | 'underline'
  | 'undo'
  | 'upload'
  | 'users'
  | 'wrap-text'
  | 'x'
  | 'zoom-in'
  | 'zoom-out';

export type IconSize = 'xs' | 'sm' | 'md' | 'lg' | 'xl';

const iconSizes: Record<IconSize, string> = {
  xs: 'h-3 w-3',
  sm: 'h-3.5 w-3.5',
  md: 'h-4 w-4',
  lg: 'h-5 w-5',
  xl: 'h-6 w-6',
};

interface IconProps extends SVGAttributes<SVGSVGElement> {
  name: IconName;
  size?: IconSize;
  title?: string;
}

function IconPath({ name }: Pick<IconProps, 'name'>) {
  switch (name) {
    case 'alert-circle':
      return <><circle cx="12" cy="12" r="9" /><path d="M12 7.5v5" /><path d="M12 16.5h.01" /></>;
    case 'align-center':
      return <><path d="M8 6h8" /><path d="M5 10h14" /><path d="M7 14h10" /><path d="M8 18h8" /></>;
    case 'align-left':
      return <><path d="M5 6h14" /><path d="M5 10h10" /><path d="M5 14h12" /><path d="M5 18h8" /></>;
    case 'align-right':
      return <><path d="M5 6h14" /><path d="M9 10h10" /><path d="M7 14h12" /><path d="M11 18h8" /></>;
    case 'align-top':
      return <><path d="M5 5h14" /><path d="M7 8h10M9 11h6M10 14h4" /></>;
    case 'align-middle':
      return <><path d="M4 12h16" /><path d="M7 7h10M9 10h6M9 14h6M7 17h10" /></>;
    case 'align-bottom':
      return <><path d="M5 19h14" /><path d="M10 10h4M9 13h6M7 16h10" /></>;
    case 'arrow-down':
      return <><path d="M12 5v14" /><path d="m6 13 6 6 6-6" /></>;
    case 'arrow-left':
      return <><path d="M19 12H5" /><path d="m11 18-6-6 6-6" /></>;
    case 'arrow-right':
      return <><path d="M5 12h14" /><path d="m13 6 6 6-6 6" /></>;
    case 'arrow-up':
      return <><path d="M12 19V5" /><path d="m6 11 6-6 6 6" /></>;
    case 'bold':
      return <path d="M7 5h5a4 4 0 0 1 2.4 7.2A4 4 0 0 1 13 19H7V5Zm3 3v3h2a1.5 1.5 0 0 0 0-3h-2Zm0 6v3h3a1.5 1.5 0 0 0 0-3h-3Z" fill="currentColor" stroke="none" />;
    case 'borders':
      return <><rect x="4" y="4" width="16" height="16" rx="1" /><path d="M4 12h16M12 4v16" strokeDasharray="2 2" /></>;
    case 'calculator':
      return <><rect x="5" y="3" width="14" height="18" rx="2" /><path d="M8 7h8M8 11h.01M12 11h.01M16 11h.01M8 15h.01M12 15h.01M16 15h.01" /></>;
    case 'chart':
      return <><path d="M5 19V5" /><path d="M5 19h14" /><path d="m8 15 3-4 3 2 4-6" /></>;
    case 'chart-column':
      return <><path d="M4 20h16" /><rect x="6" y="11" width="3" height="7" fill="#f59e0b" stroke="none" /><rect x="11" y="6" width="3" height="12" fill="#3b82f6" stroke="none" /><rect x="16" y="3" width="3" height="15" fill="#ef4444" stroke="none" /></>;
    case 'chart-bar':
      return <><path d="M4 4v16" /><rect x="6" y="6" width="12" height="3" fill="#3b82f6" stroke="none" /><rect x="6" y="11" width="8" height="3" fill="#f59e0b" stroke="none" /><rect x="6" y="16" width="14" height="3" fill="#10b981" stroke="none" /></>;
    case 'chart-line':
      return <><path d="M4 19V5M4 19h16" /><path d="m6 15 4-5 4 3 5-7" stroke="#3b82f6" strokeWidth="2" /><circle cx="10" cy="10" r="1" fill="#3b82f6" stroke="none" /><circle cx="19" cy="6" r="1" fill="#3b82f6" stroke="none" /></>;
    case 'chart-area':
      return <><path d="M4 19V5M4 19h16" /><path d="m5 17 4-6 4 3 6-8v11Z" fill="#93c5fd" stroke="#2563eb" /></>;
    case 'chart-pie':
      return <><path d="M12 3a9 9 0 1 0 9 9h-9Z" fill="#93c5fd" /><path d="M14 3.3V10h6.7A9 9 0 0 0 14 3.3Z" fill="#f59e0b" /></>;
    case 'chart-scatter':
      return <><path d="M4 5v15h16" /><circle cx="8" cy="15" r="1.5" fill="#3b82f6" stroke="none" /><circle cx="12" cy="11" r="1.5" fill="#10b981" stroke="none" /><circle cx="17" cy="7" r="1.5" fill="#ef4444" stroke="none" /></>;
    case 'data-chart':
      return <><rect x="3" y="4" width="7" height="16" rx="1" fill="#93c5fd" /><path d="M5 8h3M5 12h3M5 16h3" /><rect x="13" y="11" width="3" height="8" fill="#10b981" stroke="none" /><rect x="18" y="6" width="3" height="13" fill="#3b82f6" stroke="none" /></>;
    case 'barcode':
      return <><path d="M4 5v14M7 5v14M10 5v14M12 5v14M16 5v14M19 5v14" strokeWidth="2" /><path d="M5.5 5v14M14 5v14" /></>;
    case 'camera':
      return <><path d="M8 7 9.5 4h5L16 7h3a2 2 0 0 1 2 2v9H3V9a2 2 0 0 1 2-2Z" /><circle cx="12" cy="13" r="4" /><path d="M6 10h.01" /></>;
    case 'checkbox':
      return <><rect x="4" y="4" width="16" height="16" rx="1" /><path d="m7 12 3 3 7-7" strokeWidth="2" /></>;
    case 'check':
      return <path d="m5 12 4.5 4.5L19 7" />;
    case 'check-circle':
      return <><circle cx="12" cy="12" r="9" /><path d="m8 12 2.5 2.5L16 9" /></>;
    case 'clock':
      return <><circle cx="12" cy="12" r="8.5" /><path d="M12 7v5l3 2" /></>;
    case 'chevron-down':
      return <path d="m6 9 6 6 6-6" />;
    case 'chevron-left':
      return <path d="m15 18-6-6 6-6" />;
    case 'chevron-right':
      return <path d="m9 18 6-6-6-6" />;
    case 'chevron-up':
      return <path d="m18 15-6-6-6 6" />;
    case 'clipboard':
      return <><rect x="8" y="2" width="8" height="4" rx="1" /><rect x="4" y="6" width="16" height="15" rx="2" /></>;
    case 'comma':
      return <><path d="M5 7h3M11 7h3M17 7h2M5 13h3M11 13h3M17 13h2" /><path d="M17 18h2l-2 3" /></>;
    case 'cloud-check':
      return <><path d="M7.5 18.5h8.75a4.75 4.75 0 0 0 .84-9.43A6 6 0 0 0 5.38 10.8 4 4 0 0 0 7.5 18.5Z" /><path d="m9 14 2 2 4-4" /></>;
    case 'comment':
      return <><path d="M20 11.5a7.5 7.5 0 0 1-8 7.5 8.6 8.6 0 0 1-3.3-.66L4 20l1.66-3.86A7.2 7.2 0 0 1 4 11.5 7.5 7.5 0 0 1 12 4a7.5 7.5 0 0 1 8 7.5Z" /><path d="M8 11.5h.01M12 11.5h.01M16 11.5h.01" /></>;
    case 'columns':
      return <><rect x="5" y="5" width="5" height="14" rx="1" /><rect x="14" y="5" width="5" height="14" rx="1" /></>;
    case 'copy':
      return <><rect x="8" y="8" width="12" height="12" rx="2" /><path d="M16 8V6a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h2" /></>;
    case 'dollar-sign':
      return <><path d="M12 2v20" /><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6" /></>;
    case 'decimal-decrease':
      return <><path d="M4 7h2M10 7h2M4 12h2M10 12h2M16 8h4M18 6v4" /><path d="M16 16h4" /></>;
    case 'decimal-increase':
      return <><path d="M4 7h2M10 7h2M4 12h2M10 12h2M16 6v4M14 8h4" /><path d="M16 16h4" /></>;
    case 'download':
      return <><path d="M12 4v10" /><path d="m8 10 4 4 4-4" /><path d="M5 19h14" /></>;
    case 'external-link':
      return <><path d="M14 5h5v5" /><path d="m19 5-8 8" /><path d="M18 13v5a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h5" /></>;
    case 'eye':
      return <><path d="M3.5 12s3-5 8.5-5 8.5 5 8.5 5-3 5-8.5 5-8.5-5-8.5-5Z" /><circle cx="12" cy="12" r="2" /></>;
    case 'fill-down':
      return <><rect x="5" y="4" width="14" height="11" rx="1" /><path d="M12 8v10M8 14l4 4 4-4" /></>;
    case 'fill-up':
      return <><rect x="5" y="9" width="14" height="11" rx="1" /><path d="M12 16V6M8 10l4-4 4 4" /></>;
    case 'fill-right':
      return <><rect x="4" y="5" width="11" height="14" rx="1" /><path d="M8 12h10M14 8l4 4-4 4" /></>;
    case 'fill-left':
      return <><rect x="9" y="5" width="11" height="14" rx="1" /><path d="M16 12H6M10 8l-4 4 4 4" /></>;
    case 'file-plus':
      return <><path d="M7 3.5h6l4 4V20H7a2 2 0 0 1-2-2V5.5a2 2 0 0 1 2-2Z" /><path d="M13 3.5v4h4" /><path d="M12 11v5M9.5 13.5h5" /></>;
    case 'file-spreadsheet':
      return <><path d="M7 3.5h6l4 4V20H7a2 2 0 0 1-2-2V5.5a2 2 0 0 1 2-2Z" /><path d="M13 3.5v4h4" /><rect x="8" y="11" width="6" height="5" rx=".5" /><path d="M8 13.5h6M10 11v5M12 11v5" /></>;
    case 'file-text':
      return <><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><path d="M14 2v6h6M16 13H8M16 17H8M10 9H8" /></>;
    case 'filter':
      return <><path d="M4 6h16" /><path d="M7 12h10" /><path d="M10 18h4" /></>;
    case 'freeze':
      return <><rect x="5" y="5" width="14" height="14" rx="1" /><path d="M5 10h14M10 5v14" /><path d="m16 8 2-2" /></>;
    case 'form-control':
      return <><rect x="3" y="5" width="18" height="14" rx="1" /><rect x="6" y="8" width="5" height="5" rx=".5" /><path d="m7 10 1.3 1.3L10 9" /><path d="M13 9h5M13 12h5M6 16h12" /></>;
    case 'function':
      return <path d="M7.5 19c1.4-2.8 2.2-6.2 2.7-10.5h5.3M8 6h7M6 13h8" />;
    case 'folder':
      return <path d="M3.5 7.5a2 2 0 0 1 2-2h4l2 2h7a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2h-13a2 2 0 0 1-2-2v-9Z" />;
    case 'folder-open':
      return <><path d="M3.5 8.5a2 2 0 0 1 2-2h4l2 2h7a2 2 0 0 1 1.9 2.6l-1.6 5.4a2 2 0 0 1-1.9 1.4H5a2 2 0 0 1-1.9-1.4L1.9 11a2 2 0 0 1 1.6-2.5Z" /><path d="M3 11h16" /></>;
    case 'grid':
      return <><rect x="4.5" y="4.5" width="15" height="15" rx="2" /><path d="M4.5 10h15M4.5 15h15M10 4.5v15M15 4.5v15" /></>;
    case 'help':
      return <><circle cx="12" cy="12" r="9" /><path d="M9.6 9a2.5 2.5 0 1 1 4.1 1.94c-.9.7-1.7 1.14-1.7 2.56" /><path d="M12 16.5h.01" /></>;
    case 'home':
      return <><path d="m4 11 8-7 8 7" /><path d="M6 10v9h12v-9" /><path d="M10 19v-5h4v5" /></>;
    case 'history':
      return <><path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" /><path d="M3 3v5h5M12 7v5l4 2" /></>;
    case 'info':
      return <><circle cx="12" cy="12" r="9" /><path d="M12 11v5" /><path d="M12 8h.01" /></>;
    case 'italic':
      return <><path d="M10 5h7" /><path d="M7 19h7" /><path d="m14 5-4 14" /></>;
    case 'indent-decrease':
      return <><path d="M10 6h9M10 10h7M10 14h9M10 18h7" /><path d="m7 9-3 3 3 3" /></>;
    case 'indent-increase':
      return <><path d="M10 6h9M10 10h7M10 14h9M10 18h7" /><path d="m4 9 3 3-3 3" /></>;
    case 'keyboard':
      return <><rect x="3.5" y="6.5" width="17" height="11" rx="2" /><path d="M6.5 10h.01M9.5 10h.01M12.5 10h.01M15.5 10h.01M18.5 10h.01M6.5 14h9M17.5 14h1" /></>;
    case 'layout':
      return <><rect x="4" y="5" width="16" height="14" rx="2" /><path d="M4 10h16M10 10v9" /></>;
    case 'link':
      return <><path d="m9 15-2 2a3 3 0 0 1-4-4l4-4a3 3 0 0 1 4 0" /><path d="m15 9 2-2a3 3 0 0 1 4 4l-4 4a3 3 0 0 1-4 0" /><path d="m8 12 8 0" /></>;
    case 'loader':
      return <path d="M12 4a8 8 0 1 0 8 8" />;
    case 'lock':
      return <><rect x="5" y="10" width="14" height="10" rx="2" /><path d="M8 10V7a4 4 0 0 1 8 0v3" /><path d="M12 14v2" /></>;
    case 'maximize':
      return <><path d="M8 4H4v4M16 4h4v4M20 16v4h-4M4 16v4h4" /><path d="M4 4l5 5M20 4l-5 5M4 20l5-5M20 20l-5-5" /></>;
    case 'menu':
      return <><path d="M5 7h14M5 12h14M5 17h14" /></>;
    case 'minimize':
      return <path d="M5 17h14" />;
    case 'merge-cells':
      return <><rect x="4" y="6" width="16" height="12" rx="2" /><path d="M9 12h6M12 9v6" strokeDasharray="2 2" /><path d="m7 12 2-2M7 12l2 2M17 12l-2-2M17 12l2 2" /></>;
    case 'more-horizontal':
      return <><circle cx="5" cy="12" r="1" fill="currentColor" stroke="none" /><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none" /><circle cx="19" cy="12" r="1" fill="currentColor" stroke="none" /></>;
    case 'more-vertical':
      return <><circle cx="12" cy="5" r="1" fill="currentColor" stroke="none" /><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none" /><circle cx="12" cy="19" r="1" fill="currentColor" stroke="none" /></>;
    case 'paint-bucket':
      return <><path d="m19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2c.8.8 2 .8 2.8 0L19 11Z" /><path d="m5 2 5 5" /><path d="M2 13h15" /><path d="M22 20a2 2 0 1 1-4 0c0-1.6 2-4 2-4s2 2.4 2 4Z" /></>;
    case 'picture':
      return <><rect x="3" y="5" width="18" height="14" rx="1" /><circle cx="8" cy="10" r="2" fill="#f59e0b" stroke="none" /><path d="m5 17 5-5 3 3 2-2 4 4" fill="#93c5fd" /></>;
    case 'palette':
      return <path d="M12 4a8 8 0 0 0 0 16h1.1a1.9 1.9 0 0 0 .7-3.67A1.9 1.9 0 0 1 14.5 14H16a4 4 0 0 0 4-4 7 7 0 0 0-8-6Z" />;
    case 'pencil':
      return <><path d="m4 16-.8 4.8L8 20l10.8-10.8a2.5 2.5 0 0 0-3.5-3.5L4 16Z" /><path d="m13.5 7.5 3 3" /></>;
    case 'percent':
      return <><line x1="19" y1="5" x2="5" y2="19" /><circle cx="6.5" cy="6.5" r="2.5" /><circle cx="17.5" cy="17.5" r="2.5" /></>;
    case 'plus':
      return <><path d="M12 5v14M5 12h14" /></>;
    case 'printer':
      return <><path d="M6 9V2h12v7M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" /><rect x="6" y="14" width="12" height="8" rx="1" /></>;
    case 'redo':
      return <><path d="M17 7h4v4" /><path d="M21 7 14 14a5 5 0 0 1-7 0l-2-2" /></>;
    case 'refresh':
      return <><path d="M19 8a7 7 0 1 0 1 6" /><path d="M19 4v4h-4" /></>;
    case 'rows':
      return <><rect x="5" y="5" width="14" height="5" rx="1" /><rect x="5" y="14" width="14" height="5" rx="1" /></>;
    case 'save':
      return <><path d="M5 4h12l2 2v14H5V4Z" /><path d="M8 4v6h8V4M8 20v-6h8v6" /></>;
    case 'scissors':
      return <><circle cx="6" cy="6" r="3" /><circle cx="6" cy="18" r="3" /><path d="M20 4 8.12 15.88M14.47 14.48 20 20M8.12 8.12 12 12" /></>;
    case 'search':
      return <><circle cx="10.5" cy="10.5" r="5.5" /><path d="m15 15 4 4" /></>;
    case 'settings':
      return <><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.8 1.8 0 0 0 .36 1.98l.05.05-1.41 1.41-.05-.05a1.8 1.8 0 0 0-1.98-.36 1.8 1.8 0 0 0-1.1 1.65V20h-2v-.32a1.8 1.8 0 0 0-1.1-1.65 1.8 1.8 0 0 0-1.98.36l-.05.05-1.41-1.41.05-.05A1.8 1.8 0 0 0 9.2 15a1.8 1.8 0 0 0-1.65-1.1H7v-2h.55A1.8 1.8 0 0 0 9.2 10a1.8 1.8 0 0 0-.36-1.98l-.05-.05L10.2 6.6l.05.05a1.8 1.8 0 0 0 1.98.36 1.8 1.8 0 0 0 1.1-1.65V5h2v.32a1.8 1.8 0 0 0 1.1 1.65 1.8 1.8 0 0 0 1.98-.36l.05-.05 1.41 1.41-.05.05A1.8 1.8 0 0 0 19.4 10c.2.6.75 1 1.4 1H21v2h-.2a1.5 1.5 0 0 0-1.4 1Z" /></>;
    case 'shape-circle':
      return <circle cx="12" cy="12" r="8" />;
    case 'shape-square':
      return <rect x="5" y="5" width="14" height="14" rx="1" />;
    case 'share':
      return <><circle cx="18" cy="5" r="2.5" /><circle cx="6" cy="12" r="2.5" /><circle cx="18" cy="19" r="2.5" /><path d="m8.2 10.8 7.6-4.5M8.2 13.2l7.6 4.5" /></>;
    case 'sliders':
      return <><path d="M4 7h16M4 17h16" /><circle cx="9" cy="7" r="2" /><circle cx="15" cy="17" r="2" /></>;
    case 'sort':
      return <><path d="M8 5v14" /><path d="m5 8 3-3 3 3" /><path d="M16 19V5" /><path d="m13 16 3 3 3-3" /></>;
    case 'star':
      return <path d="m12 3 2.8 5.7 6.2.9-4.5 4.4 1.1 6.2-5.6-2.9-5.6 2.9 1.1-6.2L3 9.6l6.2-.9L12 3Z" />;
    case 'sparkles':
      return <><path d="m12 3 1.3 5.7L19 10l-5.7 1.3L12 17l-1.3-5.7L5 10l5.7-1.3L12 3Z" /><path d="m19 16 .5 2.2L22 19l-2.5.8L19 22l-.5-2.2L16 19l2.5-.8L19 16Z" /></>;
    case 'sparkline':
      return <><path d="m3 16 5-6 4 4 6-8 3 3" /><circle cx="18" cy="6" r="1.5" fill="currentColor" /></>;
    case 'strikethrough':
      return <><path d="M16 4H9a3 3 0 0 0 0 6h6a3 3 0 0 1 0 6H8M4 12h16" /></>;
    case 'table':
      return <><rect x="4" y="5" width="16" height="14" rx="1" /><path d="M4 10h16M4 14h16M10 5v14M15 5v14" /></>;
    case 'table-sheet':
      return <><rect x="3" y="4" width="18" height="16" rx="1" /><path d="M3 9h18M8 4v16M14 4v16" /><path d="M5 6h1M10 6h2M16 6h3" stroke="#3b82f6" /></>;
    case 'gantt-sheet':
      return <><rect x="3" y="4" width="18" height="16" rx="1" /><path d="M9 4v16M3 9h18M3 14h18" /><rect x="11" y="6" width="7" height="2" fill="#3b82f6" stroke="none" /><rect x="13" y="11" width="6" height="2" fill="#10b981" stroke="none" /><rect x="10" y="16" width="5" height="2" fill="#f59e0b" stroke="none" /></>;
    case 'report-sheet':
      return <><path d="M6 3h9l4 4v14H6Z" /><path d="M15 3v5h4M9 11h7M9 15h7M9 18h5" /><rect x="3" y="7" width="6" height="6" fill="#93c5fd" /></>;
    case 'table-pivot':
      return <><rect x="4" y="4" width="16" height="16" rx="2" /><path d="M4 9h16M9 4v16M4 14h5M9 14h11" /></>;
    case 'trash':
      return <><path d="M5 7h14M10 4h4l1 3H9l1-3Z" /><path d="m7 7 .8 13h8.4L17 7" /><path d="M10 11v5M14 11v5" /></>;
    case 'type':
      return <><path d="M4 7V4h16v3M9 20h6M12 4v16" /></>;
    case 'textbox':
      return <><rect x="3" y="5" width="18" height="14" rx="1" /><path d="M7 9V7h10v2M10 17h4M12 7v10" /></>;
    case 'underline':
      return <><path d="M7 5v5a5 5 0 0 0 10 0V5" /><path d="M5 19h14" /></>;
    case 'undo':
      return <><path d="M7 7H3v4" /><path d="M3 7 10 14a5 5 0 0 0 7 0l2-2" /></>;
    case 'upload':
      return <><path d="M12 20V10" /><path d="m8 14 4-4 4 4" /><path d="M5 5h14" /></>;
    case 'users':
      return <><circle cx="9" cy="8" r="3" /><path d="M3.5 19a5.5 5.5 0 0 1 11 0" /><path d="M16 5.5a3 3 0 0 1 0 5.8M17 13.5a5.5 5.5 0 0 1 3.5 5" /></>;
    case 'wrap-text':
      return <><path d="M4 6h16M4 18h10M4 12h13a3 3 0 0 1 0 6h-2" /><path d="m17 16-2 2 2 2" /></>;
    case 'x':
      return <><path d="m6 6 12 12M18 6 6 18" /></>;
    case 'zoom-in':
      return <><circle cx="10.5" cy="10.5" r="5.5" /><path d="m15 15 4 4M10.5 8v5M8 10.5h5" /></>;
    case 'zoom-out':
      return <><circle cx="10.5" cy="10.5" r="5.5" /><path d="m15 15 4 4M8 10.5h5" /></>;
  }
}

export function Icon({ name, size = 'md', className, title, ...props }: IconProps) {
  return (
    <svg
      aria-hidden={title ? undefined : true}
      aria-label={title}
      className={cn('shrink-0 fill-none stroke-current stroke-[1.7]', iconSizes[size], className)}
      fill="none"
      focusable="false"
      role={title ? 'img' : undefined}
      viewBox="0 0 24 24"
      {...props}
    >
      {title ? <title>{title}</title> : null}
      <IconPath name={name} />
    </svg>
  );
}
