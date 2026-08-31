import React from 'react';
import {
  Box,
  Button,
  ColorPicker,
  DropdownMenu,
  Inline,
  Select,
  Stack,
  Text,
  TextInput,
  type RibbonLayoutState,
} from '@react-sheets/ui-system';
import {
  type CompiledFeatureSurfaceSchema,
  type HomeRibbonState,
  type RibbonCommandContext,
  type RibbonCommandId,
  type RibbonControlId,
  type RibbonSurfaceDefinition,
  type RibbonMergeOperation,
  EXCEL_KEY_TIP_BINDINGS,
} from '@react-sheets/spreadsheet-app';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import { HOME_NUMBER_FORMAT_OPTIONS, homeText } from './home/home-localization';
import { HomeRibbonIcon, type HomeRibbonIconName } from './home/HomeRibbonIcon';
import { FontFamilyControl } from './FontFamilyControl';
import { RibbonLayoutRenderer } from './RibbonLayoutRenderer';

export interface HomeRibbonCommandOptions {
  className?: string;
  iconNode?: React.ReactNode;
  iconOnly?: boolean;
  iconOverride?: import('@react-sheets/ui-system').IconName;
  labelOverride?: string;
  ribbonLayoutNodeId?: string;
  ribbonSurfaceId?: string;
  testId?: string;
  tile?: boolean;
  trailingNode?: React.ReactNode;
}

export interface HomeRibbonProps {
  locale: Locale;
  layout: RibbonLayoutState;
  context: RibbonCommandContext;
  homeState: HomeRibbonState;
  disabled: boolean;
  featureSurfaceSchema: CompiledFeatureSurfaceSchema;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
  onEmitStyle: (style: Record<string, unknown>) => void;
  onBeginFormatPainter: (locked?: boolean) => void;
  formatPainterActive: boolean;
  onMergeCells: (operation: RibbonMergeOperation) => void;
  onOpenColumnWidth: () => void;
  onAutoFitColumns: () => void;
  onHideColumns: () => void;
  onUnhideColumns: () => void;
  onOpenDefaultColumnWidth: () => void;
  onOpenRowHeight: () => void;
  onAutoFitRows: () => void;
  onHideRows: () => void;
  onUnhideRows: () => void;
}

type HomeGroup = 'clipboard' | 'font' | 'alignment' | 'number' | 'styles' | 'cells' | 'editing';

const HOME_GROUPS: readonly HomeGroup[] = ['clipboard', 'font', 'alignment', 'number', 'styles', 'cells', 'editing'];
const EXCEL_FONT_SIZE_STEPS = [1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72, 409] as const;

function adjacentExcelFontSize(current: number, direction: 1 | -1): number {
  if (direction > 0) return EXCEL_FONT_SIZE_STEPS.find((value) => value > current) ?? 409;
  return [...EXCEL_FONT_SIZE_STEPS].reverse().find((value) => value < current) ?? 1;
}

const HomeRibbonLocaleContext = React.createContext<Locale>('en-US');

function surfaceLabel(locale: Locale, controlId: RibbonControlId): string {
  switch (controlId) {
    case 'format-painter': return homeText(locale, 'formatPainter');
    case 'font-family': return homeText(locale, 'fontFamily');
    case 'font-size': return homeText(locale, 'fontSize');
    case 'font-increase': return homeText(locale, 'increaseFontSize');
    case 'font-decrease': return homeText(locale, 'decreaseFontSize');
    case 'font-borders-menu': return homeText(locale, 'borders');
    case 'font-color': return homeText(locale, 'textColor');
    case 'fill-color': return homeText(locale, 'fillBackground');
    case 'number-format': return translateRibbonText(locale, 'groups.number');
    case 'alignment-menu': return homeText(locale, 'alignment');
    case 'orientation-menu': return homeText(locale, 'textOrientation');
    case 'cells-insert-menu': return homeText(locale, 'insert');
    case 'cells-delete-menu': return homeText(locale, 'delete');
    case 'cells-format-menu': return homeText(locale, 'format');
    case 'cell-styles-menu': return homeText(locale, 'cellStyles');
    case 'clear-menu': return homeText(locale, 'clear');
    case 'row-height': return homeText(locale, 'rowHeight');
    case 'auto-fit-row-height': return homeText(locale, 'autoFitRowHeight');
    case 'hide-rows': return homeText(locale, 'hideRows');
    case 'unhide-rows': return homeText(locale, 'unhideRows');
    case 'column-width': return homeText(locale, 'columnWidth');
    case 'auto-fit-column-width': return homeText(locale, 'autoFitColumnWidth');
    case 'hide-columns': return homeText(locale, 'hideColumns');
    case 'unhide-columns': return homeText(locale, 'unhideColumns');
    case 'default-column-width': return homeText(locale, 'defaultColumnWidth');
    case 'merge-menu': return homeText(locale, 'mergeCenter');
    case 'auto-sum-menu': return translateRibbonText(locale, 'commands.autoSum');
  }
}

const HOME_SMALL_ACTION_CLASS = '!h-6 !min-h-0 !w-6 !rounded-[var(--home-ribbon-radius)] !px-1';
const HOME_ALIGNMENT_ACTION_CLASS = '!h-[22px] !min-h-0 !w-[22px] !rounded-[var(--home-ribbon-radius)] !px-1';
const HOME_INLINE_ACTION_CLASS = '!h-6 !min-h-0 justify-start gap-1 rounded-[var(--home-ribbon-radius)] !px-1 text-[12px] font-normal leading-[14px] text-[var(--home-ribbon-color-text)]';
const HOME_LARGE_TILE_CLASS = '!h-[104px] !min-h-0 !max-w-none flex-col gap-1 rounded-[var(--home-ribbon-radius)] !px-2 !py-1 text-center text-[13px] font-normal leading-[16px] text-[var(--home-ribbon-color-text)] !whitespace-normal';
const HOME_EDITING_ACTION_CLASS = '!h-6 !min-h-0 justify-start gap-1 overflow-hidden rounded-[var(--home-ribbon-radius)] !px-1 text-[12px] font-normal leading-[14px] text-[var(--home-ribbon-color-text)] whitespace-nowrap';
const HOME_EDITING_MENU_CLASS = '!h-6 !min-h-0 !w-4 rounded-[var(--home-ribbon-radius)] !px-0';

const HOME_SURFACE_ICONS: Readonly<Partial<Record<string, { name: HomeRibbonIconName; size: React.ComponentProps<typeof HomeRibbonIcon>['size'] }>>> = {
  'clipboard.paste': { name: 'clipboard-paste', size: 'xl' },
  'clipboard.cut': { name: 'scissors', size: 'md' },
  'clipboard.copy': { name: 'layers-2', size: 'md' },
  'font.bold': { name: 'bold', size: 'md' },
  'font.italic': { name: 'italic', size: 'md' },
  'font.underline': { name: 'underline', size: 'md' },
  'alignment.top': { name: 'align-vertical-space-around', size: 'sm' },
  'alignment.middle': { name: 'align-center-vertical', size: 'sm' },
  'alignment.bottom': { name: 'align-vertical-justify-end', size: 'sm' },
  'alignment.left': { name: 'align-left', size: 'sm' },
  'alignment.center': { name: 'align-center', size: 'sm' },
  'alignment.right': { name: 'align-right', size: 'sm' },
  'alignment.wrap': { name: 'wrap-text', size: 'sm' },
  'number.decimal-increase': { name: 'arrow-down-01', size: 'sm' },
  'number.decimal-decrease': { name: 'arrow-down-01', size: 'sm' },
  'styles.conditional-format': { name: 'file-spreadsheet', size: 'xl' },
  'styles.table': { name: 'table', size: 'xl' },
  'editing.fill-down': { name: 'arrow-down-square', size: 'md' },
  'editing.sort': { name: 'sort-asc', size: 'md' },
  'editing.find': { name: 'search', size: 'md' },
};

function iconForSurface(surfaceId: string): React.ReactNode {
  const asset = HOME_SURFACE_ICONS[surfaceId];
  return asset ? <HomeRibbonIcon name={asset.name} size={asset.size} /> : undefined;
}

function HomeTile({ children, className, ...props }: React.ComponentProps<typeof Button>) {
  return <Button {...props} className={`${HOME_LARGE_TILE_CLASS} ${className ?? ''}`}>{children}</Button>;
}

/** HOME visual composition. Surface identity and responsive membership come only from the catalog. */
export function HomeRibbon({
  locale,
  layout,
  context,
  homeState,
  disabled,
  featureSurfaceSchema,
  renderCommand,
  onEmitStyle,
  onBeginFormatPainter,
  formatPainterActive,
  onMergeCells,
  onOpenColumnWidth,
  onAutoFitColumns,
  onHideColumns,
  onUnhideColumns,
  onOpenDefaultColumnWidth,
  onOpenRowHeight,
  onAutoFitRows,
  onHideRows,
  onUnhideRows,
}: HomeRibbonProps) {
  const cellStyle = homeState.style;
  const mixed = (key: keyof typeof cellStyle) => homeState.mixedStyleKeys.includes(key as never);
  const canFormat = !disabled && homeState.canFormat;
  const fontSizeMixed = mixed('fontSizePx');
  const fontSizeValue = fontSizeMixed ? '' : String(Math.round(pixelsToPoints(cellStyle.fontSizePx ?? pointsToPixels(11))));
  const [fontSizeDraft, setFontSizeDraft] = React.useState<string | undefined>(undefined);
  const fontSizeDraftRef = React.useRef<string | undefined>(undefined);
  const fontSizeCommitRef = React.useRef(false);
  const fontSizeCancelRef = React.useRef(false);
  React.useEffect(() => {
    fontSizeDraftRef.current = undefined;
    fontSizeCommitRef.current = false;
    fontSizeCancelRef.current = false;
    setFontSizeDraft(undefined);
  }, [fontSizeMixed, fontSizeValue]);
  const commitFontSizeDraft = (): void => {
    const raw = fontSizeDraftRef.current;
    if (raw === undefined) return;
    const value = Number(raw);
    fontSizeDraftRef.current = undefined;
    setFontSizeDraft(undefined);
    if (Number.isFinite(value) && value >= 1 && value <= 409) {
      fontSizeCommitRef.current = true;
      onEmitStyle({ fontSizePx: pointsToPixels(value) });
    }
  };
  const allSurfaces = (group: HomeGroup): readonly RibbonSurfaceDefinition[] => {
    const seen = new Set<string>();
    return (['wide', 'compact', 'narrow'] as const)
      .flatMap((candidate) => featureSurfaceSchema.ribbon
        .filter((surface) => surface.tab === 'home' && surface.group === group && surface.breakpoints.includes(candidate))
        .sort((left, right) => left.order - right.order))
      .filter((surface) => {
        if (seen.has(surface.id)) return false;
        seen.add(surface.id);
        return true;
      });
  };
  const menuMembers = (menuId: string): readonly RibbonSurfaceDefinition[] => HOME_GROUPS.flatMap((group) => allSurfaces(group)).filter((surface) => surface.menuId === menuId);

  const renderSurface = (surface: RibbonSurfaceDefinition, mode: RibbonLayoutState['mode'] | 'menu'): React.ReactNode => {
    if (surface.controlId) return renderControl(surface.controlId, mode, surface.id);
    if (!surface.commandId) return null;
    if (mode === 'menu') return renderCommand(surface.commandId, { className: 'w-full justify-start', ribbonSurfaceId: surface.id });
    const iconNode = iconForSurface(surface.id);
    if (surface.id.endsWith('.dialog-launcher')) return renderCommand(surface.commandId, { iconOverride: 'arrow-down', iconOnly: true, className: '!h-4 !min-h-0 !w-4 rotate-[-45deg] rounded-none p-0', ribbonSurfaceId: surface.id });
    if (surface.id === 'clipboard.paste') {
      return <Inline gap="none" className="h-[104px] items-stretch">
        {renderCommand(surface.commandId, { iconNode, tile: true, className: `${HOME_LARGE_TILE_CLASS} !w-[42px] !min-w-[42px] rounded-r-none`, ribbonSurfaceId: surface.id })}
        <DropdownMenu align="left" trigger={<Button aria-label={translateRibbonText(locale, 'commands.pasteSpecial')} disabled={disabled} iconNode={<HomeRibbonIcon name="chevron-down" size="xs" />} iconOnly size="sm" variant="ghost" className="!h-[104px] !min-h-0 !w-4 rounded-l-none px-0" />}>
          <Stack gap="none" className="min-w-[15rem] p-1">{menuMembers('clipboard.paste').map((member) => renderSurface(member, 'menu'))}</Stack>
        </DropdownMenu>
      </Inline>;
    }
    if (surface.id === 'clipboard.cut' || surface.id === 'clipboard.copy') {
      return renderCommand(surface.commandId, { iconNode, className: `${HOME_INLINE_ACTION_CLASS} ${surface.id === 'clipboard.cut' ? '!w-[52px]' : '!w-[52px]'}`, ribbonSurfaceId: surface.id });
    }
    if (surface.group === 'font') {
      return renderCommand(surface.commandId, { iconNode, iconOnly: true, className: HOME_SMALL_ACTION_CLASS, ribbonSurfaceId: surface.id });
    }
    if (surface.group === 'alignment') {
      if (surface.id === 'alignment.wrap') return renderCommand(surface.commandId, { iconNode, className: `${HOME_INLINE_ACTION_CLASS} !w-[74px]`, ribbonSurfaceId: surface.id });
      return renderCommand(surface.commandId, { iconNode, iconOnly: true, className: HOME_ALIGNMENT_ACTION_CLASS, ribbonSurfaceId: surface.id });
    }
    if (surface.id === 'number.percent' || surface.id === 'number.comma') {
      const symbol = surface.id === 'number.percent' ? '%' : ',';
      return renderCommand(surface.commandId, {
        iconNode: <Text className="text-[11px] font-bold leading-[13px] text-black">{symbol}</Text>,
        iconOnly: true,
        className: `!h-[19px] !min-h-0 rounded-[var(--home-ribbon-radius)] border border-[var(--home-ribbon-color-border)] px-[3px] ${surface.id === 'number.percent' ? '!w-[17px]' : '!w-[10px]'}`,
        ribbonSurfaceId: surface.id,
      });
    }
    if (surface.id === 'number.decimal-increase' || surface.id === 'number.decimal-decrease') {
      return renderCommand(surface.commandId, { iconNode, iconOnly: true, className: '!h-[18px] !min-h-0 !w-[18px] rounded-[var(--home-ribbon-radius)] border border-[var(--home-ribbon-color-border)] px-[3px]', ribbonSurfaceId: surface.id });
    }
    if (surface.id === 'styles.conditional-format' || surface.id === 'styles.table') {
      const widthClass = surface.id === 'styles.conditional-format'
        ? (mode === 'wide' ? '!w-[74px] !min-w-[74px]' : '!w-[58px] !min-w-[58px]')
        : (mode === 'wide' ? '!w-[98px] !min-w-[98px]' : '!w-[80px] !min-w-[80px]');
      return renderCommand(surface.commandId, { iconNode, tile: true, trailingNode: <HomeRibbonIcon name="chevron-down" size="xs" />, className: `${HOME_LARGE_TILE_CLASS} ${widthClass}`, ribbonSurfaceId: surface.id });
    }
    if (surface.id === 'editing.fill-down') {
      return <Inline gap="none" className={`${mode === 'wide' ? 'w-16' : 'w-14'} h-6 items-stretch`}>
        {renderCommand(surface.commandId, { iconNode, className: `${HOME_EDITING_ACTION_CLASS} !w-12`, labelOverride: homeText(locale, 'fill'), ribbonSurfaceId: surface.id })}
        <DropdownMenu align="left" trigger={<Button aria-label="Fill options" title="Fill options" disabled={disabled} iconNode={<HomeRibbonIcon name="chevron-down" size="xs" />} iconOnly size="sm" variant="ghost" className={HOME_EDITING_MENU_CLASS} />}>
          <Stack gap="none" className="min-w-[10rem] p-1">{menuMembers('editing.fill-down').map((member) => renderSurface(member, 'menu'))}</Stack>
        </DropdownMenu>
      </Inline>;
    }
    if (surface.id === 'editing.sort' || surface.id === 'editing.find') {
      const isSort = surface.id === 'editing.sort';
      const triggerLabel = homeText(locale, isSort ? 'sortAndFilter' : 'findAndSelect');
      return <DropdownMenu align="left" disabled={disabled} trigger={<Button aria-label={triggerLabel} data-ribbon-surface={surface.id} title={triggerLabel} disabled={disabled} iconNode={iconNode} size="sm" variant="ghost" className={`${HOME_EDITING_ACTION_CLASS} ${mode === 'wide' ? '!w-[100px]' : '!w-[84px]'}`}>{triggerLabel}<HomeRibbonIcon name="chevron-down" size="xs" /></Button>}>
        <Stack gap="none" className="min-w-[11rem] p-1">
          {renderCommand(surface.commandId, { className: 'w-full justify-start', ribbonSurfaceId: surface.id })}
          {menuMembers(surface.id).map((member) => renderSurface(member, 'menu'))}
        </Stack>
      </DropdownMenu>;
    }
    const tile = surface.appearance === 'large' || surface.appearance === 'tile';
    return renderCommand(surface.commandId, { iconNode, tile, iconOnly: !tile && surface.appearance === 'small', className: tile ? HOME_LARGE_TILE_CLASS : HOME_SMALL_ACTION_CLASS, ribbonSurfaceId: surface.id });
  };

  const renderControl = (controlId: RibbonControlId, mode: RibbonLayoutState['mode'] | 'menu', surfaceId: string): React.ReactNode => {
    const label = surfaceLabel(locale, controlId);
    const keyTip = EXCEL_KEY_TIP_BINDINGS.find((binding) => binding.target.kind === 'command' && binding.target.id === controlId)?.sequence;
    const compactMenu = controlId === 'font-borders-menu' || controlId === 'orientation-menu';
    const menuTrigger = (iconName: HomeRibbonIconName, presentation: 'tile' | 'inline' | 'editing-inline' | 'style' | 'cell' = 'tile') => mode !== 'menu'
      ? compactMenu
        ? <Button aria-label={label} data-ribbon-keytip={keyTip} data-ribbon-surface={surfaceId} title={label} disabled={disabled} iconNode={<HomeRibbonIcon name={iconName} size="md" />} iconOnly size="sm" variant="ghost" className={controlId === 'orientation-menu' ? HOME_SMALL_ACTION_CLASS : HOME_SMALL_ACTION_CLASS} />
        : presentation === 'inline' || presentation === 'editing-inline'
          ? <Button aria-label={label} data-ribbon-keytip={keyTip} data-ribbon-surface={surfaceId} title={label} disabled={disabled} iconNode={<HomeRibbonIcon name={iconName} size={presentation === 'inline' ? 'sm' : 'md'} />} size="sm" variant="ghost" className={presentation === 'editing-inline' ? `${HOME_EDITING_ACTION_CLASS} ${mode === 'wide' ? '!w-16' : '!w-14'}` : `${HOME_INLINE_ACTION_CLASS} ${mode === 'wide' ? '!w-[97px]' : '!w-[84px]'}`}>{label}<HomeRibbonIcon name="chevron-down" size="xs" /></Button>
        : <HomeTile aria-label={label} data-ribbon-surface={surfaceId} title={label} disabled={disabled} iconNode={<HomeRibbonIcon name={iconName} size="xl" />} type="button" className={presentation === 'style' ? `${mode === 'wide' ? '!w-[86px] !min-w-[86px]' : '!w-[70px] !min-w-[70px]'}` : presentation === 'cell' ? `${mode === 'wide' ? '!w-[50px] !min-w-[50px]' : '!w-[44px] !min-w-[44px]'}` : undefined}><Inline gap="none" className="gap-0.5">{label}<HomeRibbonIcon name="chevron-down" size="xs" /></Inline></HomeTile>
      : <Button aria-label={label} data-ribbon-surface={surfaceId} title={label} disabled={disabled} iconNode={<HomeRibbonIcon name={iconName} size="md" />} size="sm" variant="ghost" className="w-full justify-start">{label}</Button>;
    switch (controlId) {
      case 'format-painter':
        return <Button aria-label={label} aria-pressed={formatPainterActive} data-ribbon-keytip="HFP" data-ribbon-surface={surfaceId} data-testid="home-format-painter" disabled={!canFormat} iconNode={<HomeRibbonIcon name="paintbrush-2" size="md" />} iconOnly={false} size="sm" title={homeText(locale, 'formatPainterHint')} variant="ghost" className={mode === 'wide' ? `${HOME_INLINE_ACTION_CLASS} !w-16` : 'w-full justify-start'} onClick={() => onBeginFormatPainter(false)} onDoubleClick={() => onBeginFormatPainter(true)}>{label}</Button>;
      case 'font-family':
        return <Box data-ribbon-surface={surfaceId} className={mode === 'menu' ? 'w-full' : mode === 'wide' ? 'w-[214px] shrink-0' : 'w-[100px] shrink-0'}><FontFamilyControl
          value={cellStyle.fontFamily}
          fallbackValue="Microsoft YaHei"
          mixed={mixed('fontFamily')}
          mixedPlaceholder={homeText(locale, 'mixed')}
          className="!h-6 !w-full !rounded-[var(--home-ribbon-radius)] !border-[var(--home-ribbon-color-border)] !px-1.5 text-[11px] leading-[13px]"
          disabled={!canFormat}
          label={label}
          testId="home-font-family"
          onCommit={(fontFamily) => onEmitStyle({ fontFamily })}
        /></Box>;
      case 'font-size':
        return <TextInput
          aria-label={label}
          data-ribbon-surface={surfaceId}
          className={mode === 'menu' ? 'w-full' : '!h-6 !w-[45px] !rounded-[var(--home-ribbon-radius)] !border-[var(--home-ribbon-color-border)] !px-1.5 text-[11px] leading-[13px]'}
          disabled={!canFormat}
          inputMode="decimal"
          value={fontSizeDraft ?? fontSizeValue}
          onChange={(event) => {
            fontSizeCommitRef.current = false;
            fontSizeCancelRef.current = false;
            fontSizeDraftRef.current = event.target.value;
            setFontSizeDraft(event.target.value);
          }}
          onBlur={() => {
            if (fontSizeCommitRef.current || fontSizeCancelRef.current) {
              fontSizeCommitRef.current = false;
              fontSizeCancelRef.current = false;
              return;
            }
            commitFontSizeDraft();
          }}
          onKeyDown={(event) => {
            if (event.key === 'Enter') {
              event.preventDefault();
              commitFontSizeDraft();
              event.currentTarget.blur();
            } else if (event.key === 'Escape') {
              event.preventDefault();
              fontSizeDraftRef.current = undefined;
              fontSizeCancelRef.current = true;
              setFontSizeDraft(undefined);
              event.currentTarget.blur();
            }
          }}
        />;
      case 'font-increase':
      case 'font-decrease':
        return <Button aria-label={label} disabled={!canFormat} size="sm" variant="ghost" className={mode !== 'menu' ? '!h-[19px] !min-h-0 !w-[25px] rounded-[var(--home-ribbon-radius)] border border-[var(--home-ribbon-color-border)] px-[3px] text-[11px] font-bold leading-[13px] text-black' : 'w-full justify-start'} onClick={() => onEmitStyle({ fontSizePx: pointsToPixels(adjacentExcelFontSize(pixelsToPoints(cellStyle.fontSizePx ?? pointsToPixels(11)), controlId === 'font-increase' ? 1 : -1)) })}>{mode !== 'menu' ? (controlId === 'font-increase' ? 'A↑' : 'A↓') : label}</Button>;
      case 'font-color':
      case 'fill-color':
        return <DropdownMenu disabled={!canFormat} trigger={mode !== 'menu' ? <Button aria-label={label} data-ribbon-keytip={keyTip} disabled={!canFormat} iconNode={<HomeRibbonIcon name={controlId === 'font-color' ? 'text-align-center' : 'paint-bucket'} size="md" />} iconOnly size="sm" variant="ghost" className={HOME_SMALL_ACTION_CLASS} /> : menuTrigger(controlId === 'font-color' ? 'text-align-center' : 'paint-bucket')}>
          {({ close }) => <ColorPicker color={controlId === 'font-color' ? cellStyle.textColor ?? '#1e293b' : cellStyle.background ?? '#ffffff'} onChange={(color) => { onEmitStyle({ [controlId === 'font-color' ? 'textColor' : 'background']: color }); close(); }} />}
        </DropdownMenu>;
      case 'font-borders-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('grid-2x2')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.font-borders-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
      case 'number-format':
        return <Select aria-label={label} data-ribbon-surface={surfaceId} className="!h-6 !w-[120px] !rounded-[var(--home-ribbon-radius)] !border-[var(--home-ribbon-color-border)] !px-1.5 text-[11px] leading-[13px]" disabled={!canFormat} sizeVariant="sm" value={mixed('numberFormat') ? '__mixed__' : cellStyle.numberFormat || 'general'} onChange={(event) => { if (event.target.value === '__more__') { context.dispatchSessionIntent({ type: 'dialog.open', dialog: 'format-cells', formatCellsTab: 'number' }); return; } if (event.target.value !== '__mixed__') onEmitStyle({ numberFormat: event.target.value }); }}>
          {mixed('numberFormat') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
          {HOME_NUMBER_FORMAT_OPTIONS.map(({ value, labelKey }) => <option key={value} value={value}>{homeText(locale, labelKey)}</option>)}
          <option value="__more__">{homeText(locale, 'numberPresetMore')}</option>
        </Select>;
      case 'alignment-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('align-left')}>
          <Stack gap="none" className="min-w-[15rem] p-1">{menuMembers('control.alignment-menu').map((surface) => renderSurface(surface, 'menu'))}</Stack>
        </DropdownMenu>;
      case 'orientation-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('text-wrap')}>
          <Stack gap="none" className="min-w-[15rem] p-1">{menuMembers('control.orientation-menu').map((surface) => renderSurface(surface, 'menu'))}</Stack>
        </DropdownMenu>;
      case 'cells-insert-menu':
      case 'cells-delete-menu':
      case 'cells-format-menu': {
        const icon: HomeRibbonIconName = controlId === 'cells-format-menu' ? 'table-properties' : 'file-spreadsheet';
        return <DropdownMenu align="left" trigger={menuTrigger(icon, 'cell')}><Stack gap="none" className="min-w-[13rem] p-1">{menuMembers(`control.${controlId}`).map((surface) => renderSurface(surface, 'menu'))}</Stack></DropdownMenu>;
      }
      case 'cell-styles-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('file-spreadsheet', 'style')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.cell-styles-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
      case 'clear-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('wand-sparkles', 'editing-inline')}><Stack gap="none" className="min-w-[14rem] p-1">{menuMembers('control.clear-menu').map((surface) => renderSurface(surface, 'menu'))}</Stack></DropdownMenu>;
      case 'column-width':
      case 'auto-fit-column-width':
      case 'hide-columns':
      case 'unhide-columns':
      case 'default-column-width': {
        const actions: Record<typeof controlId, () => void> = {
          'column-width': onOpenColumnWidth,
          'auto-fit-column-width': onAutoFitColumns,
          'hide-columns': onHideColumns,
          'unhide-columns': onUnhideColumns,
          'default-column-width': onOpenDefaultColumnWidth,
        };
        const commandId = controlId === 'hide-columns' || controlId === 'unhide-columns' ? 'sheet.columns.visibility.set'
          : controlId === 'default-column-width' ? 'sheet.column.defaultWidth.set' : 'sheet.dimensions.apply';
        const permitted = !disabled && (!context.canExecute || context.canExecute(commandId));
        return <Button aria-label={label} data-ribbon-surface={surfaceId} title={permitted ? label : `${label} — permission denied`} disabled={!permitted} size="sm" variant="ghost" className="w-full justify-start" onClick={actions[controlId]}>{label}</Button>;
      }
      case 'row-height':
      case 'auto-fit-row-height':
      case 'hide-rows':
      case 'unhide-rows': {
        const actions: Record<typeof controlId, () => void> = {
          'row-height': onOpenRowHeight,
          'auto-fit-row-height': onAutoFitRows,
          'hide-rows': onHideRows,
          'unhide-rows': onUnhideRows,
        };
        const commandId = controlId === 'hide-rows' || controlId === 'unhide-rows' ? 'sheet.rows.visibility.set' : 'sheet.dimensions.apply';
        const permitted = !disabled && (!context.canExecute || context.canExecute(commandId));
        return <Button aria-label={label} data-ribbon-surface={surfaceId} title={permitted ? label : `${label} — permission denied`} disabled={!permitted} size="sm" variant="ghost" className="w-full justify-start" onClick={actions[controlId]}>{label}</Button>;
      }
      case 'merge-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('table-cells-merge', 'inline')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.merge-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
      case 'auto-sum-menu':
        return <Inline gap="none" className={mode !== 'menu' ? 'h-6 w-[88px] items-stretch' : 'w-full'}>
          {renderCommand('autoSum', { iconNode: <HomeRibbonIcon name="sigma-square" size="md" />, className: mode !== 'menu' ? `${HOME_EDITING_ACTION_CLASS} !w-[72px]` : 'w-full justify-start', labelOverride: homeText(locale, 'autoSumCompact'), ribbonSurfaceId: surfaceId })}
          <DropdownMenu align="left" trigger={<Button aria-label={`${label} options`} title={`${label} options`} disabled={disabled} iconNode={<HomeRibbonIcon name="chevron-down" size="xs" />} iconOnly size="sm" variant="ghost" className={mode !== 'menu' ? HOME_EDITING_MENU_CLASS : 'w-5 shrink-0 justify-center px-0'} />}>
            <Stack gap="none" className="min-w-[10rem] p-1">
              {menuMembers('control.auto-sum-menu').map((surface) => renderSurface(surface, 'menu'))}
            </Stack>
          </DropdownMenu>
        </Inline>;
    }
  };

  return (
    <HomeRibbonLocaleContext.Provider value={locale}>
      <RibbonLayoutRenderer
        tab="home"
        locale={locale}
        layout={layout}
        featureSurfaceSchema={featureSurfaceSchema}
        renderCommand={renderCommand}
        renderSurface={(surface, context) => renderSurface(surface, context.mode)}
      />
    </HomeRibbonLocaleContext.Provider>
  );
}
