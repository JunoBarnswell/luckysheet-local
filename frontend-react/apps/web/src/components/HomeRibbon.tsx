import React, { useMemo } from 'react';
import {
  Box,
  Button,
  ColorPicker,
  Divider,
  DropdownMenu,
  Inline,
  Select,
  Stack,
  Text,
  TextInput,
  type RibbonLayoutState,
} from '@react-sheets/ui-system';
import {
  getRibbonGroupDefinition,
  getRibbonSurfaces,
  type HomeRibbonState,
  type RibbonCommandContext,
  type RibbonCommandId,
  type RibbonControlId,
  type RibbonSurfaceBreakpoint,
  type RibbonSurfaceDefinition,
  type RibbonMergeOperation,
} from '@react-sheets/spreadsheet-app';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import { HOME_NUMBER_FORMAT_OPTIONS, homeText } from './home/home-localization';

export interface HomeRibbonCommandOptions {
  className?: string;
  iconOnly?: boolean;
  testId?: string;
  tile?: boolean;
}

export interface HomeRibbonProps {
  locale: Locale;
  layout: RibbonLayoutState;
  context: RibbonCommandContext;
  homeState: HomeRibbonState;
  disabled: boolean;
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
}

type HomeBreakpoint = RibbonSurfaceBreakpoint;
type HomeGroup = 'history' | 'clipboard' | 'font' | 'alignment' | 'number' | 'styles' | 'cells' | 'editing';

const HOME_GROUPS: readonly HomeGroup[] = ['history', 'clipboard', 'font', 'alignment', 'number', 'styles', 'cells', 'editing'];

function breakpointFor(layout: RibbonLayoutState): HomeBreakpoint {
  if (layout.width >= 1280) return 'wide';
  if (layout.width >= 1024) return 'compact';
  return 'narrow';
}

const HomeRibbonLocaleContext = React.createContext<Locale>('en-US');

function surfaceLabel(locale: Locale, controlId: RibbonControlId): string {
  switch (controlId) {
    case 'format-painter': return homeText(locale, 'formatPainter');
    case 'font-family': return homeText(locale, 'fontFamily');
    case 'font-size': return homeText(locale, 'fontSize');
    case 'font-increase': return homeText(locale, 'increaseFontSize');
    case 'font-decrease': return homeText(locale, 'decreaseFontSize');
    case 'font-color': return homeText(locale, 'textColor');
    case 'fill-color': return homeText(locale, 'fillBackground');
    case 'number-format': return translateRibbonText(locale, 'groups.number');
    case 'cells-insert-menu': return homeText(locale, 'insert');
    case 'cells-delete-menu': return homeText(locale, 'delete');
    case 'cells-format-menu': return homeText(locale, 'format');
    case 'clear-menu': return homeText(locale, 'clear');
    case 'column-width': return homeText(locale, 'columnWidth');
    case 'auto-fit-column-width': return homeText(locale, 'autoFitColumnWidth');
    case 'hide-columns': return homeText(locale, 'hideColumns');
    case 'unhide-columns': return homeText(locale, 'unhideColumns');
    case 'default-column-width': return homeText(locale, 'defaultColumnWidth');
    case 'merge-menu': return homeText(locale, 'mergeCenter');
  }
}

function HomeTile({ children, className, ...props }: React.ComponentProps<typeof Button>) {
  return <Button {...props} className={`!h-[68px] !min-h-0 !w-[72px] flex-col gap-1 rounded-none px-1 text-[10px] leading-3 [&>svg]:!h-6 [&>svg]:!w-6 ${className ?? ''}`}>{children}</Button>;
}

/** HOME visual composition. Surface identity and responsive membership come only from the catalog. */
export function HomeRibbon({
  locale,
  layout,
  context,
  homeState,
  disabled,
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
}: HomeRibbonProps) {
  const breakpoint = breakpointFor(layout);
  const cellStyle = homeState.style;
  const mixed = (key: keyof typeof cellStyle) => homeState.mixedStyleKeys.includes(key as never);
  const canFormat = !disabled && homeState.canFormat;
  const surfacesByGroup = useMemo(
    () => new Map(HOME_GROUPS.map((group) => [group, getRibbonSurfaces('home', group, breakpoint)] as const)),
    [breakpoint],
  );
  const allSurfaces = (group: HomeGroup): readonly RibbonSurfaceDefinition[] => {
    const seen = new Set<string>();
    return (['wide', 'compact', 'narrow'] as const)
      .flatMap((candidate) => getRibbonSurfaces('home', group, candidate))
      .filter((surface) => {
        if (seen.has(surface.id)) return false;
        seen.add(surface.id);
        return true;
      });
  };
  const topLevelSurfaces = (group: HomeGroup): readonly RibbonSurfaceDefinition[] => allSurfaces(group).filter((surface) => !surface.menuId);
  const menuMembers = (menuId: string): readonly RibbonSurfaceDefinition[] => HOME_GROUPS.flatMap((group) => allSurfaces(group)).filter((surface) => surface.menuId === menuId);

  const renderSurface = (surface: RibbonSurfaceDefinition, mode: 'wide' | 'menu'): React.ReactNode => {
    if (surface.controlId) return renderControl(surface.controlId, mode);
    if (!surface.commandId) return null;
    if (mode === 'menu') return renderCommand(surface.commandId, { className: 'w-full justify-start' });
    const tile = surface.appearance === 'large' || surface.appearance === 'tile';
    return renderCommand(surface.commandId, { tile, iconOnly: !tile && surface.appearance === 'small', className: tile ? undefined : '!h-8 !min-h-0 !rounded-none' });
  };

  const renderControl = (controlId: RibbonControlId, mode: 'wide' | 'menu'): React.ReactNode => {
    const label = surfaceLabel(locale, controlId);
    const menuTrigger = (icon: React.ComponentProps<typeof Button>['icon']) => mode === 'wide'
      ? <HomeTile aria-label={label} title={label} disabled={disabled} icon={icon} type="button">{label}</HomeTile>
      : <Button aria-label={label} title={label} disabled={disabled} icon={icon} size="sm" variant="ghost" className="w-full justify-start">{label}</Button>;
    switch (controlId) {
      case 'format-painter':
        return <Button aria-label={label} aria-pressed={formatPainterActive} data-testid="home-format-painter" disabled={!canFormat} icon="palette" iconOnly={mode === 'wide'} size="sm" title={homeText(locale, 'formatPainterHint')} variant="ghost" className={mode === 'wide' ? '!h-8 !min-h-0 !w-8 !rounded-none' : 'w-full justify-start'} onClick={() => onBeginFormatPainter(false)} onDoubleClick={() => onBeginFormatPainter(true)}>{mode === 'menu' ? label : null}</Button>;
      case 'font-family':
        return <Box className={mode === 'wide' ? 'w-[124px] shrink-0' : 'w-full'}><Select aria-label={label} className="w-full" disabled={!canFormat} sizeVariant="sm" value={mixed('fontFamily') ? '__mixed__' : cellStyle.fontFamily ?? 'Microsoft YaHei'} onChange={(event) => { if (event.target.value !== '__mixed__') onEmitStyle({ fontFamily: event.target.value }); }}>
          {mixed('fontFamily') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
          <option value="Microsoft YaHei">微软雅黑</option><option value="Arial">Arial</option><option value="Calibri">Calibri</option><option value="Segoe UI">Segoe UI</option><option value="Times New Roman">Times New Roman</option>
        </Select></Box>;
      case 'font-size':
        return <TextInput aria-label={label} className={mode === 'wide' ? '!w-[48px]' : 'w-full'} disabled={!canFormat} inputMode="decimal" value={mixed('fontSizePx') ? '' : String(Math.round(pixelsToPoints(cellStyle.fontSizePx ?? pointsToPixels(11))))} onChange={(event) => { const value = Number(event.target.value); if (Number.isFinite(value) && value >= 1 && value <= 409) onEmitStyle({ fontSizePx: pointsToPixels(value) }); }} />;
      case 'font-increase':
      case 'font-decrease':
        return <Button aria-label={label} disabled={!canFormat} size="sm" variant="ghost" className={mode === 'wide' ? '!h-8 !min-h-0 !w-7 rounded-none px-0 font-semibold text-[#2572bc]' : 'w-full justify-start'} onClick={() => onEmitStyle({ fontSizePx: controlId === 'font-increase' ? Math.min(pointsToPixels(409), (cellStyle.fontSizePx ?? pointsToPixels(11)) + pointsToPixels(1)) : Math.max(pointsToPixels(1), (cellStyle.fontSizePx ?? pointsToPixels(11)) - pointsToPixels(1)) })}>{mode === 'wide' ? 'A' : label}</Button>;
      case 'font-color':
      case 'fill-color':
        return <DropdownMenu disabled={!canFormat} trigger={mode === 'wide' ? <Button aria-label={label} disabled={!canFormat} icon={controlId === 'font-color' ? 'type' : 'paint-bucket'} iconOnly size="sm" variant="ghost" className="!h-8 !min-h-0 !w-8 !rounded-none" /> : menuTrigger(controlId === 'font-color' ? 'type' : 'paint-bucket')}>
          {({ close }) => <ColorPicker color={controlId === 'font-color' ? cellStyle.textColor ?? '#1e293b' : cellStyle.background ?? '#ffffff'} onChange={(color) => { onEmitStyle({ [controlId === 'font-color' ? 'textColor' : 'background']: color }); close(); }} />}
        </DropdownMenu>;
      case 'number-format':
        return <Select aria-label={label} className="w-full" disabled={!canFormat} sizeVariant="sm" value={mixed('numberFormat') ? '__mixed__' : cellStyle.numberFormat || 'general'} onChange={(event) => { if (event.target.value !== '__mixed__') onEmitStyle({ numberFormat: event.target.value }); }}>
          {mixed('numberFormat') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
          {HOME_NUMBER_FORMAT_OPTIONS.map(({ value, labelKey }) => <option key={value} value={value}>{homeText(locale, labelKey)}</option>)}
        </Select>;
      case 'cells-insert-menu':
      case 'cells-delete-menu':
      case 'cells-format-menu': {
        const icon = controlId === 'cells-insert-menu' ? 'plus' : controlId === 'cells-delete-menu' ? 'trash' : 'columns';
        return <DropdownMenu align="left" trigger={menuTrigger(icon)}><Stack gap="none" className="min-w-[13rem] p-1">{menuMembers(`control.${controlId}`).map((surface) => renderSurface(surface, 'menu'))}</Stack></DropdownMenu>;
      }
      case 'clear-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('trash')}><Stack gap="none" className="min-w-[14rem] p-1">{menuMembers('control.clear-menu').map((surface) => renderSurface(surface, 'menu'))}</Stack></DropdownMenu>;
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
        return <Button aria-label={label} title={label} disabled={disabled} size="sm" variant="ghost" className="w-full justify-start" onClick={actions[controlId]}>{label}</Button>;
      }
      case 'merge-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('columns')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.merge-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
    }
  };

  const renderGroupMenu = (group: HomeGroup) => {
    const label = translateRibbonText(locale, getRibbonGroupDefinition(group).labelKey);
    return <DropdownMenu key={group} align="left" trigger={<Button aria-label={label} title={label} disabled={disabled} icon="more-horizontal" size="sm" variant="ghost" className="h-[68px] min-w-0 flex-1 flex-col gap-1 rounded-none px-1 text-[10px] leading-3">{label}</Button>}>
      <Stack gap="none" className="min-w-[14rem] p-1">{topLevelSurfaces(group).map((surface) => renderSurface(surface, 'menu'))}</Stack>
    </DropdownMenu>;
  };

  const groupLabel = (group: HomeGroup) => translateRibbonText(locale, getRibbonGroupDefinition(group).labelKey);

  return (
    <HomeRibbonLocaleContext.Provider value={locale}>
      {breakpoint !== 'wide' ? (
        <Inline gap="none" className="h-[102px] w-full min-w-0 flex-nowrap items-start overflow-visible" data-testid="home-ribbon-groups" data-ribbon-breakpoint={breakpoint}>
          {HOME_GROUPS.map((group, index) => <React.Fragment key={group}>{index > 0 ? <Divider orientation="vertical" className="h-[96px]" /> : null}{renderGroupMenu(group)}</React.Fragment>)}
        </Inline>
      ) : (
        <Inline gap="none" className="h-[102px] w-full min-w-0 flex-nowrap items-start overflow-visible" data-testid="home-ribbon-groups" data-ribbon-breakpoint={breakpoint}>
          {HOME_GROUPS.map((group, index) => {
            const surfaces = (surfacesByGroup.get(group) ?? []).filter((surface) => !surface.menuId);
            return <React.Fragment key={group}>
              {index > 0 ? <Divider orientation="vertical" className="h-[96px]" /> : null}
              <Stack gap="none" className="h-[102px] min-w-0 flex-1 justify-between overflow-visible px-1">
                <Inline gap="none" className="min-h-0 flex-1 items-start justify-center overflow-visible pt-2">{surfaces.map((surface) => renderSurface(surface, 'wide'))}</Inline>
                <Text size="xs" tone="subtle" className="pointer-events-none h-4 shrink-0 text-center text-[10px] font-medium text-[#5b555a]">{groupLabel(group)}</Text>
              </Stack>
            </React.Fragment>;
          })}
        </Inline>
      )}
    </HomeRibbonLocaleContext.Provider>
  );
}
