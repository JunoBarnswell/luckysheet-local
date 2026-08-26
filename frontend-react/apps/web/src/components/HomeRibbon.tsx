import React from 'react';
import {
  Box,
  Button,
  ColorPicker,
  DropdownMenu,
  Inline,
  Select,
  Stack,
  TextInput,
  type RibbonLayoutState,
} from '@react-sheets/ui-system';
import {
  getRibbonSurfaces,
  type HomeRibbonState,
  type RibbonCommandContext,
  type RibbonCommandId,
  type RibbonControlId,
  type RibbonSurfaceDefinition,
  type RibbonMergeOperation,
} from '@react-sheets/spreadsheet-app';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import { HOME_NUMBER_FORMAT_OPTIONS, homeText } from './home/home-localization';
import { FontFamilyControl } from './FontFamilyControl';
import { RibbonLayoutRenderer } from './RibbonLayoutRenderer';

export interface HomeRibbonCommandOptions {
  className?: string;
  iconOnly?: boolean;
  iconOverride?: import('@react-sheets/ui-system').IconName;
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
  onOpenRowHeight: () => void;
  onAutoFitRows: () => void;
  onHideRows: () => void;
  onUnhideRows: () => void;
}

type HomeGroup = 'history' | 'clipboard' | 'font' | 'alignment' | 'number' | 'styles' | 'cells' | 'editing';

const HOME_GROUPS: readonly HomeGroup[] = ['history', 'clipboard', 'font', 'alignment', 'number', 'styles', 'cells', 'editing'];

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
    case 'cell-styles-menu': return translateRibbonText(locale, 'groups.styles');
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
      .flatMap((candidate) => getRibbonSurfaces('home', group, candidate))
      .filter((surface) => {
        if (seen.has(surface.id)) return false;
        seen.add(surface.id);
        return true;
      });
  };
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
        return <Box className={mode === 'wide' ? 'w-[124px] shrink-0' : 'w-full'}><FontFamilyControl
          value={cellStyle.fontFamily}
          fallbackValue="Microsoft YaHei"
          mixed={mixed('fontFamily')}
          mixedPlaceholder={homeText(locale, 'mixed')}
          className="w-full"
          disabled={!canFormat}
          label={label}
          testId="home-font-family"
          onCommit={(fontFamily) => onEmitStyle({ fontFamily })}
        /></Box>;
      case 'font-size':
        return <TextInput
          aria-label={label}
          className={mode === 'wide' ? '!w-[48px]' : 'w-full'}
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
        return <Button aria-label={label} disabled={!canFormat} size="sm" variant="ghost" className={mode === 'wide' ? '!h-8 !min-h-0 !w-7 rounded-none px-0 font-semibold text-[#2572bc]' : 'w-full justify-start'} onClick={() => onEmitStyle({ fontSizePx: controlId === 'font-increase' ? Math.min(pointsToPixels(409), (cellStyle.fontSizePx ?? pointsToPixels(11)) + pointsToPixels(1)) : Math.max(pointsToPixels(1), (cellStyle.fontSizePx ?? pointsToPixels(11)) - pointsToPixels(1)) })}>{mode === 'wide' ? 'A' : label}</Button>;
      case 'font-color':
      case 'fill-color':
        return <DropdownMenu disabled={!canFormat} trigger={mode === 'wide' ? <Button aria-label={label} disabled={!canFormat} icon={controlId === 'font-color' ? 'type' : 'paint-bucket'} iconOnly size="sm" variant="ghost" className="!h-8 !min-h-0 !w-8 !rounded-none" /> : menuTrigger(controlId === 'font-color' ? 'type' : 'paint-bucket')}>
          {({ close }) => <ColorPicker color={controlId === 'font-color' ? cellStyle.textColor ?? '#1e293b' : cellStyle.background ?? '#ffffff'} onChange={(color) => { onEmitStyle({ [controlId === 'font-color' ? 'textColor' : 'background']: color }); close(); }} />}
        </DropdownMenu>;
      case 'font-borders-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('borders')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.font-borders-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
      case 'number-format':
        return <Select aria-label={label} className="w-full" disabled={!canFormat} sizeVariant="sm" value={mixed('numberFormat') ? '__mixed__' : cellStyle.numberFormat || 'general'} onChange={(event) => { if (event.target.value !== '__mixed__') onEmitStyle({ numberFormat: event.target.value }); }}>
          {mixed('numberFormat') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
          {HOME_NUMBER_FORMAT_OPTIONS.map(({ value, labelKey }) => <option key={value} value={value}>{homeText(locale, labelKey)}</option>)}
        </Select>;
      case 'alignment-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('align-left')}>
          <Stack gap="none" className="min-w-[15rem] p-1">{menuMembers('control.alignment-menu').map((surface) => renderSurface(surface, 'menu'))}</Stack>
        </DropdownMenu>;
      case 'orientation-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('type')}>
          <Stack gap="none" className="min-w-[15rem] p-1">{menuMembers('control.orientation-menu').map((surface) => renderSurface(surface, 'menu'))}</Stack>
        </DropdownMenu>;
      case 'cells-insert-menu':
      case 'cells-delete-menu':
      case 'cells-format-menu': {
        const icon = controlId === 'cells-insert-menu' ? 'plus' : controlId === 'cells-delete-menu' ? 'trash' : 'columns';
        return <DropdownMenu align="left" trigger={menuTrigger(icon)}><Stack gap="none" className="min-w-[13rem] p-1">{menuMembers(`control.${controlId}`).map((surface) => renderSurface(surface, 'menu'))}</Stack></DropdownMenu>;
      }
      case 'cell-styles-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('star')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.cell-styles-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
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
        return <Button aria-label={label} title={label} disabled={disabled} size="sm" variant="ghost" className="w-full justify-start" onClick={actions[controlId]}>{label}</Button>;
      }
      case 'merge-menu':
        return <DropdownMenu align="left" trigger={menuTrigger('columns')}>
          <Stack gap="none" className="min-w-[14rem] p-1">
            {menuMembers('control.merge-menu').map((surface) => renderSurface(surface, 'menu'))}
          </Stack>
        </DropdownMenu>;
      case 'auto-sum-menu':
        return <Inline gap="none" className={mode === 'wide' ? 'h-8 items-stretch' : 'w-full'}>
          {renderCommand('autoSum', { iconOnly: mode === 'wide', className: mode === 'wide' ? '!h-8 !min-h-0 !rounded-none' : 'w-full justify-start' })}
          <DropdownMenu align="left" trigger={<Button aria-label={`${label} options`} title={`${label} options`} disabled={disabled} icon="chevron-down" iconOnly size="sm" variant="ghost" className={mode === 'wide' ? '!h-8 !min-h-0 !w-5 rounded-none px-0' : 'w-5 shrink-0 justify-center px-0'} />}>
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
        renderCommand={renderCommand}
        renderSurface={(surface, context) => renderSurface(surface, context.mode)}
      />
    </HomeRibbonLocaleContext.Provider>
  );
}
