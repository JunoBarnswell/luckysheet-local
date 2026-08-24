import React, { useMemo } from 'react';
import {
  Box,
  Button,
  ColorPicker,
  Divider,
  DropdownMenu,
  Icon,
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
  type RibbonCommandResult,
} from '@react-sheets/spreadsheet-app';
import { pixelsToPoints, pointsToPixels } from '@react-sheets/exchange-excel-ooxml';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import { homeText } from './home/home-localization';

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
  onMergeCells: () => void;
  onOpenColumnWidth: () => void;
  onAutoFitColumns: () => void;
  onHideColumns: () => void;
  onUnhideColumns: () => void;
  onOpenDefaultColumnWidth: () => void;
}

type Breakpoint = 'wide' | 'compact' | 'narrow';

function breakpointFor(layout: RibbonLayoutState): Breakpoint {
  if (layout.width >= 1280) return 'wide';
  if (layout.width >= 1024) return 'compact';
  return 'narrow';
}

function HomeGroup({ children, group, className }: { children: React.ReactNode; group: 'history' | 'clipboard' | 'font' | 'alignment' | 'number' | 'styles' | 'cells' | 'editing'; className?: string }) {
  const definition = getRibbonGroupDefinition(group);
  const locale = React.useContext(HomeRibbonLocaleContext);
  return (
    <Stack gap="none" className={`h-[102px] min-w-0 shrink-0 justify-between overflow-hidden ${className ?? ''}`}>
      <Box className="min-h-0 flex-1 overflow-hidden">{children}</Box>
      <Text size="xs" tone="subtle" className="pointer-events-none h-4 shrink-0 text-center text-[10px] font-medium text-[#5b555a]">
        {translateRibbonText(locale, definition.labelKey)}
      </Text>
    </Stack>
  );
}

const HomeRibbonLocaleContext = React.createContext<Locale>('en-US');

function Tile({ children, className, ...props }: React.ComponentProps<typeof Button>) {
  return <Button {...props} className={`!h-[68px] !min-h-0 !w-[72px] flex-col gap-1 rounded-none px-1 text-[10px] leading-3 [&>svg]:!h-6 [&>svg]:!w-6 ${className ?? ''}`}>{children}</Button>;
}

/** Home visual composition. It never constructs a workbook mutation itself. */
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
  const surfaces = useMemo(() => new Map([
    ...['history', 'clipboard', 'font', 'alignment', 'number', 'styles', 'cells', 'editing'].flatMap((group) => getRibbonSurfaces('home', group as never, breakpoint)),
  ].map((surface) => [surface.id, surface])), [breakpoint]);
  const command = (surfaceId: string, options?: HomeRibbonCommandOptions) => {
    const id = surfaces.get(surfaceId)?.commandId;
    return id ? renderCommand(id, options) : null;
  };
  const compact = breakpoint === 'compact';
  const dense = layout.width < 1800;
  const narrow = breakpoint === 'narrow';

  if (narrow) {
    const menu = (group: 'history' | 'clipboard' | 'font' | 'alignment' | 'number' | 'styles' | 'cells' | 'editing') => (
      <DropdownMenu
        align="left"
        trigger={<Button size="sm" variant="ghost" icon="more-horizontal">{translateRibbonText(locale, getRibbonGroupDefinition(group).labelKey)}</Button>}
      >
        <Stack gap="none" className="min-w-[13rem] p-1">
          {[...getRibbonSurfaces('home', group, 'wide'), ...getRibbonSurfaces('home', group, 'compact')]
            .filter((surface, index, entries) => surface.commandId && entries.findIndex((entry) => entry.commandId === surface.commandId) === index)
            .flatMap((surface) => surface.commandId ? [renderCommand(surface.commandId, { className: 'w-full justify-start' })] : [])}
        </Stack>
      </DropdownMenu>
    );
    return <Inline gap="xs" className="h-[96px] items-center overflow-hidden">{menu('history')}{menu('clipboard')}{menu('font')}{menu('alignment')}{menu('number')}{menu('styles')}{menu('cells')}{menu('editing')}</Inline>;
  }

  return (
    <HomeRibbonLocaleContext.Provider value={locale}>
      <Inline gap="none" className="h-[102px] flex-nowrap items-start overflow-hidden" data-testid="home-ribbon-groups">
        <HomeGroup className={dense ? 'w-[50px]' : 'w-[70px]'} group="history">
          <Stack gap="none" className="items-center pt-1">
            {command('history.undo', { iconOnly: true, className: 'h-8 w-8 !rounded-none' })}
            {command('history.redo', { iconOnly: true, className: 'h-8 w-8 !rounded-none' })}
          </Stack>
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={dense ? 'w-[92px]' : 'w-[112px]'} group="clipboard">
          <Inline gap="none" className="h-[78px] items-start pt-1">
            <Tile type="button" disabled={disabled} icon="clipboard" className={dense ? '!w-[60px]' : undefined} onClick={() => context.actions.onPaste()} title={homeText(locale, 'pasteAll')}>
              {homeText(locale, 'pasteAll')}
            </Tile>
            <Stack gap="none" className="w-9">
              {command('clipboard.cut', { iconOnly: true, className: '!h-[19px] !min-h-0 !w-8 !rounded-none' })}
              {command('clipboard.copy', { iconOnly: true, className: '!h-[19px] !min-h-0 !w-8 !rounded-none' })}
              <Button
                aria-label={homeText(locale, 'formatPainter')}
                aria-pressed={formatPainterActive}
                data-testid="home-format-painter"
                disabled={!canFormat}
                icon="palette"
                iconOnly
                size="xs"
                title={homeText(locale, 'formatPainterHint')}
                variant="ghost"
                className="!h-[19px] !min-h-0 !w-8 !rounded-none"
                onClick={() => onBeginFormatPainter(false)}
                onDoubleClick={() => onBeginFormatPainter(true)}
              />
              {command('clipboard.paste-special', { iconOnly: true, className: '!h-[19px] !min-h-0 !w-8 !rounded-none' })}
            </Stack>
          </Inline>
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={dense ? 'w-[230px]' : 'w-[260px]'} group="font">
          <Stack gap="xs" className="px-1 pt-2">
            <Inline gap="xs" className="h-8 flex-nowrap items-start">
              <Box className={dense ? 'w-[96px] shrink-0' : 'w-[124px] shrink-0'}>
                <Select aria-label={homeText(locale, 'fontFamily')} className="w-full" disabled={!canFormat} sizeVariant="sm" value={mixed('fontFamily') ? '__mixed__' : cellStyle.fontFamily ?? 'Microsoft YaHei'} onChange={(event) => { if (event.target.value !== '__mixed__') onEmitStyle({ fontFamily: event.target.value }); }}>
                  {mixed('fontFamily') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
                  <option value="Microsoft YaHei">微软雅黑</option><option value="Arial">Arial</option><option value="Calibri">Calibri</option><option value="Segoe UI">Segoe UI</option><option value="Times New Roman">Times New Roman</option>
                </Select>
              </Box>
              <TextInput aria-label={homeText(locale, 'fontSize')} className={dense ? '!w-[42px]' : '!w-[48px]'} disabled={!canFormat} inputMode="decimal" value={mixed('fontSizePx') ? '' : String(Math.round(pixelsToPoints(cellStyle.fontSizePx ?? pointsToPixels(11))))} onChange={(event) => { const value = Number(event.target.value); if (Number.isFinite(value) && value >= 1 && value <= 409) onEmitStyle({ fontSizePx: pointsToPixels(value) }); }} />
              <Button aria-label={homeText(locale, 'increaseFontSize')} disabled={!canFormat} onClick={() => onEmitStyle({ fontSizePx: Math.min(pointsToPixels(409), (cellStyle.fontSizePx ?? pointsToPixels(11)) + pointsToPixels(1)) })} size="sm" variant="ghost" className={`${dense ? '!w-5 text-sm' : '!w-7 text-base'} !h-8 !min-h-0 rounded-none px-0 font-semibold text-[#2572bc]`}>A</Button>
              <Button aria-label={homeText(locale, 'decreaseFontSize')} disabled={!canFormat} onClick={() => onEmitStyle({ fontSizePx: Math.max(pointsToPixels(1), (cellStyle.fontSizePx ?? pointsToPixels(11)) - pointsToPixels(1)) })} size="sm" variant="ghost" className={`${dense ? '!w-5 text-[10px]' : '!w-7 text-xs'} !h-8 !min-h-0 rounded-none px-0 font-semibold text-[#2572bc]`}>A</Button>
            </Inline>
            <Inline gap="none" className="h-8 flex-nowrap items-start">
              {command('font.bold', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}
              {command('font.italic', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}
              {command('font.underline', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}
              {command('font.strikethrough', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              {command('font.borders', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              <DropdownMenu disabled={!canFormat} trigger={<Button aria-label={homeText(locale, 'textColor')} disabled={!canFormat} icon="type" iconOnly size="sm" variant="ghost" className="!h-8 !min-h-0 !w-8 !rounded-none" />}>
                {({ close }) => <ColorPicker color={cellStyle.textColor ?? '#1e293b'} onChange={(color) => { onEmitStyle({ textColor: color }); close(); }} />}
              </DropdownMenu>
              <DropdownMenu disabled={!canFormat} trigger={<Button aria-label={homeText(locale, 'fillBackground')} disabled={!canFormat} icon="paint-bucket" iconOnly size="sm" variant="ghost" className="!h-8 !min-h-0 !w-8 !rounded-none" />}>
                {({ close }) => <ColorPicker color={cellStyle.background ?? '#ffffff'} onChange={(color) => { onEmitStyle({ background: color }); close(); }} />}
              </DropdownMenu>
            </Inline>
          </Stack>
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={dense ? 'w-[205px]' : compact ? 'w-[250px]' : 'w-[315px]'} group="alignment">
          <Inline gap="none" className="h-[78px] px-1 pt-2">
            <Stack gap="none" className="w-[112px]">
              <Inline gap="none">{command('alignment.left', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}{command('alignment.center', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}{command('alignment.right', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}</Inline>
              <Inline gap="none">
                {command('alignment.top', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
                {command('alignment.middle', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
                {command('alignment.bottom', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              </Inline>
            </Stack>
            <Stack gap="none" className={dense ? 'w-[100px]' : compact ? 'w-[128px]' : 'w-[190px]'}>
              <Button aria-label={homeText(locale, 'wrapText')} disabled={!canFormat} icon="wrap-text" size="sm" variant="ghost" className="!h-8 !min-h-0 justify-start rounded-none" onClick={() => onEmitStyle({ wrapText: !cellStyle.wrapText })}>{homeText(locale, 'wrapText')}</Button>
              <Button aria-label={homeText(locale, 'mergeCenter')} disabled={!canFormat} icon="merge-cells" size="sm" variant="ghost" className="!h-8 !min-h-0 justify-start rounded-none" onClick={onMergeCells}>{homeText(locale, 'mergeCenter')}</Button>
              <Inline gap="none" className="h-7">{command('alignment.indent-decrease', { iconOnly: true, className: '!h-7 !min-h-0 !w-8 !rounded-none' })}{command('alignment.indent-increase', { iconOnly: true, className: '!h-7 !min-h-0 !w-8 !rounded-none' })}{command('alignment.orientation', { iconOnly: true, className: '!h-7 !min-h-0 !w-8 !rounded-none' })}</Inline>
            </Stack>
          </Inline>
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={dense ? 'w-[128px]' : 'w-[155px]'} group="number">
          <Stack gap="xs" className="px-1 pt-2">
            <Select aria-label={translateRibbonText(locale, 'groups.number')} disabled={!canFormat} sizeVariant="sm" value={mixed('numberFormat') ? '__mixed__' : cellStyle.numberFormat || 'general'} onChange={(event) => { if (event.target.value !== '__mixed__') onEmitStyle({ numberFormat: event.target.value }); }}>
              {mixed('numberFormat') ? <option value="__mixed__" disabled>{homeText(locale, 'mixed')}</option> : null}
              <option value="general">常规</option><option value="$#,##0">货币</option><option value="0%">百分比</option><option value="#,##0">千位分隔</option><option value="0.00">小数</option>
            </Select>
            <Inline gap="none" className="h-8">
              {command('number.currency', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}
              {command('number.percent', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none text-[#2572bc]' })}
              {command('number.comma', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none px-0 text-[#2572bc]' })}
              <DropdownMenu align="left" trigger={<Button aria-label={homeText(locale, 'decimalFormats')} icon="more-horizontal" iconOnly size="sm" variant="ghost" className="!h-8 !min-h-0 !w-8 !rounded-none" />}>
                <Stack gap="none" className="min-w-[12rem] p-1">{command('number.decimal', { className: 'w-full justify-start' })}{command('number.decimal-decrease', { className: 'w-full justify-start' })}{command('number.decimal-increase', { className: 'w-full justify-start' })}</Stack>
              </DropdownMenu>
            </Inline>
          </Stack>
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={dense ? 'w-[365px]' : compact ? 'w-[290px]' : 'w-[365px]'} group="styles">
          <Inline gap="none" className="h-[78px] items-start px-1 pt-1">
            {command('styles.conditional-format', { tile: true })}
            {command('styles.table', { tile: true })}
            {command('styles.format-cells', { tile: true, testId: 'ribbon-format-cells' })}
            {!compact ? command('styles.template', { tile: true }) : null}
            {!compact ? command('styles.editor', { tile: true }) : null}
            {compact ? (
              <DropdownMenu align="left" trigger={<Tile aria-label={homeText(locale, 'moreStyles')} disabled={disabled} icon="more-horizontal" type="button">{homeText(locale, 'moreStyles')}</Tile>}>
                <Stack gap="none" className="min-w-[13rem] p-1">
                  {command('styles.template', { className: 'w-full justify-start' })}
                  {command('styles.editor', { className: 'w-full justify-start' })}
                </Stack>
              </DropdownMenu>
            ) : null}
          </Inline>
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={compact ? 'w-[72px]' : 'w-[220px]'} group="cells">
          {compact ? (
            <DropdownMenu align="left" trigger={<Button aria-label="Cells" icon="columns" iconOnly size="sm" variant="ghost" className="m-5 !h-10 !w-10 rounded-none" />}>
              <Stack gap="none" className="min-w-[13rem] p-1">
                {command('cells.insert', { className: 'w-full justify-start' })}
                {command('cells.delete', { className: 'w-full justify-start' })}
                {renderCommand('insertColumnHome', { className: 'w-full justify-start' })}
                {renderCommand('deleteColumn', { className: 'w-full justify-start' })}
                <Button size="sm" variant="ghost" className="justify-start" onClick={onOpenColumnWidth}>列宽…</Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={onAutoFitColumns}>自动调整列宽</Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={onHideColumns}>隐藏列</Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={onUnhideColumns}>取消隐藏列</Button>
                <Button size="sm" variant="ghost" className="justify-start" onClick={onOpenDefaultColumnWidth}>默认列宽…</Button>
              </Stack>
            </DropdownMenu>
          ) : (
            <Inline gap="none" className="h-[78px] px-1 pt-1">
              <DropdownMenu align="left" trigger={<Tile aria-label={homeText(locale, 'insert')} icon="plus" type="button">{homeText(locale, 'insert')}</Tile>}>
                <Stack gap="none" className="min-w-[13rem] p-1">{renderCommand('insertRowHome', { className: 'w-full justify-start' })}{renderCommand('insertColumnHome', { className: 'w-full justify-start' })}{renderCommand('shiftCells', { className: 'w-full justify-start' })}</Stack>
              </DropdownMenu>
              <DropdownMenu align="left" trigger={<Tile aria-label={homeText(locale, 'delete')} icon="trash" type="button">{homeText(locale, 'delete')}</Tile>}>
                <Stack gap="none" className="min-w-[13rem] p-1">{renderCommand('deleteRow', { className: 'w-full justify-start' })}{renderCommand('deleteColumn', { className: 'w-full justify-start' })}{renderCommand('shiftCells', { className: 'w-full justify-start' })}</Stack>
              </DropdownMenu>
              <DropdownMenu align="left" trigger={<Tile aria-label={homeText(locale, 'format')} icon="columns" type="button">{homeText(locale, 'format')}</Tile>}>
                <Stack gap="none" className="min-w-[13rem] p-1">
                  <Button size="sm" variant="ghost" className="justify-start" onClick={onOpenColumnWidth}>{homeText(locale, 'columnWidth')}</Button><Button size="sm" variant="ghost" className="justify-start" onClick={onAutoFitColumns}>{homeText(locale, 'autoFitColumnWidth')}</Button><Button size="sm" variant="ghost" className="justify-start" onClick={onHideColumns}>{homeText(locale, 'hideColumns')}</Button><Button size="sm" variant="ghost" className="justify-start" onClick={onUnhideColumns}>{homeText(locale, 'unhideColumns')}</Button><Button size="sm" variant="ghost" className="justify-start" onClick={onOpenDefaultColumnWidth}>{homeText(locale, 'defaultColumnWidth')}</Button>
                </Stack>
              </DropdownMenu>
            </Inline>
          )}
        </HomeGroup>
        <Divider orientation="vertical" className="h-[96px]" />

        <HomeGroup className={compact ? 'w-[118px]' : 'w-[235px]'} group="editing">
          {compact ? (
            <Inline gap="none" className="h-[78px] flex-wrap px-1 pt-1">
              {command('editing.autosum', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              {command('editing.fill-down', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              {command('editing.clear', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              {command('editing.sort', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              {command('editing.filter', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
              {command('editing.find', { iconOnly: true, className: '!h-8 !min-h-0 !w-8 !rounded-none' })}
            </Inline>
          ) : (
            <Inline gap="none" className="h-[78px] flex-wrap content-start px-1 pt-1">
              {command('editing.autosum', { className: '!h-[34px] !min-h-0 !w-[72px] justify-start rounded-none px-1 text-[10px]' })}
              {command('editing.fill-down', { className: '!h-[34px] !min-h-0 !w-[72px] justify-start rounded-none px-1 text-[10px]' })}
              {command('editing.clear', { className: '!h-[34px] !min-h-0 !w-[72px] justify-start rounded-none px-1 text-[10px]' })}
              {command('editing.sort', { className: '!h-[34px] !min-h-0 !w-[72px] justify-start rounded-none px-1 text-[10px]' })}
              {command('editing.filter', { className: '!h-[34px] !min-h-0 !w-[72px] justify-start rounded-none px-1 text-[10px]' })}
              {command('editing.find', { className: '!h-[34px] !min-h-0 !w-[72px] justify-start rounded-none px-1 text-[10px]' })}
            </Inline>
          )}
        </HomeGroup>
      </Inline>
    </HomeRibbonLocaleContext.Provider>
  );
}
