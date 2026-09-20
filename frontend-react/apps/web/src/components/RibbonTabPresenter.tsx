import React from 'react';
import { Box, Button, Divider, DropdownMenu, Icon, Inline, RIBBON_DENSITY_CLASSES, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { RIBBON_COMMAND_CATALOG, RIBBON_GROUP_CATALOG, type RibbonCatalogTabId, type RibbonCommandId } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { collapsedRibbonGroups, RIBBON_GROUP_ICONS } from './RibbonLayoutRenderer';

export interface RibbonTabPresenterProps {
  tab: RibbonCatalogTabId;
  locale: Locale;
  layout: RibbonLayoutState;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
}

/** Catalog-only presenter used by every non-Home/Insert tab. */
export function RibbonTabPresenter({ tab, locale, layout, renderCommand }: RibbonTabPresenterProps) {
  const groups = RIBBON_GROUP_CATALOG.filter((group) => group.tab === tab).sort((left, right) => left.priority - right.priority);
  const entries = groups.map(group => ({ ...group, commands: RIBBON_COMMAND_CATALOG.filter(command => command.placements.some(placement => placement.tab === tab && placement.group === group.id)).sort((left, right) => left.priority - right.priority) }));
  const collapsed = collapsedRibbonGroups(entries.map(group => ({ id: group.id, width: Math.ceil(group.commands.length / 3) * 156 + 12 })), layout.width);
  return (
    <Inline gap="none" tabIndex={0} className={`${RIBBON_DENSITY_CLASSES.commandArea} w-full min-w-0 flex-nowrap items-start overflow-x-auto overflow-y-hidden [scrollbar-width:thin]`} data-testid={`ribbon-groups-${tab}`} data-ribbon-breakpoint={layout.mode}>
      {entries.map((group, groupIndex) => {
        const label = translateRibbonText(locale, group.labelKey);
        const inMenu = collapsed.has(group.id);
        const commands = group.commands.map(command => <React.Fragment key={command.id}>{renderCommand(command.id, { tile: false, className: '!h-6 !min-h-0 w-full justify-start overflow-hidden !px-1.5 text-xs' })}</React.Fragment>);
        return <React.Fragment key={group.id}>
          {groupIndex ? <Divider orientation="vertical" className="my-3 h-[72px] border-slate-200" /> : null}
          {inMenu ? <Stack gap="none" className="h-[104px] w-[68px] shrink-0 justify-center px-1">
            <DropdownMenu align="left" trigger={<Button aria-label={`${label}工具`} title={label} variant="ghost" size="sm" className="h-[76px] !w-[60px] !min-w-0 flex-col gap-2 !px-1 text-xs"><Icon name={RIBBON_GROUP_ICONS[group.id] ?? 'grid'} size="lg" /><Text className="max-w-full truncate">{label}</Text><Icon name="chevron-down" size="xs" /></Button>}>
              <Stack gap="xs" className="max-h-[60vh] min-w-[14rem] overflow-y-auto p-3"><Text size="xs" tone="muted">{label}</Text>{commands}</Stack>
            </DropdownMenu>
          </Stack> : <Stack gap="none" className="h-[104px] shrink-0 justify-between px-1.5 py-1" data-ribbon-group={group.id}>
            <Box className="grid h-[80px] auto-cols-[156px] grid-flow-col grid-rows-3 items-center">{commands}</Box>
            <Text size="xs" tone="subtle" className={`${RIBBON_DENSITY_CLASSES.groupCaption} text-center text-[10px] text-slate-500`}>{label}</Text>
          </Stack>}
        </React.Fragment>;
      })}
    </Inline>
  );
}
