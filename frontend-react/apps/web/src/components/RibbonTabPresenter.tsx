import React from 'react';
import { Divider, Inline, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { RIBBON_COMMAND_CATALOG, RIBBON_GROUP_CATALOG, type RibbonCatalogTabId, type RibbonCommandId } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';

export interface RibbonTabPresenterProps {
  tab: RibbonCatalogTabId;
  locale: Locale;
  layout: RibbonLayoutState;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
}

/** Catalog-only presenter used by every non-Home/Insert tab. */
export function RibbonTabPresenter({ tab, locale, layout, renderCommand }: RibbonTabPresenterProps) {
  const groups = RIBBON_GROUP_CATALOG.filter((group) => group.tab === tab).sort((left, right) => left.priority - right.priority);
  return (
    <Inline gap="none" className="h-[102px] min-w-0 flex-nowrap items-start overflow-hidden" data-testid={`ribbon-groups-${tab}`}>
      {groups.map((group, groupIndex) => {
        const commands = RIBBON_COMMAND_CATALOG.filter((command) => command.placements.some((placement) => placement.tab === tab && placement.group === group.id)).sort((left, right) => left.priority - right.priority);
        const collapsed = layout.mode === 'narrow' || (layout.mode === 'compact' && group.priority >= 50);
        return (
          <React.Fragment key={group.id}>
            {groupIndex ? <Divider orientation="vertical" className="h-[96px]" /> : null}
            <Stack gap="none" className="h-[102px] min-w-[76px] shrink-0 justify-between px-1">
              <Inline gap="xs" className="min-h-0 flex-1 flex-nowrap items-start pt-2">
                {commands.map((command) => <React.Fragment key={command.id}>{renderCommand(command.id, collapsed ? { iconOnly: true } : { tile: command.display === 'large' })}</React.Fragment>)}
              </Inline>
              <Text size="xs" tone="subtle" className="h-4 text-center text-[10px] font-medium text-[#5b555a]">{translateRibbonText(locale, group.labelKey)}</Text>
            </Stack>
          </React.Fragment>
        );
      })}
    </Inline>
  );
}
