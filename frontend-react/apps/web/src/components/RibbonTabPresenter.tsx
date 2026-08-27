import React from 'react';
import { Divider, Inline, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { RIBBON_COMMAND_CATALOG, RIBBON_GROUP_CATALOG, type RibbonCatalogTabId, type RibbonCommandId } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { ribbonGroupWidthClass } from './RibbonLayoutRenderer';

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
    <Inline gap="none" className="h-[86px] w-full min-w-0 flex-nowrap items-start overflow-hidden" data-testid={`ribbon-groups-${tab}`} data-ribbon-breakpoint={layout.mode}>
      {groups.map((group, groupIndex) => {
        const commands = RIBBON_COMMAND_CATALOG.filter((command) => command.placements.some((placement) => placement.tab === tab && placement.group === group.id)).sort((left, right) => left.priority - right.priority);
        return (
          <React.Fragment key={group.id}>
            {groupIndex ? <Divider orientation="vertical" className="h-[82px]" /> : null}
            <Stack gap="none" className={`h-[86px] min-w-0 shrink-0 justify-between overflow-hidden px-1 ${ribbonGroupWidthClass(group.id, layout.mode, layout.width)}`}>
              <Inline gap="xs" className="h-[72px] min-h-0 flex-wrap content-center items-center">
                {commands.map((command) => <React.Fragment key={command.id}>{renderCommand(command.id, { tile: command.display === 'large' })}</React.Fragment>)}
              </Inline>
              <Text size="xs" tone="subtle" className="h-[14px] text-center text-[10px] font-medium leading-[14px] text-[#5b555a]">{translateRibbonText(locale, group.labelKey)}</Text>
            </Stack>
          </React.Fragment>
        );
      })}
    </Inline>
  );
}
