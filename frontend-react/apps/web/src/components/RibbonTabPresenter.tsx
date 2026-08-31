import React from 'react';
import { Divider, Inline, RIBBON_DENSITY_CLASSES, Stack, Text, type RibbonLayoutState } from '@react-sheets/ui-system';
import { getRibbonCommandDefinition, getRibbonGroupDefinition, type CompiledFeatureSurfaceSchema, type RibbonCatalogTabId, type RibbonCommandId, type RibbonGroupId } from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';
import { ribbonGroupWidthClass } from './RibbonLayoutRenderer';

export interface RibbonTabPresenterProps {
  tab: RibbonCatalogTabId;
  locale: Locale;
  layout: RibbonLayoutState;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => React.ReactNode;
  featureSurfaceSchema: CompiledFeatureSurfaceSchema;
}

/** Feature-manifest surface presenter used by every non-Home/Insert tab. */
export function RibbonTabPresenter({ tab, locale, layout, renderCommand, featureSurfaceSchema }: RibbonTabPresenterProps) {
  const surfaces = featureSurfaceSchema.ribbon.filter((surface) => surface.tab === tab && surface.commandId)
    .concat(featureSurfaceSchema.contextualTabs.filter((surface) => surface.tab === tab && surface.commandId));
  const groups = [...new Set(surfaces.map((surface) => surface.group as RibbonGroupId))]
    .map((id) => getRibbonGroupDefinition(id))
    .sort((left, right) => left.priority - right.priority);
  return (
    <Inline gap="none" tabIndex={0} className={`${RIBBON_DENSITY_CLASSES.commandArea} w-full min-w-0 flex-nowrap items-start overflow-x-auto overflow-y-hidden [scrollbar-width:thin]`} data-testid={`ribbon-groups-${tab}`} data-ribbon-breakpoint={layout.mode}>
      {groups.map((group, groupIndex) => {
        const commands = surfaces
          .filter((surface) => surface.group === group.id && surface.commandId)
          .map((surface) => ({ surface, definition: getRibbonCommandDefinition(surface.commandId as RibbonCommandId) }))
          .sort((left, right) => left.surface.order - right.surface.order);
        return (
          <React.Fragment key={group.id}>
            {groupIndex ? <Divider orientation="vertical" className={RIBBON_DENSITY_CLASSES.groupContent} /> : null}
            <Stack gap="none" className={`${RIBBON_DENSITY_CLASSES.groupContent} min-w-0 shrink-0 justify-between overflow-hidden px-1 ${ribbonGroupWidthClass(group.id, layout.mode, layout.width)}`}>
              <Inline gap="xs" className={`${RIBBON_DENSITY_CLASSES.groupControls} min-h-0 flex-wrap content-center items-center`}>
                {commands.map(({ surface, definition }) => <React.Fragment key={surface.id}>{renderCommand(surface.commandId as RibbonCommandId, { tile: definition.display === 'large', ribbonSurfaceId: surface.id })}</React.Fragment>)}
              </Inline>
              <Text size="xs" tone="subtle" className={`${RIBBON_DENSITY_CLASSES.groupCaption} text-center text-[10px] font-medium text-[#5b555a]`}>{translateRibbonText(locale, group.labelKey)}</Text>
            </Stack>
          </React.Fragment>
        );
      })}
    </Inline>
  );
}
