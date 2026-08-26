import React, { type ReactNode } from 'react';
import {
  Button,
  Divider,
  DropdownMenu,
  Inline,
  Stack,
  Text,
  type RibbonLayoutState,
} from '@react-sheets/ui-system';
import {
  DESIGNER_ICON_TO_RIBBON_ICON,
  RIBBON_LAYOUT_SPECS,
  RIBBON_TAB_SURFACES,
  type RibbonCommandId,
  type RibbonGroupId,
  type RibbonLayoutNode,
  type RibbonLayoutSpec,
  type RibbonSurfaceDefinition,
} from '@react-sheets/spreadsheet-app';
import type { Locale } from '../i18n';
import { translateRibbonText } from '../i18n';
import type { HomeRibbonCommandOptions } from './HomeRibbon';

export interface RibbonLayoutRendererProps {
  tab: RibbonLayoutSpec['tab'];
  locale: Locale;
  layout: RibbonLayoutState;
  renderCommand: (id: RibbonCommandId, options?: HomeRibbonCommandOptions) => ReactNode;
  renderSurface?: (surface: RibbonSurfaceDefinition, context: { inMenu: boolean; mode: 'wide' | 'menu' }) => ReactNode;
}

interface NodeRenderContext {
  inMenu: boolean;
}

/**
 * The Designer keeps a stable footprint for each ribbon group so dense
 * groups wrap inside their own area instead of pushing neighboring groups
 * into the viewport or clipping their controls vertically.
 */
const RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  history: 'w-[144px]',
  clipboard: 'w-[216px]',
  font: 'w-[240px]',
  alignment: 'w-[240px]',
  number: 'w-[196px]',
  styles: 'w-[216px]',
  cells: 'w-[216px]',
  editing: 'w-[216px]',
  insertSheets: 'w-[220px]',
  insertTables: 'w-[220px]',
  insertCharts: 'w-[220px]',
  insertDataCharts: 'w-[104px]',
  illustrations: 'w-[288px]',
  insertLinks: 'w-[104px]',
  insertControls: 'w-[152px]',
  sortFilter: 'w-[240px]',
  dataTools: 'w-[232px]',
  findTransform: 'w-[256px]',
  outline: 'w-[224px]',
  whatIf: 'w-[128px]',
};

export function ribbonGroupWidthClass(groupId: RibbonGroupId): string {
  return RIBBON_GROUP_WIDTH_CLASSES[groupId] ?? 'min-w-[76px]';
}

function iconFor(node: { icon: keyof typeof DESIGNER_ICON_TO_RIBBON_ICON }) {
  return DESIGNER_ICON_TO_RIBBON_ICON[node.icon];
}

function commandOptions(node: Extract<RibbonLayoutNode, { kind: 'command' }>, context: NodeRenderContext): HomeRibbonCommandOptions {
  return {
    iconOverride: node.icon ? iconFor(node as { icon: keyof typeof DESIGNER_ICON_TO_RIBBON_ICON }) : undefined,
    iconOnly: false,
    ribbonLayoutNodeId: node.id,
    tile: node.size === 'large' && !context.inMenu,
    className: context.inMenu ? 'w-full justify-start rounded-none' : undefined,
  };
}

function renderLayoutNode(node: RibbonLayoutNode, context: NodeRenderContext, props: RibbonLayoutRendererProps): ReactNode {
  const { renderCommand, renderSurface } = props;
  switch (node.kind) {
    case 'column':
      return <Stack key={node.id} gap="none" className="min-w-0 items-stretch">{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
    case 'row':
      return <Inline key={node.id} gap="none" className="w-full min-w-0 flex-wrap content-start items-start">{node.children.map((child) => renderLayoutNode(child, context, props))}</Inline>;
    case 'stack':
      return <Stack key={node.id} gap="none" className="min-w-0 items-stretch">{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
    case 'command':
      return <React.Fragment key={node.id}>{renderCommand(node.commandId, commandOptions(node, context))}</React.Fragment>;
    case 'surface': {
      const surface = RIBBON_TAB_SURFACES.find((candidate) => candidate.id === node.surfaceId);
      return surface && renderSurface ? <React.Fragment key={node.id}>{renderSurface(surface, { inMenu: context.inMenu, mode: context.inMenu ? 'menu' : 'wide' })}</React.Fragment> : null;
    }
    case 'split':
      return (
        <Inline key={node.id} gap="none" className={context.inMenu ? 'w-full flex-nowrap' : 'flex-nowrap'}>
          {renderCommand(node.primary, { iconOverride: iconFor({ icon: node.primaryIcon }), ribbonLayoutNodeId: node.id, tile: !context.inMenu, className: context.inMenu ? 'min-w-0 flex-1 justify-start rounded-none' : undefined })}
          <DropdownMenu
            align="left"
            trigger={<Button aria-label="More options" data-ribbon-layout-node={`${node.id}.menu`} icon="chevron-down" iconOnly size="sm" variant="ghost" className="!h-8 !w-5 rounded-none px-0" />}
          >
            <Stack gap="none" className="min-w-[12rem] p-1">
              {node.items.map((item) => <React.Fragment key={item.commandId}>{renderCommand(item.commandId, { iconOverride: iconFor(item), ribbonLayoutNodeId: `${node.id}.item.${item.commandId}`, className: 'w-full justify-start rounded-none' })}</React.Fragment>)}
            </Stack>
          </DropdownMenu>
        </Inline>
      );
    case 'dropdown':
      return (
        <DropdownMenu
          key={node.id}
          align="left"
          trigger={renderCommand(node.trigger, { iconOverride: iconFor({ icon: node.triggerIcon }), ribbonLayoutNodeId: node.id, tile: !context.inMenu, className: context.inMenu ? 'w-full justify-start rounded-none' : undefined })}
        >
          <Stack gap="none" className="min-w-[12rem] p-1">
            {node.items.map((item) => <React.Fragment key={item.commandId}>{renderCommand(item.commandId, { iconOverride: iconFor(item), ribbonLayoutNodeId: `${node.id}.item.${item.commandId}`, className: 'w-full justify-start rounded-none' })}</React.Fragment>)}
          </Stack>
        </DropdownMenu>
      );
    case 'checkbox':
    case 'spinner':
    case 'combo':
    case 'launcher':
      return <React.Fragment key={node.id}>{renderCommand(node.commandId, { iconOverride: iconFor(node), ribbonLayoutNodeId: node.id, className: context.inMenu ? 'w-full justify-start rounded-none' : undefined })}</React.Fragment>;
    case 'separator':
      return <Divider key={node.id} orientation="vertical" className="mx-0.5 h-8" />;
  }
}

export function RibbonLayoutRenderer(props: RibbonLayoutRendererProps): React.ReactElement {
  const { tab, locale, layout } = props;
  const spec = RIBBON_LAYOUT_SPECS[tab];
  return (
    <Inline gap="none" className="h-[102px] w-max min-w-full flex-nowrap items-start overflow-visible" data-testid={tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`} data-ribbon-layout={tab} data-ribbon-breakpoint={layout.mode}>
      {spec.groups.map((group, index) => {
        const collapsed = layout.mode === 'narrow' || (layout.mode === 'compact' && group.collapsePriority >= 50);
        const groupLabel = translateRibbonText(locale, `groups.${group.id}`);
        const content = group.children.map((node) => renderLayoutNode(node, { inMenu: collapsed }, props));
        return (
          <React.Fragment key={group.id}>
            {index > 0 ? <Divider orientation="vertical" className="h-[96px]" /> : null}
            {collapsed ? (
              <DropdownMenu
                align="left"
                trigger={<Button aria-label={groupLabel} data-ribbon-group={group.id} title={groupLabel} icon="chevron-down" size="sm" variant="ghost" className="h-[68px] min-w-[76px] shrink-0 flex-col gap-1 rounded-none px-1 text-[10px] leading-3">{groupLabel}</Button>}
              >
                <Stack gap="none" className="min-w-[14rem] p-1">{content}</Stack>
              </DropdownMenu>
            ) : (
              <Stack data-ribbon-group={group.id} gap="none" className={`h-[102px] shrink-0 justify-between overflow-hidden px-1 ${ribbonGroupWidthClass(group.id)}`}>
                <Inline gap="none" className="min-h-0 flex-1 flex-wrap content-start items-start pt-2">{content}</Inline>
                <Text size="xs" tone="subtle" className="h-4 text-center text-[10px] font-medium text-[#5b555a]">{groupLabel}</Text>
              </Stack>
            )}
          </React.Fragment>
        );
      })}
    </Inline>
  );
}
