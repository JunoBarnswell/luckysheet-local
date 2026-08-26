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
  tab: RibbonLayoutSpec['tab'];
}

/**
 * The Designer keeps a stable footprint for each ribbon group so dense
 * groups wrap inside their own area instead of pushing neighboring groups
 * into the viewport or clipping their controls vertically.
 */
const WIDE_RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  history: 'w-[64px]',
  clipboard: 'w-[138px]',
  font: 'w-[248px]',
  alignment: 'w-[260px]',
  number: 'w-[160px]',
  styles: 'w-[460px]',
  cells: 'w-[170px]',
  editing: 'w-[190px]',
  pageSetup: 'w-[300px]',
  scaleToFit: 'w-[112px]',
  sheetOptions: 'w-[220px]',
  calculation: 'w-[190px]',
  functionLibrary: 'w-[220px]',
  formulaAudit: 'w-[276px]',
  definedNames: 'w-[128px]',
  insertSheets: 'w-[204px]',
  insertTables: 'w-[204px]',
  insertCharts: 'w-[204px]',
  insertDataCharts: 'w-[92px]',
  illustrations: 'w-[264px]',
  insertLinks: 'w-[92px]',
  insertControls: 'w-[136px]',
  sortFilter: 'w-[220px]',
  dataTools: 'w-[292px]',
  findTransform: 'w-[292px]',
  outline: 'w-[292px]',
  whatIf: 'w-[160px]',
};

const COMPACT_RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  history: 'w-[48px]',
  clipboard: 'w-[96px]',
  font: 'w-[178px]',
  alignment: 'w-[202px]',
  number: 'w-[122px]',
  styles: 'w-[238px]',
  cells: 'w-[92px]',
  editing: 'w-[150px]',
  pageSetup: 'w-[260px]',
  scaleToFit: 'w-[88px]',
  sheetOptions: 'w-[192px]',
  calculation: 'w-[156px]',
  functionLibrary: 'w-[180px]',
  formulaAudit: 'w-[236px]',
  definedNames: 'w-[96px]',
  insertSheets: 'w-[108px]',
  insertTables: 'w-[108px]',
  insertCharts: 'w-[108px]',
  insertDataCharts: 'w-[72px]',
  illustrations: 'w-[140px]',
  insertLinks: 'w-[72px]',
  insertControls: 'w-[84px]',
  sortFilter: 'w-[180px]',
  dataTools: 'w-[240px]',
  findTransform: 'w-[240px]',
  outline: 'w-[240px]',
  whatIf: 'w-[112px]',
};

export function ribbonGroupWidthClass(groupId: RibbonGroupId, mode: RibbonLayoutState['mode'] = 'wide'): string {
  return (mode === 'wide' ? WIDE_RIBBON_GROUP_WIDTH_CLASSES : COMPACT_RIBBON_GROUP_WIDTH_CLASSES)[groupId]
    ?? (mode === 'wide' ? 'w-[112px]' : 'min-w-[72px] flex-1');
}

function iconFor(node: { icon: keyof typeof DESIGNER_ICON_TO_RIBBON_ICON }) {
  return DESIGNER_ICON_TO_RIBBON_ICON[node.icon];
}

function commandOptions(node: Extract<RibbonLayoutNode, { kind: 'command' }>, context: NodeRenderContext): HomeRibbonCommandOptions {
  const compactClass = context.tab === 'data' ? '!h-6 !min-h-0 gap-1 px-1 text-[11px]' : '!h-7 !min-h-0';
  return {
    iconOverride: node.icon ? iconFor(node as { icon: keyof typeof DESIGNER_ICON_TO_RIBBON_ICON }) : undefined,
    iconOnly: false,
    ribbonLayoutNodeId: node.id,
    tile: node.size === 'large' && !context.inMenu,
    className: context.inMenu ? 'w-full justify-start rounded-none' : node.size === 'small' ? compactClass : undefined,
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
          {renderCommand(node.primary, { iconOverride: iconFor({ icon: node.primaryIcon }), ribbonLayoutNodeId: node.id, className: context.inMenu ? 'min-w-0 flex-1 justify-start rounded-none' : context.tab === 'data' ? '!h-6 !min-h-0 gap-1 px-1 text-[11px]' : '!h-7 !min-h-0' })}
          <DropdownMenu
            align="left"
            trigger={<Button aria-label="More options" data-ribbon-layout-node={`${node.id}.menu`} icon="chevron-down" iconOnly size="sm" variant="ghost" className="!h-7 !w-5 rounded-none px-0" />}
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
          trigger={renderCommand(node.trigger, { iconOverride: iconFor({ icon: node.triggerIcon }), ribbonLayoutNodeId: node.id, className: context.inMenu ? 'w-full justify-start rounded-none' : context.tab === 'data' ? '!h-6 !min-h-0 gap-1 px-1 text-[11px]' : '!h-7 !min-h-0' })}
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
    <Inline gap="none" className="h-[115px] w-full min-w-0 flex-nowrap items-start overflow-hidden" data-testid={tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`} data-ribbon-layout={tab} data-ribbon-breakpoint={layout.mode}>
      {spec.groups.map((group, index) => {
        const groupLabel = translateRibbonText(locale, `groups.${group.id}`);
        const content = group.children.map((node) => renderLayoutNode(node, { inMenu: false, tab }, props));
        return (
          <React.Fragment key={group.id}>
            {index > 0 ? <Divider orientation="vertical" className="h-[108px]" /> : null}
            <Stack data-ribbon-group={group.id} gap="none" className={`h-[115px] min-w-0 shrink-0 justify-between overflow-hidden px-1 ${ribbonGroupWidthClass(group.id, layout.mode)}`}>
              <Inline gap="none" className="min-h-0 flex-1 flex-wrap content-start items-start pt-2">{content}</Inline>
              <Text size="xs" tone="subtle" className="h-4 shrink-0 truncate text-center text-[10px] font-medium text-[#5b555a]">{groupLabel}</Text>
            </Stack>
          </React.Fragment>
        );
      })}
    </Inline>
  );
}
