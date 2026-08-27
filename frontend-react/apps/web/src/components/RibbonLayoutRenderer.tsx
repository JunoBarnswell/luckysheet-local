import React, { type ReactNode } from 'react';
import {
  Box,
  Button,
  Divider,
  DropdownMenu,
  Inline,
  Stack,
  Text,
  RIBBON_DENSITY_CLASSES,
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
  history: 'w-[56px]',
  clipboard: 'w-[124px]',
  font: 'w-[224px]',
  alignment: 'w-[248px]',
  number: 'w-[154px]',
  styles: 'w-[260px]',
  cells: 'w-[144px]',
  editing: 'w-[196px]',
  pageSetup: 'w-[300px]',
  scaleToFit: 'w-[112px]',
  sheetOptions: 'w-[220px]',
  calculation: 'w-[190px]',
  functionLibrary: 'w-[220px]',
  formulaAudit: 'w-[276px]',
  definedNames: 'w-[128px]',
  insertSheets: 'w-[152px]',
  insertTables: 'w-[152px]',
  insertCharts: 'w-[240px]',
  insertDataCharts: 'w-[72px]',
  illustrations: 'w-[200px]',
  insertLinks: 'w-[68px]',
  insertControls: 'w-[108px]',
  sortFilter: 'w-[220px]',
  dataTools: 'w-[292px]',
  findTransform: 'w-[292px]',
  outline: 'w-[400px]',
  whatIf: 'w-[160px]',
};

const COMPACT_RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  ...WIDE_RIBBON_GROUP_WIDTH_CLASSES,
};

const DENSE_COMPACT_RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  ...WIDE_RIBBON_GROUP_WIDTH_CLASSES,
};

const HOME_RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  history: 'w-[64px]',
  clipboard: 'w-[118px]',
  font: 'w-[205px]',
  alignment: 'w-[180px]',
  number: 'w-[120px]',
  styles: 'w-[166px]',
  cells: 'w-[118px]',
  editing: 'w-[204px]',
};

export function ribbonGroupWidthClass(groupId: RibbonGroupId, mode: RibbonLayoutState['mode'] = 'wide', width = 0, tab?: RibbonLayoutSpec['tab']): string {
  if (tab === 'home' && mode !== 'narrow') return HOME_RIBBON_GROUP_WIDTH_CLASSES[groupId] ?? 'min-w-[72px] flex-1';
  const widths = mode === 'wide'
    ? WIDE_RIBBON_GROUP_WIDTH_CLASSES
    : width >= 1440 ? DENSE_COMPACT_RIBBON_GROUP_WIDTH_CLASSES : COMPACT_RIBBON_GROUP_WIDTH_CLASSES;
  return widths[groupId] ?? (mode === 'wide' ? 'w-[112px]' : 'min-w-[72px] flex-1');
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
      return <Stack key={node.id} gap="none" className="min-w-0 items-center justify-center">{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
    case 'row':
      return <Inline key={node.id} gap="none" className="min-w-0 flex-nowrap items-center content-center">{node.children.map((child) => renderLayoutNode(child, context, props))}</Inline>;
    case 'stack':
      return <Stack key={node.id} gap="none" className="min-w-0 items-center justify-center">{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
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
  const isHome = tab === 'home';
  return (
    <Inline aria-label={`${tab} ribbon commands`} gap="none" tabIndex={0} className={`${RIBBON_DENSITY_CLASSES.commandArea} ${isHome ? 'px-3' : ''} w-full min-w-0 flex-nowrap items-start overflow-x-auto overflow-y-hidden [scrollbar-width:thin]`} data-testid={tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`} data-ribbon-layout={tab} data-ribbon-breakpoint={layout.mode}>
      {spec.groups.map((group, index) => {
        const groupLabel = translateRibbonText(locale, `groups.${group.id}`);
        const content = group.children.map((node) => renderLayoutNode(node, { inMenu: false, tab }, props));
        return (
          <React.Fragment key={group.id}>
            {index > 0 ? isHome
              ? <Box className={`${RIBBON_DENSITY_CLASSES.groupContent} relative mx-1.5 w-0 shrink-0`}><Divider orientation="vertical" className="absolute left-0 h-full" /></Box>
              : <Divider orientation="vertical" className={RIBBON_DENSITY_CLASSES.groupContent} /> : null}
            <Stack data-ribbon-group={group.id} gap="none" className={`${RIBBON_DENSITY_CLASSES.groupContent} min-w-0 shrink-0 justify-between overflow-hidden px-1 ${ribbonGroupWidthClass(group.id, layout.mode, layout.width, tab)}`}>
              <Inline gap="none" className={`${RIBBON_DENSITY_CLASSES.groupControls} min-h-0 flex-nowrap items-center justify-center content-center`}>{content}</Inline>
              <Text size="xs" tone="subtle" className={`${RIBBON_DENSITY_CLASSES.groupCaption} ${isHome ? 'text-[10px] font-normal text-[#999]' : 'font-medium text-[#5b555a]'} shrink-0 truncate text-center select-none`}>{groupLabel}</Text>
            </Stack>
          </React.Fragment>
        );
      })}
    </Inline>
  );
}
