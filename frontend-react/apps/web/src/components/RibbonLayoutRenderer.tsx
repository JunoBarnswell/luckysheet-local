import React, { type ReactNode } from 'react';
import {
  Box,
  Button,
  Icon,
  type IconName,
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
  renderSurface?: (surface: RibbonSurfaceDefinition, context: { inMenu: boolean; mode: RibbonLayoutState['mode'] | 'menu' }) => ReactNode;
}

interface NodeRenderContext {
  inMenu: boolean;
  tab: RibbonLayoutSpec['tab'];
}

const RIBBON_GROUP_WIDTHS: Partial<Record<RibbonGroupId, number>> = {
  pageSetup: 420, scaleToFit: 100, sheetOptions: 300,
  calculation: 250, functionLibrary: 220, formulaAudit: 280, definedNames: 200,
  tables: 248, illustrations: 324, controls: 68, charts: 216,
  sparklines: 188, filters: 140, links: 68, insertComments: 80, text: 320, symbols: 140,
  sortFilter: 280, dataTools: 340, findTransform: 300, outline: 440, whatIf: 180,
};

/** Reserve every group entry before expanding the highest-priority groups. */
export function collapsedRibbonGroups(groups: readonly { id: RibbonGroupId; width: number }[], width: number, preferred: readonly RibbonGroupId[] = []): Set<RibbonGroupId> {
  const collapsed = new Set(groups.map(group => group.id));
  let remaining = width - groups.length * 69 - 12;
  const ordered = [...groups].sort((left, right) => {
    const priority = (id: RibbonGroupId) => preferred.includes(id) ? preferred.indexOf(id) : preferred.length + groups.findIndex(group => group.id === id);
    return priority(left.id) - priority(right.id);
  });
  for (const group of ordered) {
    const extra = Math.max(0, group.width - 68);
    if (extra <= remaining) { collapsed.delete(group.id); remaining -= extra; }
  }
  return collapsed;
}

export function ribbonGroupWidth(groupId: RibbonGroupId): number {
  return RIBBON_GROUP_WIDTHS[groupId] ?? 220;
}

const HOME_RIBBON_GROUP_WIDTH_CLASSES: Partial<Record<RibbonGroupId, string>> = {
  clipboard: 'w-[140px]', font: 'w-[244px]', alignment: 'w-[224px]',
  number: 'w-[140px]', styles: 'w-[188px]', cells: 'w-[164px]', editing: 'w-[210px]',
};

export const RIBBON_GROUP_ICONS: Partial<Record<RibbonGroupId, IconName>> = {
  styles: 'file-spreadsheet', cells: 'grid', editing: 'search', alignment: 'align-left', number: 'calculator',
  tables: 'table', illustrations: 'picture', controls: 'checkbox', charts: 'chart', sparklines: 'sparkline', filters: 'filter', links: 'link', insertComments: 'comment', text: 'file-spreadsheet', symbols: 'calculator',
};

function collapseHomeGroup(groupId: RibbonGroupId, width: number): boolean {
  return (width < 1500 && (groupId === 'styles' || groupId === 'cells'))
    || (width < 1320 && groupId === 'editing')
    || (width < 1120 && groupId === 'alignment')
    || (width < 980 && groupId === 'number');
}

export function ribbonGroupWidthClass(groupId: RibbonGroupId, _mode: RibbonLayoutState['mode'] = 'wide', _width = 0, tab?: RibbonLayoutSpec['tab']): string {
  if (tab === 'home') {
    const widths = HOME_RIBBON_GROUP_WIDTH_CLASSES;
    return widths[groupId] ?? 'min-w-[72px] flex-1';
  }
  return 'shrink-0';
}

function iconFor(node: { icon: keyof typeof DESIGNER_ICON_TO_RIBBON_ICON }) {
  return DESIGNER_ICON_TO_RIBBON_ICON[node.icon];
}

function commandOptions(node: Extract<RibbonLayoutNode, { kind: 'command' }>, context: NodeRenderContext): HomeRibbonCommandOptions {
  const compactClass = context.tab === 'data' ? '!h-6 !min-h-0 gap-1 px-1 text-[11px]' : '!h-6 !min-h-0';
  return {
    iconOverride: node.icon ? iconFor(node as { icon: keyof typeof DESIGNER_ICON_TO_RIBBON_ICON }) : undefined,
    iconOnly: false,
    ribbonLayoutNodeId: node.id,
    tile: node.size === 'large' && !context.inMenu,
    className: context.inMenu ? 'w-full justify-start rounded-none' : node.size === 'small' ? compactClass : undefined,
  };
}

const HOME_COLUMN_CLASSES: Readonly<Record<string, string>> = {
  'clipboard.secondary': 'gap-0.5 items-start justify-center',
  'font.layout': 'gap-1 items-start justify-center',
  'alignment.controls': 'gap-1 items-start justify-center',
  'alignment.wrap-merge': 'gap-1 items-start justify-center',
  'number.layout': 'gap-1 items-start justify-center',
  'editing.stack': 'gap-0.5 items-start justify-center',
  'editing.search-stack': 'gap-1 items-start justify-center',
};

const HOME_ROW_CLASSES: Readonly<Record<string, string>> = {
  'clipboard.layout': 'h-[80px] gap-1',
  'font.controls': 'gap-1',
  'font.actions': 'gap-1',
  'alignment.layout': 'h-[80px] gap-2',
  'alignment.controls.top': 'gap-1',
  'alignment.controls.bottom': 'gap-1',
  'number.actions': 'gap-1',
  'styles.actions': 'h-[80px] gap-1',
  'cells.actions': 'h-[80px] gap-1',
  'editing.layout': 'h-[80px] w-full gap-1',
};

function renderLayoutNode(node: RibbonLayoutNode, context: NodeRenderContext, props: RibbonLayoutRendererProps): ReactNode {
  const { renderCommand, renderSurface } = props;
  switch (node.kind) {
    case 'column':
      return <Stack key={node.id} gap="none" className={`min-w-0 items-center justify-center ${context.tab === 'home' && !context.inMenu ? HOME_COLUMN_CLASSES[node.id] ?? '' : ''}`}>{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
    case 'row':
      if (context.inMenu) return <Stack key={node.id} gap="xs" className="w-full items-stretch">{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
      return <Inline key={node.id} gap="none" className={`min-w-0 flex-nowrap items-center content-center ${context.tab === 'home' && !context.inMenu ? HOME_ROW_CLASSES[node.id] ?? '' : ''}`}>{node.children.map((child) => renderLayoutNode(child, context, props))}</Inline>;
    case 'stack':
      return <Stack key={node.id} gap="none" className="min-w-0 items-center justify-center">{node.children.map((child) => renderLayoutNode(child, context, props))}</Stack>;
    case 'command':
      return <React.Fragment key={node.id}>{renderCommand(node.commandId, commandOptions(node, context))}</React.Fragment>;
    case 'surface': {
      const surface = RIBBON_TAB_SURFACES.find((candidate) => candidate.id === node.surfaceId);
      return surface && renderSurface ? <React.Fragment key={node.id}>{renderSurface(surface, { inMenu: context.inMenu, mode: context.inMenu ? 'menu' : props.layout.mode })}</React.Fragment> : null;
    }
    case 'split':
      return (
        <Inline key={node.id} gap="none" className={context.inMenu ? 'w-full flex-nowrap' : 'flex-nowrap'}>
          {renderCommand(node.primary, { iconOverride: iconFor({ icon: node.primaryIcon }), ribbonLayoutNodeId: node.id, className: context.inMenu ? 'min-w-0 flex-1 justify-start rounded-none' : context.tab === 'data' ? '!h-6 !min-h-0 gap-1 px-1 text-[11px]' : '!h-6 !min-h-0' })}
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
          trigger={renderCommand(node.trigger, { iconOverride: iconFor({ icon: node.triggerIcon }), ribbonLayoutNodeId: node.id, className: context.inMenu ? 'w-full justify-start rounded-none' : context.tab === 'data' ? '!h-6 !min-h-0 gap-1 px-1 text-[11px]' : '!h-6 !min-h-0' })}
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
  const collapsedGroups = collapsedRibbonGroups(spec.groups.map(group => ({ id: group.id, width: ribbonGroupWidth(group.id) })), layout.width, tab === 'insert' ? ['tables', 'charts', 'illustrations'] : []);
  const groups = spec.groups.map((group, index) => {
    const groupLabel = translateRibbonText(locale, `groups.${group.id}`);
    const collapsed = isHome ? collapseHomeGroup(group.id, layout.width) : collapsedGroups.has(group.id);
    const content = group.children.map((node) => renderLayoutNode(node, { inMenu: collapsed, tab }, props));
    return (
      <React.Fragment key={group.id}>
        {index > 0 ? <Divider orientation="vertical" className="my-3 h-[72px] border-slate-200" /> : null}
        {collapsed ? (
          <Stack data-ribbon-group={group.id} gap="none" className="h-[104px] w-[68px] shrink-0 justify-center px-1">
            <DropdownMenu align="left" trigger={<Button aria-label={`${groupLabel}工具`} title={`${groupLabel}工具`} iconNode={<Icon name={RIBBON_GROUP_ICONS[group.id] ?? 'grid'} size="lg" />} size="sm" variant="ghost" className="h-[76px] !w-[60px] !min-w-0 flex-col gap-1.5 rounded-md border border-slate-200 bg-slate-50/80 !px-1 text-[10px] text-slate-700 hover:border-emerald-300 hover:bg-emerald-50">{groupLabel}<Icon name="chevron-down" size="xs" /></Button>}>
              <Stack gap="sm" className="max-h-[60vh] min-w-[14rem] overflow-y-auto p-3" data-ribbon-overflow={group.id}>
                <Text size="xs" tone="muted">{groupLabel}</Text>
                {content}
              </Stack>
            </DropdownMenu>
          </Stack>
        ) : (
          <Stack data-ribbon-group={group.id} gap="none" style={isHome ? undefined : { width: ribbonGroupWidth(group.id) }} className={`h-[104px] px-1.5 relative min-w-0 shrink-0 justify-between ${ribbonGroupWidthClass(group.id, layout.mode, layout.width, tab)}`}>
            <Inline gap="none" className={`h-[80px] min-h-0 flex-nowrap items-center justify-center`}>{content}</Inline>
            <Text size="xs" tone="subtle" className={`${RIBBON_DENSITY_CLASSES.groupCaption} shrink-0 truncate text-center text-[10px] font-normal text-slate-500 select-none`}>{groupLabel}</Text>
          </Stack>
        )}
      </React.Fragment>
    );
  });
  return (
    <Inline aria-label={`${tab} ribbon commands`} gap="none" tabIndex={0} className={`h-[112px] w-full min-w-0 flex-nowrap items-start overflow-x-auto overflow-y-hidden [scrollbar-width:thin]`} data-testid={tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`} data-ribbon-layout={tab} data-ribbon-breakpoint={layout.mode}>
      <Inline gap="none" className="h-full min-w-max flex-1 bg-white py-1">{groups}</Inline>
    </Inline>
  );
}
