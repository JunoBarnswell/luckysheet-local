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
  renderSurface?: (surface: RibbonSurfaceDefinition, context: { inMenu: boolean; mode: RibbonLayoutState['mode'] | 'menu' }) => ReactNode;
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
  tables: 'w-[265px]',
  illustrations: 'w-[308px]',
  controls: 'w-[68px]',
  charts: 'w-[369px]',
  sparklines: 'w-[181px]',
  filters: 'w-[132px]',
  links: 'w-[68px]',
  insertComments: 'w-[83px]',
  text: 'w-[260px]',
  symbols: 'w-[171px]',
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
  clipboard: 'w-[140px]', font: 'w-[244px]', alignment: 'w-[224px]',
  number: 'w-[140px]', styles: 'w-[188px]', cells: 'w-[164px]', editing: 'w-[210px]',
};

function collapseHomeGroup(groupId: RibbonGroupId, width: number): boolean {
  return (width < 1440 && (groupId === 'styles' || groupId === 'cells'))
    || (width < 1100 && groupId === 'editing')
    || (width < 920 && groupId === 'alignment')
    || (width < 760 && groupId === 'number');
}

export function ribbonGroupWidthClass(groupId: RibbonGroupId, mode: RibbonLayoutState['mode'] = 'wide', width = 0, tab?: RibbonLayoutSpec['tab']): string {
  if (tab === 'home') {
    const widths = HOME_RIBBON_GROUP_WIDTH_CLASSES;
    return widths[groupId] ?? 'min-w-[72px] flex-1';
  }
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
  const groups = spec.groups.map((group, index) => {
    const groupLabel = translateRibbonText(locale, `groups.${group.id}`);
    const collapsed = isHome && collapseHomeGroup(group.id, layout.width);
    const content = group.children.map((node) => renderLayoutNode(node, { inMenu: collapsed, tab }, props));
    return (
      <React.Fragment key={group.id}>
        {index > 0 ? <Divider orientation="vertical" className={isHome ? 'my-3 h-[72px] border-slate-200' : RIBBON_DENSITY_CLASSES.groupContent} /> : null}
        {collapsed ? (
          <Stack data-ribbon-group={group.id} gap="none" className="h-[104px] w-[68px] shrink-0 justify-center px-1">
            <DropdownMenu align="left" trigger={<Button aria-label={`${groupLabel}工具`} title={`${groupLabel}工具`} icon="chevron-down" size="sm" variant="ghost" className="h-[76px] w-full flex-col gap-2 text-xs">{groupLabel}</Button>}>
              <Stack gap="sm" className="min-w-[14rem] p-3" data-ribbon-overflow={group.id}>
                <Text size="xs" tone="muted">{groupLabel}</Text>
                {content}
              </Stack>
            </DropdownMenu>
          </Stack>
        ) : (
          <Stack data-ribbon-group={group.id} gap="none" className={`${isHome ? 'h-[104px] px-1.5' : RIBBON_DENSITY_CLASSES.groupContent + ' px-1'} relative min-w-0 shrink-0 justify-between ${ribbonGroupWidthClass(group.id, layout.mode, layout.width, tab)}`}>
            <Inline gap="none" className={`${isHome ? 'h-[80px]' : RIBBON_DENSITY_CLASSES.groupControls} min-h-0 flex-nowrap items-center justify-center`}>{content}</Inline>
            <Text size="xs" tone="subtle" className={`${RIBBON_DENSITY_CLASSES.groupCaption} shrink-0 truncate text-center text-[10px] font-normal text-slate-500 select-none`}>{groupLabel}</Text>
          </Stack>
        )}
      </React.Fragment>
    );
  });
  return (
    <Inline aria-label={`${tab} ribbon commands`} gap="none" tabIndex={0} className={`${isHome ? 'h-[112px]' : RIBBON_DENSITY_CLASSES.commandArea} w-full min-w-0 flex-nowrap items-start overflow-x-auto overflow-y-hidden [scrollbar-width:thin]`} data-testid={tab === 'home' ? 'home-ribbon-groups' : tab === 'insert' ? 'insert-ribbon-groups' : `ribbon-layout-${tab}`} data-ribbon-layout={tab} data-ribbon-breakpoint={layout.mode}>
      {isHome
        ? <Inline gap="none" className="h-full min-w-max flex-1 bg-[var(--home-ribbon-color-surface)] py-1 font-[var(--home-ribbon-font-family)]">{groups}</Inline>
        : tab === 'insert' ? <Inline gap="none" className="h-full min-w-[1905px] flex-1 bg-[#fffdf9]">{groups}</Inline> : groups}
    </Inline>
  );
}
