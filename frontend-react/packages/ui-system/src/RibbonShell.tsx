import { useLayoutEffect, useRef, useState, type ReactNode, type RefObject } from 'react';
import { Button } from './Button';
import { DropdownMenu } from './DropdownMenu';
import { Box, Inline, Text } from './layout';
import { RIBBON_DENSITY, RIBBON_DENSITY_CLASSES, RIBBON_TAB_ORDER, type RibbonKeyTipBinding, type RibbonKeyTipState, type RibbonLayoutMode, type RibbonLayoutState, type RibbonTabId } from './shell-types';
import { Tab, TabList, Tabs } from './Tabs';

export type { RibbonKeyTipBinding, RibbonKeyTipState, RibbonLayoutMode, RibbonLayoutState, RibbonTabId } from './shell-types';
export { RIBBON_TAB_ORDER } from './shell-types';

/**
 * Ribbon geometry is a fixed product surface. Viewport width only controls
 * horizontal scrolling; it must never change command density, labels or
 * available capabilities.
 */
export function ribbonLayoutModeForWidth(_width: number): RibbonLayoutMode {
  return 'wide';
}

const DENSE_COMPACT_MIN_WIDTH = 1440;
const RIBBON_TAB_WIDTHS: Partial<Record<RibbonTabId, string>> = {
  file: 'w-[52px]',
  home: 'w-[56px]',
  insert: 'w-[64px]',
  pageLayout: 'w-[82px]',
  formulas: 'w-[60px]',
  data: 'w-[52px]',
  view: 'w-[58px]',
  review: 'w-[58px]',
  settings: 'w-[58px]',
};

function isDenseCompact(width: number, mode: RibbonLayoutMode): boolean {
  return mode === 'compact' && width >= DENSE_COMPACT_MIN_WIDTH;
}

export interface RibbonShellProps {
  activeTab: RibbonTabId;
  children: ReactNode | ((layout: RibbonLayoutState) => ReactNode);
  /** Context tabs are session state supplied by the host, never workbook data. */
  contextualTabs?: readonly RibbonTabId[];
  disabled?: boolean;
  onFileEntry?: () => void;
  onTabChange: (tab: RibbonTabId) => void;
  status?: ReactNode;
  tabLabel: (tab: RibbonTabId) => string;
  keyTipState?: RibbonKeyTipState;
  keyTipBindings?: readonly RibbonKeyTipBinding[];
}

function RibbonKeyTipHints({ rootRef, state, bindings }: { rootRef: RefObject<HTMLDivElement | null>; state?: RibbonKeyTipState; bindings: readonly RibbonKeyTipBinding[] }): React.ReactElement | null {
  const [hints, setHints] = useState<Array<{ id: string; label: string; left: number; top: number }>>([]);

  useLayoutEffect(() => {
    if (!state?.active) {
      setHints([]);
      return;
    }
    const root = rootRef.current;
    if (!root) return;
    const rootRect = root.getBoundingClientRect();
    const next = bindings
      .filter((binding) => binding.sequence.startsWith(state.prefix.toLocaleUpperCase()))
      .flatMap((binding) => {
        const selector = binding.target.kind === 'tab'
          ? `[data-ribbon-keytip-tab="${binding.target.id}"]`
          : `[data-ribbon-keytip="${binding.sequence}"]`;
        const node = root.querySelector<HTMLElement>(selector);
        if (!node) return [];
        const rect = node.getBoundingClientRect();
        return [{ id: `${binding.sequence}:${binding.target.kind}:${binding.target.id}`, label: binding.sequence.slice(state.prefix.length) || binding.sequence, left: rect.left - rootRect.left + Math.max(0, rect.width - 18), top: rect.top - rootRect.top + 2 }];
      });
    setHints(next);
  }, [bindings, rootRef, state]);

  if (!state?.active) return null;
  return <>{hints.map((hint) => <Box key={hint.id} aria-hidden="true" className="pointer-events-none absolute z-[80] min-w-4 rounded border border-amber-700 bg-amber-100 px-1 text-center font-mono text-[10px] font-semibold leading-4 text-amber-950 shadow-sm" style={{ left: hint.left, top: hint.top }}>{hint.label}</Box>)}</>;
}

export function RibbonShell({
  activeTab,
  children,
  contextualTabs = [],
  disabled = false,
  onFileEntry,
  onTabChange,
  status,
  tabLabel,
  keyTipState,
  keyTipBindings = [],
}: RibbonShellProps) {
  const rootRef = useRef<HTMLDivElement>(null);
  const [layout, setLayout] = useState<RibbonLayoutState>(() => {
    const width = typeof window === 'undefined' ? 1920 : Math.round(window.innerWidth);
    return { mode: ribbonLayoutModeForWidth(width), width };
  });
  const denseCompact = isDenseCompact(layout.width, layout.mode);
  const tabs = [
    ...RIBBON_TAB_ORDER.filter((tab) => !onFileEntry || tab !== 'file'),
    ...contextualTabs.filter((tab, index, values) => !RIBBON_TAB_ORDER.includes(tab) && values.indexOf(tab) === index),
  ];

  useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root) return undefined;
    const update = (width: number) => {
      const nextWidth = Math.max(0, Math.round(width));
      setLayout((previous) => {
        const mode = ribbonLayoutModeForWidth(nextWidth);
        return previous.width === nextWidth && previous.mode === mode ? previous : { mode, width: nextWidth };
      });
    };
    update(root.getBoundingClientRect().width);
    if (typeof ResizeObserver === 'undefined') return undefined;
    const observer = new ResizeObserver((entries) => {
      const width = entries[0]?.contentRect.width;
      if (width !== undefined) update(width);
    });
    observer.observe(root);
    return () => observer.disconnect();
  }, []);

  return (
    <Tabs ref={rootRef} className={`${RIBBON_DENSITY_CLASSES.shell} relative overflow-hidden border-b border-[#e7e7e7] bg-[#f5f5f3]`} data-ribbon-layout={layout.mode} data-ribbon-height={RIBBON_DENSITY.shellHeight} data-testid="ribbon-shell">
      <Inline gap="none" className={`${RIBBON_DENSITY_CLASSES.tabStrip} min-w-0 flex-nowrap px-2`}>
        {onFileEntry ? (
          <DropdownMenu
            disabled={disabled}
            trigger={(
              <Button
                aria-label="Open workbook menu"
                data-ribbon-keytip-tab="file"
                className={`h-full shrink-0 rounded-none border-b-2 border-transparent px-0 font-semibold text-slate-700 hover:border-[#217345] hover:bg-[#f3f8f4] hover:text-[#217345] ${RIBBON_TAB_WIDTHS.file} ${denseCompact ? 'text-[14px]' : layout.mode === 'wide' ? 'text-xs' : 'text-[11px]'}`}
                size="sm"
                variant="ghost"
              >
                文件
              </Button>
            )}
          >
            {({ close }) => (
              <Button className="min-w-[9rem] justify-start" onClick={() => { close(); onFileEntry(); }} size="sm" variant="ghost">
                File / 工作簿
              </Button>
            )}
          </DropdownMenu>
        ) : null}
        <TabList label="Workbook ribbon tabs" className="h-full min-w-0 flex-1 flex-nowrap gap-0 overflow-x-auto overflow-y-hidden [scrollbar-width:thin]">
          {tabs.map((tab) => (
            <Tab
              key={tab}
              active={activeTab === tab}
              data-testid={`ribbon-tab-${tab}`}
              data-ribbon-keytip-tab={tab}
              disabled={disabled}
              onClick={() => onTabChange(tab)}
              className={`!h-full !min-h-0 min-w-0 !shrink-0 rounded-none border-b-2 border-transparent font-semibold text-[#3d3c41] aria-selected:!bg-white aria-selected:!text-[#217345] aria-selected:border-[#217345] ${RIBBON_TAB_WIDTHS[tab] ?? 'w-[58px]'} !px-0 ${denseCompact ? 'text-[14px]' : layout.mode === 'wide' ? 'text-xs' : 'text-[11px]'}`}
            >
              {tabLabel(tab)}
            </Tab>
          ))}
        </TabList>
        {status ? (
          <Inline gap="xs" className={`${layout.mode === 'wide' ? 'flex' : 'hidden'} ml-auto shrink-0 border-l border-slate-100 pl-3`}>
            {status}
          </Inline>
        ) : null}
      </Inline>
      <Box className={`${RIBBON_DENSITY_CLASSES.commandArea} overflow-hidden border-t-0 bg-white px-0 py-0`}>
        {typeof children === 'function' ? children(layout) : children}
      </Box>
      <RibbonKeyTipHints rootRef={rootRef} state={keyTipState} bindings={keyTipBindings} />
    </Tabs>
  );
}

export function RibbonEmptyState({ message }: { message: string }) {
  return (
    <Inline gap="sm" className="min-h-[88px] items-center">
      <Text size="sm" tone="muted">{message}</Text>
    </Inline>
  );
}
