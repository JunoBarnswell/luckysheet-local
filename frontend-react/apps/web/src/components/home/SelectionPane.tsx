import React, { useMemo, useRef, useState } from 'react';
import {
  Box,
  Button,
  CheckToggle,
  Inline,
  Panel,
  PanelBody,
  PanelDescription,
  PanelHeader,
  PanelTitle,
  ScrollArea,
  Stack,
  StatePanel,
  Text,
  TextInput,
  type IconName,
} from '@react-sheets/ui-system';
import type { Locale } from '../../i18n';
import { homeText, resolveHomeLocale } from './home-localization';

export type SelectionPaneItemKind = 'chart' | 'camera' | 'form-control' | 'image' | 'shape' | 'textbox' | 'connector' | 'pivot-control' | 'other';
export type DrawingSelectionMode = 'replace' | 'toggle' | 'extend';

/**
 * Read-only drawing projection consumed by the pane. The host remains the
 * canonical owner of `sheet.drawings`, selection, visibility, names and
 * z-order mutations; this component only emits user intent.
 */
export interface SelectionPaneItem {
  id: string;
  kind: SelectionPaneItemKind;
  name?: string;
  visible: boolean;
  zIndex: number;
}

export interface SelectionPaneProps {
  items: readonly SelectionPaneItem[];
  selectedIds: readonly string[];
  locale?: Locale;
  disabled?: boolean;
  onSelect: (drawingId: string, mode: DrawingSelectionMode) => void;
  onVisibilityChange: (drawingId: string, visible: boolean) => void;
  onRename?: (drawingId: string, name: string) => void;
  onReorder?: (drawingId: string, direction: 'forward' | 'backward') => void;
}

const kindIcons: Record<SelectionPaneItemKind, IconName> = {
  camera: 'camera',
  'form-control': 'form-control',
  chart: 'chart',
  image: 'file-text',
  shape: 'shape-square',
  textbox: 'type',
  connector: 'shape-square',
  'pivot-control': 'table-pivot',
  other: 'shape-circle',
};

function defaultItemName(item: SelectionPaneItem, ordinal: number, locale: Locale): string {
  return item.name?.trim() || `${homeText(locale, 'object')} ${ordinal + 1}`;
}

export function SelectionPane({
  items,
  selectedIds,
  locale,
  disabled = false,
  onSelect,
  onVisibilityChange,
  onRename,
  onReorder,
}: SelectionPaneProps): React.ReactElement {
  const activeLocale = resolveHomeLocale(locale);
  const [editingId, setEditingId] = useState<string>();
  const [draftName, setDraftName] = useState('');
  const renameFinalizedRef = useRef(false);
  const orderedItems = useMemo(
    () => [...items].sort((left, right) => right.zIndex - left.zIndex || left.id.localeCompare(right.id)),
    [items],
  );

  const beginRename = (item: SelectionPaneItem, ordinal: number) => {
    if (!onRename || disabled) return;
    renameFinalizedRef.current = false;
    setEditingId(item.id);
    setDraftName(defaultItemName(item, ordinal, activeLocale));
  };
  const finishRename = (commit: boolean) => {
    if (renameFinalizedRef.current) return;
    renameFinalizedRef.current = true;
    const drawingId = editingId;
    const name = draftName.trim();
    setEditingId(undefined);
    setDraftName('');
    if (commit && drawingId && name && onRename) onRename(drawingId, name);
  };

  return (
    <Panel className="flex min-h-0 flex-1 flex-col overflow-hidden" data-testid="selection-pane">
      <PanelHeader>
        <Box>
          <PanelTitle>{homeText(activeLocale, 'selectionPane')}</PanelTitle>
          <PanelDescription>{homeText(activeLocale, 'selectionPaneDescription')}</PanelDescription>
        </Box>
      </PanelHeader>
      <PanelBody className="min-h-0 flex-1 p-0">
        {orderedItems.length === 0 ? (
          <StatePanel
            kind="empty"
            icon="shape-square"
            title={homeText(activeLocale, 'selectionPane')}
            description={homeText(activeLocale, 'noObjects')}
            className="m-4"
          />
        ) : (
          <ScrollArea className="max-h-[calc(100vh-16rem)] p-2">
            <Stack gap="xs">
              {orderedItems.map((item, ordinal) => {
                const selected = selectedIds.includes(item.id);
                const editing = editingId === item.id;
                const name = defaultItemName(item, ordinal, activeLocale);
                return (
                  <Box
                    key={item.id}
                    className={selected
                      ? 'rounded-lg border border-blue-200 bg-blue-50 p-2 shadow-sm'
                      : 'rounded-lg border border-transparent p-2 hover:border-slate-200 hover:bg-slate-50'}
                  >
                    <Inline gap="sm" className="min-w-0">
                      <Button
                        size="sm"
                        variant="ghost"
                        icon={kindIcons[item.kind]}
                        iconOnly
                        disabled={disabled}
                        aria-label={`${homeText(activeLocale, 'select')} ${name}`}
                        aria-pressed={selected}
                        onClick={(event) => onSelect(item.id, event.shiftKey ? 'extend' : event.metaKey || event.ctrlKey ? 'toggle' : 'replace')}
                      />
                      {editing ? (
                        <TextInput
                          aria-label={homeText(activeLocale, 'rename')}
                          size={1}
                          autoFocus
                          value={draftName}
                          onChange={(event) => setDraftName(event.target.value)}
                          onBlur={() => finishRename(true)}
                          onKeyDown={(event) => {
                            if (event.key === 'Enter') finishRename(true);
                            if (event.key === 'Escape') finishRename(false);
                          }}
                        />
                      ) : (
                        <Button
                          size="sm"
                          variant="ghost"
                          disabled={disabled}
                          className="min-w-0 flex-1 justify-start truncate px-1.5"
                          aria-pressed={selected}
                          onClick={(event) => onSelect(item.id, event.shiftKey ? 'extend' : event.metaKey || event.ctrlKey ? 'toggle' : 'replace')}
                          onDoubleClick={() => beginRename(item, ordinal)}
                        >
                          {name}
                        </Button>
                      )}
                      <CheckToggle
                        checked={item.visible}
                        disabled={disabled}
                        aria-label={`${name}: ${homeText(activeLocale, 'visible')}`}
                        onChange={(event) => onVisibilityChange(item.id, event.target.checked)}
                      />
                    </Inline>
                    <Inline gap="xs" className="mt-1 justify-end">
                      {onRename ? (
                        <Button
                          size="xs"
                          variant="ghost"
                          disabled={disabled || editing}
                          onClick={() => beginRename(item, ordinal)}
                        >
                          {homeText(activeLocale, 'rename')}
                        </Button>
                      ) : null}
                      {onReorder ? (
                        <>
                          <Button
                            size="xs"
                            variant="ghost"
                            icon="arrow-up"
                            iconOnly
                            disabled={disabled}
                            aria-label={homeText(activeLocale, 'bringForward')}
                            title={homeText(activeLocale, 'bringForward')}
                            onClick={() => onReorder(item.id, 'forward')}
                          />
                          <Button
                            size="xs"
                            variant="ghost"
                            icon="arrow-down"
                            iconOnly
                            disabled={disabled}
                            aria-label={homeText(activeLocale, 'sendBackward')}
                            title={homeText(activeLocale, 'sendBackward')}
                            onClick={() => onReorder(item.id, 'backward')}
                          />
                        </>
                      ) : null}
                    </Inline>
                  </Box>
                );
              })}
            </Stack>
          </ScrollArea>
        )}
      </PanelBody>
    </Panel>
  );
}
