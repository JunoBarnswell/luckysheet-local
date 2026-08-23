import { Box, Button, ContextMenu, Inline, Stack, Text, type ContextMenuItem } from '@react-sheets/ui-system';
import { useMemo, useState } from 'react';
import type { PivotResultNode, PivotResultTree, PivotSourceRowPath } from '@react-sheets/core-model';

export interface PivotResultViewProps {
  tree: PivotResultTree;
  disabled?: boolean;
  expandedFieldIds?: readonly string[];
  onExpandedChange?: (fieldId: string, expanded: boolean) => void;
  onShowDetails?: (paths: PivotSourceRowPath[]) => void;
}

function NodeRows({ disabled, expandedFieldIds, node, onExpandedChange, onShowDetails }: { disabled: boolean; expandedFieldIds?: readonly string[]; node: PivotResultNode; onExpandedChange?: (fieldId: string, expanded: boolean) => void; onShowDetails?: (paths: PivotSourceRowPath[]) => void }) {
  const [menu, setMenu] = useState({ x: 0, y: 0, open: false });
  const hasChildren = node.children.length > 0;
  const expanded = hasChildren && (expandedFieldIds === undefined || expandedFieldIds.includes(node.field ?? ''));
  const menuItems = useMemo<ContextMenuItem[]>(() => [
    { id: 'expand', label: 'Expand', disabled: disabled || !hasChildren || expanded, onSelect: () => node.field && onExpandedChange?.(node.field, true) },
    { id: 'collapse', label: 'Collapse', disabled: disabled || !hasChildren || !expanded, onSelect: () => node.field && onExpandedChange?.(node.field, false) },
    { id: 'expand-field', label: 'Expand Entire Field', disabled: disabled || !hasChildren || expanded, onSelect: () => node.field && onExpandedChange?.(node.field, true) },
    { id: 'collapse-field', label: 'Collapse Entire Field', disabled: disabled || !hasChildren || !expanded, onSelect: () => node.field && onExpandedChange?.(node.field, false) },
    { id: 'details', label: 'Show Details', disabled: disabled || !onShowDetails, onSelect: () => onShowDetails?.(node.sourceRowPaths) },
  ], [disabled, expanded, hasChildren, node.field, node.sourceRowPaths, onExpandedChange, onShowDetails]);
  return (
    <Stack gap="xs" className="border-l border-slate-200 pl-2">
      <Inline gap="xs" className="min-h-8 items-center rounded-md border border-slate-200 bg-white px-2 py-1" onContextMenu={(event) => { event.preventDefault(); setMenu({ x: event.clientX, y: event.clientY, open: true }); }}>
        {hasChildren ? <Button disabled={disabled} icon={expanded ? 'chevron-down' : 'chevron-right'} iconOnly size="xs" variant="ghost" aria-label={`${expanded ? 'Collapse' : 'Expand'} ${node.label}`} onClick={() => node.field && onExpandedChange?.(node.field, !expanded)} /> : <Box className="w-6" />}
        <Text size="xs" weight={node.kind === 'subtotal' ? 'semibold' : 'medium'} className="min-w-28">{node.label}</Text>
        {onShowDetails ? <Button disabled={disabled} size="xs" variant="ghost" onClick={() => onShowDetails(node.sourceRowPaths)}>Details</Button> : null}
        {node.values.flatMap((cell) => cell.values).map((value, index) => <Text key={`${node.label}-${index}`} size="xs" tone={node.kind === 'subtotal' ? 'default' : 'muted'} className="min-w-20 text-right">{value == null ? '—' : String(value)}</Text>)}
      </Inline>
      <ContextMenu x={menu.x} y={menu.y} open={menu.open} items={menuItems} onClose={() => setMenu((current) => ({ ...current, open: false }))} />
      {expanded ? node.children.map((child) => <NodeRows key={`${child.field ?? 'field'}-${String(child.key)}-${child.depth}`} disabled={disabled} expandedFieldIds={expandedFieldIds} node={child} onExpandedChange={onExpandedChange} onShowDetails={onShowDetails} />) : null}
    </Stack>
  );
}

export function PivotResultView({ disabled = false, expandedFieldIds, onExpandedChange, onShowDetails, tree }: PivotResultViewProps) {
  return (
    <Box as="section" aria-label="Pivot table result" className="border-t border-line/80 pt-3">
      <Text size="xs" weight="semibold" tone="muted" className="mb-2 block">RESULT</Text>
      <Stack gap="xs" className="overflow-x-auto">
        <Inline gap="xs" className="min-h-8 items-center rounded-md bg-slate-100 px-2 py-1">
          <Text size="xs" weight="semibold" className="min-w-28">Rows</Text>
          {tree.columnPaths.flatMap((path, pathIndex) => path.length > 0 ? <Text key={`path-${pathIndex}`} size="xs" weight="semibold" className="min-w-20 text-right">{path.join(' / ')}</Text> : <Text key={`path-${pathIndex}`} size="xs" weight="semibold" className="min-w-20 text-right">Values</Text>)}
        </Inline>
        {tree.rows.map((node) => <NodeRows key={`${node.field ?? 'field'}-${String(node.key)}-${node.depth}`} disabled={disabled} expandedFieldIds={expandedFieldIds} node={node} onExpandedChange={onExpandedChange} onShowDetails={onShowDetails} />)}
        {tree.grandTotal ? <Inline gap="xs" className="min-h-8 items-center rounded-md bg-blue-50 px-2 py-1"><Text size="xs" weight="bold" className="min-w-28">Grand Total</Text>{tree.grandTotal.values.map((value, index) => <Text key={`grand-${index}`} size="xs" weight="bold" className="min-w-20 text-right">{value == null ? '—' : String(value)}</Text>)}</Inline> : null}
      </Stack>
    </Box>
  );
}
