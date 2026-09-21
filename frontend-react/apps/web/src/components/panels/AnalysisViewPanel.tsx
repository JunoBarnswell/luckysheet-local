import React, { useMemo, useState } from 'react';
import { Box, Button, Inline, Panel, PanelBody, PanelHeader, PanelTitle, Stack, StatePanel, Text, TextInput } from '@react-sheets/ui-system';
import type { AnalysisViewDefinition, AnalysisViewLayout, WorkbookTableModel } from '@react-sheets/core-model';

export interface AnalysisViewPanelProps {
  views: readonly AnalysisViewDefinition[];
  tables: readonly WorkbookTableModel[];
  onSetView: (view: AnalysisViewDefinition) => void;
  onRemoveView: (viewId: string) => void;
  onClose?: () => void;
}

const DEFAULT_LAYOUT: AnalysisViewLayout = { columns: 2, rowHeightPx: 240, gapPx: 12 };

export function AnalysisViewPanel({ views, tables, onSetView, onRemoveView, onClose }: AnalysisViewPanelProps) {
  const [name, setName] = useState('分析视图');
  const source = tables[0];
  const defaultView = useMemo(() => {
    if (!source) return undefined;
    return {
      kind: 'analysis' as const,
      id: `analysis-${Date.now().toString(36)}`,
      name: name.trim() || '分析视图',
      tableId: source.id,
      fields: source.fields.slice(0, 12).map((field) => ({ fieldId: field.id, caption: field.name })),
      filters: [],
      charts: [],
      layout: DEFAULT_LAYOUT,
      revision: 0,
    } satisfies AnalysisViewDefinition;
  }, [name, source]);

  return (
    <Box as="aside" aria-label="Analysis views" className="flex h-full min-h-0 flex-1 flex-col bg-white">
      <PanelHeader>
        <Inline gap="sm" className="items-center justify-between">
          <Inline gap="sm"><Text size="sm" weight="semibold">共享分析视图</Text><Text size="xs" tone="muted">服务端状态</Text></Inline>
          {onClose ? <Button size="sm" variant="ghost" onClick={onClose}>关闭</Button> : null}
        </Inline>
      </PanelHeader>
      <PanelBody className="min-h-0 flex-1 overflow-auto p-3">
        <Stack gap="md">
          <Panel className="border border-slate-200 shadow-none">
            <PanelHeader><PanelTitle as="h3" size="sm">新建分析视图</PanelTitle></PanelHeader>
            <PanelBody>
              <Stack gap="sm">
                <TextInput aria-label="分析视图名称" value={name} onChange={(event) => setName(event.target.value)} />
                <Text size="xs" tone="muted">从第一个工作簿表创建字段映射；保存后通过协作 mutation 共享。</Text>
                <Button size="sm" variant="primary" disabled={!defaultView} onClick={() => { if (defaultView) onSetView(defaultView); }}>创建并共享</Button>
              </Stack>
            </PanelBody>
          </Panel>
          {views.length === 0 ? <StatePanel kind="empty" title="暂无共享分析视图" description="创建后，筛选条件和图表字段映射会随工作簿保存。" /> : views.map((view) => (
            <Panel key={view.id} className="border border-slate-200 shadow-none">
              <PanelHeader>
                <Inline gap="sm" className="items-center justify-between">
                  <Stack gap="none"><Text size="sm" weight="semibold">{view.name}</Text><Text size="xs" tone="muted">{view.tableId} · v{view.revision}</Text></Stack>
                  <Button size="sm" variant="ghost" onClick={() => onRemoveView(view.id)}>删除</Button>
                </Inline>
              </PanelHeader>
              <PanelBody className="py-2">
                <Inline gap="md" className="text-xs text-slate-600"><Text size="xs">字段 {view.fields.length}</Text><Text size="xs">筛选 {view.filters.length}</Text><Text size="xs">图表 {view.charts.length}</Text></Inline>
              </PanelBody>
            </Panel>
          ))}
        </Stack>
      </PanelBody>
    </Box>
  );
}
