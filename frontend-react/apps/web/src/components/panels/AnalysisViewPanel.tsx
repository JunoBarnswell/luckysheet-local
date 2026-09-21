import React, { useEffect, useMemo, useState } from 'react';
import {
  Box,
  Button,
  CheckToggle,
  Inline,
  Panel,
  PanelBody,
  PanelHeader,
  PanelTitle,
  Select,
  Stack,
  StatePanel,
  Text,
  TextInput,
} from '@react-sheets/ui-system';
import type {
  AnalysisChartFieldMap,
  AnalysisFilterClause,
  AnalysisFilterOperator,
  AnalysisViewDefinition,
  AnalysisViewLayout,
  TableScalar,
  WorkbookTableField,
  WorkbookTableModel,
} from '@react-sheets/core-model';

export interface AnalysisViewPanelProps {
  views: readonly AnalysisViewDefinition[];
  tables: readonly WorkbookTableModel[];
  chartIds?: readonly string[];
  onSetView: (view: AnalysisViewDefinition) => void;
  onRemoveView: (viewId: string) => void;
  onClose?: () => void;
}

const DEFAULT_LAYOUT: AnalysisViewLayout = { columns: 2, rowHeightPx: 240, gapPx: 12 };
const FILTER_OPERATORS: readonly AnalysisFilterOperator[] = ['equals', 'not-equals', 'contains', 'in', 'between'];
const FILTER_OPERATOR_LABELS: Record<AnalysisFilterOperator, string> = {
  equals: '等于',
  'not-equals': '不等于',
  contains: '包含',
  in: '属于（逗号分隔）',
  between: '介于（两个值）',
};
const CHART_FIELD_ROLES: readonly (keyof AnalysisChartFieldMap)[] = ['category', 'series', 'value', 'color', 'size', 'tooltip'];
const CHART_FIELD_ROLE_LABELS: Record<keyof AnalysisChartFieldMap, string> = {
  category: '类别',
  series: '系列',
  value: '值',
  color: '颜色',
  size: '大小',
  tooltip: '提示',
};

type DraftUpdater = (draft: AnalysisViewDefinition) => AnalysisViewDefinition;

function newId(prefix: string): string {
  return prefix + '-' + globalThis.crypto.randomUUID();
}

function defaultFilterValue(field: WorkbookTableField): TableScalar {
  if (field.type === 'number') return 0;
  if (field.type === 'boolean') return true;
  return '';
}

function formatScalar(value: TableScalar): string {
  return value === null ? 'null' : String(value);
}

function parseScalar(field: WorkbookTableField, value: string): TableScalar | undefined {
  if (field.type === 'number') {
    if (value.trim() === '') return undefined;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : undefined;
  }
  if (field.type === 'boolean') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true') return true;
    if (normalized === 'false') return false;
    return undefined;
  }
  if (field.type === 'date' && value.trim() === '') return undefined;
  return value;
}

function parseFilterValues(field: WorkbookTableField, operator: AnalysisFilterOperator, input: string): TableScalar[] | undefined {
  const parts = operator === 'in' || operator === 'between'
    ? input.split(',').map((part) => part.trim()).filter((part) => part.length > 0)
    : [input];
  if (operator === 'between' && parts.length !== 2) return undefined;
  if (operator === 'in' && parts.length === 0) return undefined;
  if (operator !== 'between' && operator !== 'in' && parts.length !== 1) return undefined;
  const parsed = parts.map((part) => parseScalar(field, part));
  return parsed.every((value): value is TableScalar => value !== undefined) ? parsed : undefined;
}

function filterText(filter: AnalysisFilterClause): string {
  return filter.values.map(formatScalar).join(', ');
}

function tableForView(view: AnalysisViewDefinition | undefined, tables: readonly WorkbookTableModel[]): WorkbookTableModel | undefined {
  return view ? tables.find((table) => table.id === view.tableId) : undefined;
}

function validateDraft(view: AnalysisViewDefinition, table: WorkbookTableModel | undefined, chartIds: readonly string[]): string[] {
  const errors: string[] = [];
  if (view.name.trim().length === 0) errors.push('分析视图名称不能为空。');
  if (!table) errors.push('分析视图引用的表不存在，请重新选择表。');
  if (table) {
    const fieldIds = new Set(table.fields.map((field) => field.id));
    const selectedIds = new Set<string>();
    for (const field of view.fields) {
      if (!fieldIds.has(field.fieldId)) errors.push('字段“' + field.fieldId + '”已不存在。');
      if (!selectedIds.add(field.fieldId)) errors.push('字段“' + field.fieldId + '”重复选择。');
      if (field.caption.trim().length === 0) errors.push('字段“' + field.fieldId + '”的显示名称不能为空。');
    }
    for (const filter of view.filters) {
      if (!fieldIds.has(filter.fieldId)) errors.push('筛选字段“' + filter.fieldId + '”已不存在。');
      else if (!selectedIds.has(filter.fieldId)) errors.push('筛选字段“' + filter.fieldId + '”未包含在当前字段选择中，请先重新选择该字段。');
      if (filter.values.length === 0) errors.push('筛选“' + filter.id + '”至少需要一个值。');
    }
    for (const chart of view.charts) {
      if (chartIds.length > 0 && !chartIds.includes(chart.chartId)) errors.push('图表“' + chart.chartId + '”已不存在，请移除该绑定。');
      for (const fieldId of Object.values(chart.fieldMap)) {
        if (fieldId !== undefined && !fieldIds.has(fieldId)) errors.push('图表“' + chart.chartId + '”引用了不存在的字段。');
        else if (fieldId !== undefined && !selectedIds.has(fieldId)) errors.push('图表“' + chart.chartId + '”引用了未选择的字段，请先重新选择该字段。');
      }
    }
  }
  if (!Number.isSafeInteger(view.layout.columns) || view.layout.columns < 1 || view.layout.columns > 12) errors.push('布局列数必须是 1 到 12。');
  if (!Number.isFinite(view.layout.rowHeightPx) || view.layout.rowHeightPx < 1) errors.push('图表行高必须大于 0。');
  if (!Number.isFinite(view.layout.gapPx) || view.layout.gapPx < 0) errors.push('图表间距不能为负数。');
  return errors;
}

function FieldLabel({ children }: { children: React.ReactNode }) {
  return <Text size="xs" weight="medium" className="mb-1 text-slate-700">{children}</Text>;
}

export function AnalysisViewPanel({ views, tables, chartIds = [], onSetView, onRemoveView, onClose }: AnalysisViewPanelProps) {
  const [name, setName] = useState('分析视图');
  const [activeViewId, setActiveViewId] = useState<string>();
  const [draft, setDraft] = useState<AnalysisViewDefinition>();
  const [filterInputText, setFilterInputText] = useState<Record<string, string>>({});
  const [filterErrors, setFilterErrors] = useState<Record<string, string>>({});
  const [validationErrors, setValidationErrors] = useState<string[]>([]);

  const activeView = useMemo(() => views.find((view) => view.id === activeViewId), [activeViewId, views]);
  const activeTable = tableForView(draft ?? activeView, tables);
  const source = tables[0];

  useEffect(() => {
    if (views.length === 0) {
      setActiveViewId(undefined);
      setDraft(undefined);
      return;
    }
    setActiveViewId((current) => current && views.some((view) => view.id === current) ? current : views[0]!.id);
  }, [views]);

  useEffect(() => {
    if (!activeView) {
      setDraft(undefined);
      return;
    }
    setDraft((current) => current && current.id === activeView.id && current.revision === activeView.revision ? current : structuredClone(activeView));
    setFilterInputText({});
    setFilterErrors({});
    setValidationErrors([]);
  }, [activeView?.id, activeView?.revision]);

  const updateDraft = (updater: DraftUpdater) => {
    setDraft((current) => current ? updater(structuredClone(current)) : current);
    setValidationErrors([]);
  };

  const createView = () => {
    if (!source) return;
    const view: AnalysisViewDefinition = {
      kind: 'analysis',
      id: newId('analysis'),
      name: name.trim() || '分析视图',
      tableId: source.id,
      fields: source.fields.slice(0, 12).map((field) => ({ fieldId: field.id, caption: field.name })),
      filters: [],
      charts: [],
      layout: structuredClone(DEFAULT_LAYOUT),
      revision: 0,
    };
    onSetView(view);
    setActiveViewId(view.id);
    setDraft(view);
    setName('分析视图');
  };

  const selectView = (view: AnalysisViewDefinition) => {
    setActiveViewId(view.id);
    setDraft(structuredClone(view));
    setFilterInputText({});
    setFilterErrors({});
    setValidationErrors([]);
  };

  const applyDraft = () => {
    if (!draft) return;
    const errors = [...validateDraft(draft, activeTable, chartIds), ...Object.values(filterErrors)];
    if (errors.length > 0) {
      setValidationErrors([...new Set(errors)]);
      return;
    }
    const next = structuredClone(draft);
    next.name = next.name.trim();
    next.revision += 1;
    onSetView(next);
    setDraft(next);
    setValidationErrors([]);
  };

  const toggleField = (field: WorkbookTableField, checked: boolean) => {
    updateDraft((current) => ({
      ...current,
      fields: checked
        ? [...current.fields, { fieldId: field.id, caption: field.name }]
        : current.fields.filter((entry) => entry.fieldId !== field.id),
    }));
  };

  const moveField = (fieldId: string, direction: -1 | 1) => {
    updateDraft((current) => {
      const index = current.fields.findIndex((field) => field.fieldId === fieldId);
      const target = index + direction;
      if (index < 0 || target < 0 || target >= current.fields.length) return current;
      const fields = [...current.fields];
      [fields[index], fields[target]] = [fields[target]!, fields[index]!];
      return { ...current, fields };
    });
  };

  const addFilter = () => {
    if (!activeTable || activeTable.fields.length === 0) return;
    const field = activeTable.fields[0]!;
    const filter: AnalysisFilterClause = { id: newId('filter'), fieldId: field.id, operator: 'equals', values: [defaultFilterValue(field)] };
    updateDraft((current) => ({ ...current, filters: [...current.filters, filter] }));
  };

  const removeFilter = (filterId: string) => {
    updateDraft((current) => ({ ...current, filters: current.filters.filter((filter) => filter.id !== filterId) }));
    setFilterInputText((current) => { const next = { ...current }; delete next[filterId]; return next; });
    setFilterErrors((current) => { const next = { ...current }; delete next[filterId]; return next; });
  };

  const updateFilterField = (filterId: string, fieldId: string) => {
    const field = activeTable?.fields.find((candidate) => candidate.id === fieldId);
    if (!field) return;
    updateDraft((current) => ({
      ...current,
      filters: current.filters.map((filter) => filter.id === filterId ? { ...filter, fieldId, values: [defaultFilterValue(field)] } : filter),
    }));
    setFilterInputText((current) => { const next = { ...current }; delete next[filterId]; return next; });
    setFilterErrors((current) => { const next = { ...current }; delete next[filterId]; return next; });
  };

  const updateFilterOperator = (filterId: string, operator: AnalysisFilterOperator) => {
    updateDraft((current) => ({
      ...current,
      filters: current.filters.map((filter) => {
        if (filter.id !== filterId) return filter;
        const values = operator === 'between'
          ? [filter.values[0] ?? '', filter.values[1] ?? filter.values[0] ?? '']
          : [filter.values[0] ?? ''];
        return { ...filter, operator, values };
      }),
    }));
    setFilterInputText((current) => { const next = { ...current }; delete next[filterId]; return next; });
    setFilterErrors((current) => { const next = { ...current }; delete next[filterId]; return next; });
  };

  const updateFilterValue = (filter: AnalysisFilterClause, value: string) => {
    const field = activeTable?.fields.find((candidate) => candidate.id === filter.fieldId);
    if (!field) return;
    setFilterInputText((current) => ({ ...current, [filter.id]: value }));
    const parsed = parseFilterValues(field, filter.operator, value);
    if (!parsed) {
      setFilterErrors((current) => ({ ...current, [filter.id]: field.name + ' 的筛选值格式不正确。' }));
      return;
    }
    setFilterErrors((current) => { const next = { ...current }; delete next[filter.id]; return next; });
    updateDraft((current) => ({ ...current, filters: current.filters.map((entry) => entry.id === filter.id ? { ...entry, values: parsed } : entry) }));
  };

  const addChartBinding = () => {
    if (!draft || chartIds.length === 0) return;
    const chartId = chartIds.find((candidate) => !draft.charts.some((binding) => binding.chartId === candidate));
    if (!chartId) return;
    const first = activeTable?.fields[0]?.id;
    const second = activeTable?.fields[1]?.id ?? first;
    const fieldMap: AnalysisChartFieldMap = first ? { category: first, value: second } : {};
    updateDraft((current) => ({ ...current, charts: [...current.charts, { chartId, fieldMap }] }));
  };

  const removeChartBinding = (chartId: string) => updateDraft((current) => ({ ...current, charts: current.charts.filter((binding) => binding.chartId !== chartId) }));

  const updateChartField = (chartId: string, role: keyof AnalysisChartFieldMap, fieldId: string) => {
    updateDraft((current) => ({
      ...current,
      charts: current.charts.map((binding) => {
        if (binding.chartId !== chartId) return binding;
        const fieldMap: AnalysisChartFieldMap = { ...binding.fieldMap };
        if (fieldId) fieldMap[role] = fieldId;
        else delete fieldMap[role];
        return { ...binding, fieldMap };
      }),
    }));
  };

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
                <Button size="sm" variant="primary" disabled={!source} onClick={createView}>创建并共享</Button>
              </Stack>
            </PanelBody>
          </Panel>

          {views.length === 0 ? <StatePanel kind="empty" title="暂无共享分析视图" description="创建后可在下方编辑字段、筛选、图表绑定和布局。" /> : (
            <Stack gap="sm">
              {views.map((view) => (
                <Panel key={view.id} className={view.id === activeViewId ? 'border border-blue-300 shadow-sm' : 'border border-slate-200 shadow-none'}>
                  <PanelHeader>
                    <Inline gap="sm" className="items-center justify-between">
                      <Button size="sm" variant="ghost" className="min-w-0 flex-1 justify-start px-0 text-left" onClick={() => selectView(view)}>
                        <Stack gap="none"><Text size="sm" weight="semibold" className="truncate">{view.name}</Text><Text size="xs" tone="muted">{view.tableId} · v{view.revision}</Text></Stack>
                      </Button>
                      <Button size="sm" variant="ghost" onClick={() => onRemoveView(view.id)}>删除</Button>
                    </Inline>
                  </PanelHeader>
                  <PanelBody className="py-2">
                    <Inline gap="md" className="text-xs text-slate-600"><Text size="xs">字段 {view.fields.length}</Text><Text size="xs">筛选 {view.filters.length}</Text><Text size="xs">图表 {view.charts.length}</Text></Inline>
                  </PanelBody>
                </Panel>
              ))}
            </Stack>
          )}

          {draft ? (
            <Panel className="border border-blue-200 bg-blue-50/20 shadow-none">
              <PanelHeader><Inline gap="sm" className="items-center justify-between"><PanelTitle as="h3" size="sm">编辑分析视图</PanelTitle><Text size="xs" tone="muted">草稿 v{draft.revision}</Text></Inline></PanelHeader>
              <PanelBody>
                <Stack gap="md">
                  <Box><FieldLabel>视图名称</FieldLabel><TextInput aria-label="当前分析视图名称" value={draft.name} onChange={(event) => updateDraft((current) => ({ ...current, name: event.target.value }))} /></Box>

                  <Box>
                    <FieldLabel>字段与顺序</FieldLabel>
                    <Stack gap="xs">
                      {(activeTable?.fields ?? []).map((field) => {
                        const selected = draft.fields.find((entry) => entry.fieldId === field.id);
                        const selectedIndex = selected ? draft.fields.findIndex((entry) => entry.fieldId === field.id) : -1;
                        return <Box key={field.id} className="rounded border border-slate-200 bg-white p-2">
                          <Inline gap="xs" className="items-center">
                            <CheckToggle aria-label={'选择字段 ' + field.name} checked={Boolean(selected)} label={field.name + ' · ' + field.type} onChange={(event) => toggleField(field, event.currentTarget.checked)} />
                            {selected ? <>
                              <Button size="xs" variant="ghost" disabled={selectedIndex <= 0} aria-label={'字段 ' + field.name + ' 上移'} onClick={() => moveField(field.id, -1)}>↑</Button>
                              <Button size="xs" variant="ghost" disabled={selectedIndex < 0 || selectedIndex >= draft.fields.length - 1} aria-label={'字段 ' + field.name + ' 下移'} onClick={() => moveField(field.id, 1)}>↓</Button>
                              <TextInput aria-label={'字段 ' + field.name + ' 显示名称'} className="min-w-0 flex-1" value={selected.caption} onChange={(event) => updateDraft((current) => ({ ...current, fields: current.fields.map((entry) => entry.fieldId === field.id ? { ...entry, caption: event.target.value } : entry) }))} />
                            </> : null}
                          </Inline>
                        </Box>;
                      })}
                    </Stack>
                  </Box>

                  <Box>
                    <Inline gap="sm" className="items-center justify-between"><FieldLabel>筛选条件</FieldLabel><Button size="xs" variant="secondary" disabled={!activeTable || activeTable.fields.length === 0} onClick={addFilter}>添加筛选</Button></Inline>
                    <Stack gap="xs">
                      {draft.filters.map((filter) => {
                        const field = activeTable?.fields.find((candidate) => candidate.id === filter.fieldId);
                        return <Box key={filter.id} className="rounded border border-slate-200 bg-white p-2">
                          <Stack gap="xs">
                            <Inline gap="xs" className="items-center">
                              <Select aria-label="筛选字段" sizeVariant="sm" value={filter.fieldId} onChange={(event) => updateFilterField(filter.id, event.currentTarget.value)} options={(activeTable?.fields ?? []).map((candidate) => ({ value: candidate.id, label: candidate.name }))} />
                              <Select aria-label="筛选运算符" sizeVariant="sm" value={filter.operator} onChange={(event) => updateFilterOperator(filter.id, event.currentTarget.value as AnalysisFilterOperator)} options={FILTER_OPERATORS.map((operator) => ({ value: operator, label: FILTER_OPERATOR_LABELS[operator] }))} />
                              <Button size="xs" variant="ghost" aria-label="删除筛选" onClick={() => removeFilter(filter.id)}>删除</Button>
                            </Inline>
                            <TextInput aria-label="筛选值" value={filterInputText[filter.id] ?? filterText(filter)} placeholder={filter.operator === 'between' ? '例如 10, 100' : filter.operator === 'in' ? '例如 华东, 华南' : field?.type === 'boolean' ? 'true 或 false' : undefined} onChange={(event) => updateFilterValue(filter, event.target.value)} />
                            {filterErrors[filter.id] ? <Text size="xs" tone="danger">{filterErrors[filter.id]}</Text> : null}
                          </Stack>
                        </Box>;
                      })}
                      {draft.filters.length === 0 ? <Text size="xs" tone="muted">未设置筛选；组内条件使用 AND。</Text> : null}
                    </Stack>
                  </Box>

                  <Box>
                    <Inline gap="sm" className="items-center justify-between"><FieldLabel>图表字段映射</FieldLabel><Button size="xs" variant="secondary" disabled={chartIds.length === 0 || !activeTable} onClick={addChartBinding}>绑定图表</Button></Inline>
                    <Stack gap="xs">
                      {draft.charts.map((binding) => <Box key={binding.chartId} className="rounded border border-slate-200 bg-white p-2">
                        <Inline gap="sm" className="items-center justify-between"><Text size="xs" weight="semibold">{binding.chartId}{chartIds.includes(binding.chartId) ? '' : ' · 已不可用'}</Text><Button size="xs" variant="ghost" onClick={() => removeChartBinding(binding.chartId)}>移除</Button></Inline>
                        <Box className="mt-2 grid grid-cols-2 gap-2">
                          {CHART_FIELD_ROLES.map((role) => <Select key={role} aria-label={binding.chartId + ' ' + CHART_FIELD_ROLE_LABELS[role]} sizeVariant="sm" value={binding.fieldMap[role] ?? ''} onChange={(event) => updateChartField(binding.chartId, role, event.currentTarget.value)} options={[{ value: '', label: CHART_FIELD_ROLE_LABELS[role] + '：未映射' }, ...(activeTable?.fields ?? []).map((field) => ({ value: field.id, label: CHART_FIELD_ROLE_LABELS[role] + '：' + field.name }))]} />)}
                        </Box>
                      </Box>)}
                      {draft.charts.length === 0 ? <Text size="xs" tone="muted">尚未绑定图表；当前工作表没有可绑定图表时会保持空列表。</Text> : null}
                    </Stack>
                  </Box>

                  <Box>
                    <FieldLabel>仪表盘布局</FieldLabel>
                    <Box className="grid grid-cols-3 gap-2">
                      <TextInput aria-label="布局列数" type="number" min={1} max={12} value={String(draft.layout.columns)} onChange={(event) => { const value = Number(event.currentTarget.value); if (Number.isFinite(value)) updateDraft((current) => ({ ...current, layout: { ...current.layout, columns: Math.max(1, Math.min(12, Math.trunc(value))) } })); }} />
                      <TextInput aria-label="图表行高" type="number" min={1} value={String(draft.layout.rowHeightPx)} onChange={(event) => { const value = Number(event.currentTarget.value); if (Number.isFinite(value)) updateDraft((current) => ({ ...current, layout: { ...current.layout, rowHeightPx: Math.max(1, value) } })); }} />
                      <TextInput aria-label="图表间距" type="number" min={0} value={String(draft.layout.gapPx)} onChange={(event) => { const value = Number(event.currentTarget.value); if (Number.isFinite(value)) updateDraft((current) => ({ ...current, layout: { ...current.layout, gapPx: Math.max(0, value) } })); }} />
                    </Box>
                    <Text size="xs" tone="muted" className="mt-1">列数 / 行高 / 间距会随分析视图保存，并由仪表盘渲染器统一解释。</Text>
                  </Box>

                  {validationErrors.length > 0 ? <Box className="rounded border border-rose-200 bg-rose-50 px-3 py-2"><Stack gap="xs">{validationErrors.map((error) => <Text key={error} size="xs" tone="danger">{error}</Text>)}</Stack></Box> : null}
                  <Inline gap="sm"><Button size="sm" variant="primary" onClick={applyDraft}>保存分析视图</Button><Button size="sm" variant="ghost" onClick={() => { if (activeView) selectView(activeView); }}>取消草稿</Button></Inline>
                </Stack>
              </PanelBody>
            </Panel>
          ) : null}
        </Stack>
      </PanelBody>
    </Box>
  );
}
