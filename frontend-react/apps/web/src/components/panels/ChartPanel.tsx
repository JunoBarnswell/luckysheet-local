import React, { useEffect, useState } from 'react';
import { Box, Button, CheckToggle, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { CHART_SUBTYPES_BY_TYPE, chartStackingForSubtype, defaultChartSubtype, type ChartAxisModel, type ChartDrawingPayload, type DrawingObject, type DrawingPayload, type RangeRef } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import type { ChartElementSelection } from '@react-sheets/spreadsheet-app';
import { chartLabels, chartSubtypeLabels, chartTypes } from '../chart/chart-labels';
import { chartEditorDraft, chartPayloadFromDraft, chartSeriesDraft, parseChartRange, type ChartEditorDraft } from '../chart/chart-editor-state';
import { ChartSeriesEditor } from '../chart/ChartSeriesEditor';

export interface ChartPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  selectedChartElement?: ChartElementSelection | null;
  defaultRange?: string;
  onInsertChart: (type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype'], sourceRange: RangeRef, title: string, stacked: NonNullable<ChartDrawingPayload['stacked']>) => void;
  onCommand: (descriptor: CommandDescriptor) => void;
  onClose?: () => void;
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return <Box className="space-y-1.5"><Text size="xs" weight="medium" className="text-slate-600">{label}</Text>{children}</Box>;
}
function Group({ title, children, open = false }: { title: string; children: React.ReactNode; open?: boolean }) {
  return <details open={open} className="rounded-lg border border-slate-200 bg-white">
    <summary className="cursor-pointer select-none px-3 py-2.5 text-sm font-semibold text-slate-700 focus-visible:outline-emerald-600">{title}</summary>
    <Box className="space-y-3 border-t border-slate-100 p-3">{children}</Box>
  </details>;
}
function TypeFields({ value, onChange }: { value: Pick<ChartDrawingPayload, 'chartType' | 'subtype'>; onChange: (type: ChartDrawingPayload['chartType'], subtype: ChartDrawingPayload['subtype']) => void }) {
  return <Box className="grid grid-cols-2 gap-2">
    <Field label="图表类型"><Select aria-label="图表类型" value={value.chartType} onChange={event => { const type = event.target.value as ChartDrawingPayload['chartType']; onChange(type, defaultChartSubtype(type)); }}>{chartTypes.map(type => <option key={type} value={type}>{chartLabels[type]}</option>)}</Select></Field>
    <Field label="图表子类型"><Select aria-label="图表子类型" value={value.subtype} onChange={event => onChange(value.chartType, event.target.value as ChartDrawingPayload['subtype'])}>{CHART_SUBTYPES_BY_TYPE[value.chartType].map(type => <option key={type} value={type}>{chartSubtypeLabels[type]}</option>)}</Select></Field>
  </Box>;
}
function AxisFields({ axis, position, onChange }: { axis?: ChartAxisModel; position: ChartAxisModel['position']; onChange: (axis: ChartAxisModel) => void }) {
  const value: ChartAxisModel = axis ?? { id: position, position, axisType: position === 'bottom' ? 'category' : 'value' };
  const update = (patch: Partial<ChartAxisModel>) => onChange({ ...value, ...patch });
  return <Stack gap="sm">
    <TextInput aria-label={position + ' 轴标题'} placeholder="轴标题" value={value.title ?? ''} onChange={event => update({ title: event.target.value })} />
    {position !== 'bottom' ? <>
      <Select aria-label={position + ' 轴刻度'} value={value.scale ?? 'linear'} onChange={event => update({ scale: event.target.value as ChartAxisModel['scale'] })}><option value="linear">线性刻度</option><option value="logarithmic">对数刻度</option></Select>
      <Box className="grid grid-cols-2 gap-2"><TextInput aria-label={position + ' 轴最小值'} type="number" placeholder="自动最小值" value={value.minimum ?? ''} onChange={event => update({ minimum: event.target.value === '' ? undefined : Number(event.target.value), automaticMinimum: event.target.value === '' })} /><TextInput aria-label={position + ' 轴最大值'} type="number" placeholder="自动最大值" value={value.maximum ?? ''} onChange={event => update({ maximum: event.target.value === '' ? undefined : Number(event.target.value), automaticMaximum: event.target.value === '' })} /></Box>
    </> : null}
    <TextInput aria-label={position + ' 轴数字格式'} placeholder="数字格式，例如 0.00%" value={value.numberFormat ?? ''} onChange={event => update({ numberFormat: event.target.value || undefined })} />
    <CheckToggle label="显示主要网格线" checked={value.majorGridlines?.visible === true} onChange={event => update({ majorGridlines: { ...value.majorGridlines, visible: event.currentTarget.checked } })} />
  </Stack>;
}
const elementLabels: Record<ChartElementSelection['kind'], string> = {
  'chart-area': '图表区', 'plot-area': '绘图区', title: '标题', legend: '图例', axis: '坐标轴', 'axis-title': '轴标题',
  gridline: '网格线', 'data-table': '数据表', trendline: '趋势线', 'error-bar': '误差线', series: '数据系列', point: '数据点', 'data-label': '数据标签',
};

export function ChartPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = [], selectedChartElement = null, defaultRange, onInsertChart, onCommand, onClose }: ChartPanelProps) {
  const entries = drawings.flatMap(drawing => {
    const payload = drawingPayloads.get(drawing.payloadId);
    return drawing.kind === 'chart' && payload?.kind === 'chart' ? [{ drawing, payload }] : [];
  });
  const current = entries.find(entry => selectedDrawingIds.includes(entry.drawing.id))?.payload;
  // Editor drafts never become a second workbook state. Only chart.update writes canonical data.
  const [drafts, setDrafts] = useState<Record<string, ChartEditorDraft>>({});
  const [applied, setApplied] = useState<Record<string, ChartDrawingPayload>>({});
  const [createType, setCreateType] = useState<ChartDrawingPayload['chartType']>('column');
  const [createSubtype, setCreateSubtype] = useState<ChartDrawingPayload['subtype']>('clustered');
  const [createTitle, setCreateTitle] = useState('');
  const [createRange, setCreateRange] = useState(defaultRange ?? '');
  const [message, setMessage] = useState<string | null>(null);
  const id = current?.chartId;
  const submitted = id ? applied[id] : undefined;
  const submittedApplied = current && submitted && JSON.stringify(current) === JSON.stringify(submitted);
  useEffect(() => {
    if (!id || !submittedApplied) return;
    setDrafts(all => { const next = { ...all }; delete next[id]; return next; });
    setApplied(all => { const next = { ...all }; delete next[id]; return next; });
  }, [id, submittedApplied]);
  const storedDraft = id && !submittedApplied ? drafts[id] : undefined;
  const draft = current ? storedDraft ?? chartEditorDraft(current) : undefined;
  const payload = draft?.value;
  const conflict = Boolean(storedDraft && current && JSON.stringify(storedDraft.base) !== JSON.stringify(current));
  const preserved = current?.nativeIdentity?.status === 'preserved-native';
  let validationError: string | null = null;
  let candidate: ChartDrawingPayload | undefined;
  if (draft) {
    try { candidate = chartPayloadFromDraft(draft, sheetId); }
    catch (error) { validationError = error instanceof Error ? error.message : '图表输入无效'; }
  }
  const dirty = Boolean(storedDraft && (!candidate || JSON.stringify(candidate) !== JSON.stringify(current)));
  const edit = (change: (value: ChartEditorDraft) => ChartEditorDraft) => {
    if (!id || !draft) return;
    setDrafts(all => ({ ...all, [id]: change(structuredClone(draft)) }));
    setApplied(all => { const next = { ...all }; delete next[id]; return next; });
    setMessage(null);
  };
  const updatePayload = (patch: Partial<ChartDrawingPayload>) => edit(value => ({ ...value, value: { ...value.value, ...patch } }));
  const updateElements = (patch: Partial<ChartDrawingPayload['elements']>) => edit(value => ({ ...value, value: { ...value.value, elements: { ...value.value.elements, ...patch } } }));
  const cancel = () => {
    if (!id) return;
    setDrafts(all => { const next = { ...all }; delete next[id]; return next; });
    setApplied(all => { const next = { ...all }; delete next[id]; return next; });
    setMessage(null);
  };
  const apply = () => {
    if (!draft || !candidate || !id || conflict || preserved) return;
    onCommand({ commandId: 'chart.update', params: { sheetId, chartId: id, payload: candidate, expectedPayload: draft.base } });
    setApplied(all => ({ ...all, [id]: candidate }));
  };
  const create = () => {
    try {
      const range = parseChartRange(createRange, sheetId, '数据区域');
      onInsertChart(createType, createSubtype, range, createTitle, chartStackingForSubtype(createSubtype) ?? 'none');
      setMessage(null);
    } catch (error) { setMessage(error instanceof Error ? error.message : '创建图表失败'); }
  };
  const selectedSeriesId = selectedChartElement && 'seriesId' in selectedChartElement ? selectedChartElement.seriesId : undefined;
  return <Panel className="flex h-full min-h-0 flex-col overflow-hidden rounded-none border-0 bg-slate-50/70 shadow-none" onKeyDown={event => {
    if (event.key === 'Escape' && dirty) { event.preventDefault(); event.stopPropagation(); cancel(); }
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter' && dirty && !validationError) { event.preventDefault(); apply(); }
  }}>
    <PanelHeader className="shrink-0 border-b border-slate-200 bg-white px-4 py-3"><Stack gap="xs"><PanelTitle size="sm">{current ? '图表设计与格式' : '插入图表'}</PanelTitle><Text size="xs" tone="muted">{current ? '编辑草稿后应用 · Ctrl+Enter 应用 · Esc 取消' : '从选定的数据区域创建图表'}</Text></Stack>{onClose ? <Button icon="x" iconOnly aria-label="关闭图表面板" variant="ghost" size="sm" onClick={onClose} /> : null}</PanelHeader>
    <PanelBody className="min-h-0 flex-1 overflow-y-auto p-3"><Stack gap="sm">
      {selectedChartElement ? <Box className="rounded-lg border border-emerald-200 bg-emerald-50 p-3"><Text size="xs" className="text-emerald-800">当前选择：{elementLabels[selectedChartElement.kind]}{selectedSeriesId ? ' · ' + selectedSeriesId : ''}</Text></Box> : null}
      {conflict ? <Box role="alert" className="rounded-lg border border-amber-300 bg-amber-50 p-3 text-xs text-amber-900">图表已被其他操作修改。你的草稿仍保留；请核对当前图表，取消草稿后重新编辑。</Box> : null}
      {preserved ? <Box role="status" className="rounded-lg bg-amber-50 p-3 text-xs">该原生图表尚未支持编辑，原始内容将保留。</Box> : null}
      {validationError && dirty ? <Box role="alert" className="rounded-lg bg-red-50 p-3 text-xs text-red-700">{validationError}</Box> : null}
      {message ? <Box role="alert" className="rounded-lg bg-red-50 p-3 text-xs text-red-700">{message}</Box> : null}
      {payload && draft ? <fieldset disabled={preserved} className="min-w-0 space-y-3 border-0 p-0">
        <Group title="图表与数据" open>
          <Field label="图表标题"><TextInput aria-label="图表标题" value={payload.elements.title ?? ''} placeholder="输入标题" onChange={event => updateElements({ title: event.target.value })} /></Field>
          <TypeFields value={payload} onChange={(chartType, subtype) => edit(value => ({ ...value, value: { ...value.value, chartType, subtype, stacked: chartStackingForSubtype(subtype) }, series: value.series.map(entry => ({ ...entry, value: { ...entry.value, chartType: chartType === 'combo' ? (entry.value.chartType && ['column', 'bar', 'line', 'area'].includes(entry.value.chartType) ? entry.value.chartType : 'column') : chartType, subtype: chartType === 'combo' ? undefined : subtype } })) }))} />
          {payload.source.kind === 'worksheet-ranges' || payload.source.kind === 'report-range' ? <Field label="数据区域（多个区域用分号分隔）"><TextInput aria-label="图表数据区域" value={draft.sourceRanges} onChange={event => edit(value => ({ ...value, sourceRanges: event.target.value }))} /></Field> : <Text size="xs" tone="muted">数据绑定：{payload.source.kind === 'pivot' ? '透视结果' : 'Table'}，请在对应数据源中调整范围。</Text>}
          <Field label="分类标签范围"><TextInput aria-label="分类标签范围" placeholder="自动，或 A2:A20" value={draft.categoryRange} onChange={event => edit(value => ({ ...value, categoryRange: event.target.value }))} /></Field>
          <Button size="sm" variant="secondary" disabled={payload.source.kind === 'pivot'} onClick={() => updatePayload({ dataOrientation: payload.dataOrientation === 'rows' ? 'columns' : 'rows' })}>切换行／列（当前按{payload.dataOrientation === 'rows' ? '行' : '列'}）</Button>
        </Group>
        <Group title={'数据系列（' + (draft.series.length || '自动') + '）'} open={Boolean(selectedSeriesId)}>
          <ChartSeriesEditor chartType={payload.chartType} series={draft.series} selectedSeriesId={selectedSeriesId} onChange={series => edit(value => ({ ...value, series }))} canAdd={payload.source.kind === 'worksheet-ranges'} onAdd={() => {
            try {
              const range = parseChartRange(draft.sourceRanges.split(';')[0] ?? '', sheetId, '数据区域');
              const series = chartSeriesDraft({ id: crypto.randomUUID(), name: '系列 ' + (draft.series.length + 1), range, chartType: payload.chartType === 'combo' ? 'column' : payload.chartType });
              edit(value => ({ ...value, series: [...value.series, series] }));
            } catch (error) { setMessage(error instanceof Error ? error.message : '数据范围无效'); }
          }} />
        </Group>
        <Group title="图例与标签" open={selectedChartElement?.kind === 'legend' || selectedChartElement?.kind === 'data-label'}>
          <Field label="图例位置"><Select aria-label="图例位置" value={payload.elements.legend?.visible ? payload.elements.legend.position : 'none'} onChange={event => updateElements({ legend: { ...payload.elements.legend, visible: event.target.value !== 'none', position: event.target.value === 'none' ? 'bottom' : event.target.value as NonNullable<ChartDrawingPayload['elements']['legend']>['position'] } })}><option value="none">隐藏</option><option value="top">上方</option><option value="bottom">下方</option><option value="left">左侧</option><option value="right">右侧</option><option value="top-right">右上方</option></Select></Field>
          {(['visible', 'showValue', 'showCategoryName', 'showSeriesName', 'showPercentage'] as const).map((key, index) => <CheckToggle key={key} label={['显示数据标签', '显示数值', '显示分类', '显示系列名称', '显示百分比'][index]!} checked={payload.elements.dataLabels?.[key] === true} onChange={event => updateElements({ dataLabels: { visible: false, ...payload.elements.dataLabels, [key]: event.currentTarget.checked } })} />)}
          <TextInput aria-label="数据标签数字格式" placeholder="数字格式，例如 0.00" value={payload.elements.dataLabels?.numberFormat ?? ''} onChange={event => updateElements({ dataLabels: { visible: true, ...payload.elements.dataLabels, numberFormat: event.target.value || undefined } })} />
          <CheckToggle label="显示图表数据表" checked={payload.elements.dataTable?.visible === true} onChange={event => updateElements({ dataTable: { ...payload.elements.dataTable, visible: event.currentTarget.checked } })} />
        </Group>
        <Group title="坐标轴" open={selectedChartElement?.kind === 'axis' || selectedChartElement?.kind === 'axis-title'}>
          {(['categoryAxis', 'valueAxis', 'secondaryValueAxis'] as const).map((key, index) => <Box key={key} className="space-y-2 border-b border-slate-100 pb-3 last:border-0"><Text size="xs" weight="semibold">{['分类轴', '主数值轴', '次数值轴'][index]}</Text><AxisFields axis={payload.elements[key]} position={key === 'categoryAxis' ? 'bottom' : key === 'valueAxis' ? 'left' : 'right'} onChange={axis => updateElements({ [key]: axis })} /></Box>)}
        </Group>
        <Group title="外观与空值">
          <Field label="图表背景"><TextInput aria-label="图表背景" placeholder="自动，例如 #ffffff" value={typeof payload.elements.chartArea?.fill === 'string' ? payload.elements.chartArea.fill : payload.elements.chartArea?.fill?.color ?? ''} onChange={event => updateElements({ chartArea: { ...payload.elements.chartArea, fill: event.target.value || undefined } })} /></Field>
          <Field label="图表边框"><TextInput aria-label="图表边框" placeholder="自动，例如 #cbd5e1" value={payload.elements.chartArea?.border ?? ''} onChange={event => updateElements({ chartArea: { ...payload.elements.chartArea, border: event.target.value || undefined } })} /></Field>
          <Select aria-label="隐藏数据" value={payload.elements.hiddenData} onChange={event => updateElements({ hiddenData: event.target.value as ChartDrawingPayload['elements']['hiddenData'] })}><option value="show">显示全部数据</option><option value="hideRows">忽略隐藏行</option><option value="hideColumns">忽略隐藏列</option></Select>
          <Select aria-label="空值处理" value={payload.elements.emptyCells ?? 'gap'} onChange={event => updateElements({ emptyCells: event.target.value as ChartDrawingPayload['elements']['emptyCells'] })}><option value="gap">空值留空</option><option value="zero">空值显示为零</option><option value="connect">跨空值连接</option></Select>
        </Group>
      </fieldset> : <Group title="新建图表" open>
        <Field label="图表标题"><TextInput aria-label="图表标题" value={createTitle} placeholder="输入标题" onChange={event => setCreateTitle(event.target.value)} /></Field>
        <TypeFields value={{ chartType: createType, subtype: createSubtype }} onChange={(type, subtype) => { setCreateType(type); setCreateSubtype(subtype); }} />
        <Field label="数据区域"><TextInput aria-label="图表数据区域" value={createRange} placeholder="例如 A1:C20" onChange={event => setCreateRange(event.target.value)} /></Field>
        <Button size="sm" variant="primary" icon="plus" disabled={!createRange.trim()} onClick={create}>插入图表</Button>
      </Group>}
      {entries.length ? <Group title={'工作表中的图表（' + entries.length + '）'}>
        {entries.map(({ payload: item }) => <Box key={item.chartId} className="flex items-center justify-between gap-2"><Text size="xs">{item.elements.title || chartLabels[item.chartType]}{drafts[item.chartId] && !applied[item.chartId] ? ' · 草稿' : ''}</Text><Button size="xs" variant="ghost" icon="trash" iconOnly aria-label={'删除图表 ' + (item.elements.title || item.chartId)} onClick={() => onCommand({ commandId: 'chart.remove', params: { sheetId, chartId: item.chartId } })} /></Box>)}
      </Group> : null}
    </Stack></PanelBody>
    <PanelFooter className="shrink-0 border-t border-slate-200 bg-white px-3 py-3"><Inline className="justify-between" gap="xs">
      <Text size="xs" tone="muted">{dirty ? '有未应用的草稿' : '修改后统一应用'}</Text>
      <Inline gap="xs">{current ? <><Button size="sm" variant="ghost" disabled={!storedDraft} onClick={cancel}>取消</Button><Button size="sm" variant="primary" disabled={!dirty || Boolean(validationError) || conflict || preserved} onClick={apply}>应用</Button></> : null}</Inline>
    </Inline></PanelFooter>
  </Panel>;
}
