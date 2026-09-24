import React, { useEffect, useMemo, useRef, useState } from 'react';
import { Box, Button, CheckToggle, FileButton, Inline, Panel, PanelBody, PanelFooter, PanelHeader, PanelTitle, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { buildExplicitChartSeries, CHART_SUBTYPES_BY_TYPE, chartStackingForSubtype, defaultChartSubtype, resolveWorksheetChartRanges, type ChartAxisModel, type ChartDrawingPayload, type DrawingObject, type DrawingPayload, type FormulaValue, type RangeRef } from '@react-sheets/core-model';
import type { CommandDescriptor } from '@react-sheets/command-runtime';
import { parseGeoJsonMapResource, type ChartElementSelection } from '@react-sheets/spreadsheet-app';
import { chartLabels, chartSubtypeLabels, chartTypes } from '../chart/chart-labels';
import { chartEditorDraft, chartPayloadFromDraft, chartSeriesDraft, chartSourcePayloadFromDraft, parseChartRange, retargetChartDraft, type ChartEditorDraft } from '../chart/chart-editor-state';
import { ChartSeriesEditor } from '../chart/ChartSeriesEditor';

export interface ChartPanelProps {
  sheetId: string;
  drawings: readonly DrawingObject[];
  drawingPayloads: ReadonlyMap<string, DrawingPayload>;
  selectedDrawingIds?: readonly string[];
  selectedChartElement?: ChartElementSelection | null;
  defaultRange?: string;
  readCellValue: (sourceSheetId: string, row: number, column: number) => FormulaValue;
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
  gridline: '网格线', 'data-table': '数据表', trendline: '趋势线', 'error-bar': '误差线', series: '数据系列', point: '数据点', 'data-label': '数据标签', 'histogram-bin': '直方图分组',
};
const EMPTY_SELECTED_DRAWING_IDS: readonly string[] = [];

export function ChartPanel({ sheetId, drawings, drawingPayloads, selectedDrawingIds = EMPTY_SELECTED_DRAWING_IDS, selectedChartElement = null, defaultRange, readCellValue, onInsertChart, onCommand, onClose }: ChartPanelProps) {
  const selectedDrawingIdSet = useMemo(() => new Set(selectedDrawingIds), [selectedDrawingIds]);
  const entries = useMemo(() => drawings.flatMap(drawing => {
    const payload = drawingPayloads.get(drawing.payloadId);
    return drawing.kind === 'chart' && payload?.kind === 'chart' ? [{ drawing, payload }] : [];
  }), [drawings, drawingPayloads]);
  const current = entries.find(entry => selectedDrawingIdSet.has(entry.drawing.id))?.payload;
  // Editor drafts never become a second workbook state. Only chart.update writes canonical data.
  const [drafts, setDrafts] = useState<Record<string, ChartEditorDraft>>({});
  const [applied, setApplied] = useState<Record<string, ChartDrawingPayload>>({});
  const [createType, setCreateType] = useState<ChartDrawingPayload['chartType']>('column');
  const [createSubtype, setCreateSubtype] = useState<ChartDrawingPayload['subtype']>('clustered');
  const [createTitle, setCreateTitle] = useState('');
  const [createRange, setCreateRange] = useState(defaultRange ?? '');
  const [message, setMessage] = useState<string | null>(null);
  const id = current?.chartId;
  const selectedChartIdRef = useRef(id);
  const currentPayloadRef = useRef(current);
  const mapImportRef = useRef<symbol | undefined>(undefined);
  selectedChartIdRef.current = id;
  currentPayloadRef.current = current;
  const previousDefaultRangeRef = useRef(defaultRange ?? '');
  useEffect(() => {
    if (id) return;
    const nextDefault = defaultRange ?? '';
    setCreateRange(value => !value.trim() || value === previousDefaultRangeRef.current ? nextDefault : value);
    previousDefaultRangeRef.current = nextDefault;
  }, [defaultRange, id]);
  const chartIdsKey = useMemo(() => entries.map((entry) => entry.payload.chartId).join('\u0000'), [entries]);
  useEffect(() => {
    mapImportRef.current = undefined;
    const liveIds = new Set(chartIdsKey ? chartIdsKey.split('\u0000') : []);
    setDrafts(all => Object.fromEntries(Object.entries(all).filter(([chartId]) => liveIds.has(chartId))));
    setApplied(all => Object.fromEntries(Object.entries(all).filter(([chartId]) => liveIds.has(chartId))));
  }, [chartIdsKey]);
  const submitted = id ? applied[id] : undefined;
  const currentFingerprint = useMemo(() => current ? JSON.stringify(current) : '', [current]);
  const submittedFingerprint = useMemo(() => submitted ? JSON.stringify(submitted) : '', [submitted]);
  const submittedApplied = Boolean(current && submitted && currentFingerprint === submittedFingerprint);
  useEffect(() => {
    if (!id || !submittedApplied) return;
    setDrafts(all => { const next = { ...all }; delete next[id]; return next; });
    setApplied(all => { const next = { ...all }; delete next[id]; return next; });
  }, [id, submittedApplied]);
  const storedDraft = id && !submittedApplied ? drafts[id] : undefined;
  const cleanDraft = useMemo(() => current ? chartEditorDraft(current) : undefined, [current]);
  const draft = current ? storedDraft ?? cleanDraft : undefined;
  const payload = draft?.value;
  const storedBaseFingerprint = useMemo(() => storedDraft ? JSON.stringify(storedDraft.base) : '', [storedDraft]);
  const conflict = Boolean(storedDraft && current && storedBaseFingerprint !== currentFingerprint);
  const preserved = current?.nativeIdentity?.status === 'preserved-native';
  const validation = useMemo((): { candidate?: ChartDrawingPayload; error: string | null } => {
    if (!draft) return { error: null };
    try { return { candidate: chartPayloadFromDraft(draft, sheetId), error: null }; }
    catch (error) { return { error: error instanceof Error ? error.message : '图表输入无效' }; }
  }, [draft, sheetId]);
  const candidate = validation.candidate;
  const validationError = validation.error;
  const candidateFingerprint = useMemo(() => candidate ? JSON.stringify(candidate) : '', [candidate]);
  const dirty = Boolean(storedDraft && (!candidate || candidateFingerprint !== currentFingerprint));
  const edit = (change: (value: ChartEditorDraft) => ChartEditorDraft) => {
    if (!id || !current) return;
    setDrafts(all => {
      const latestPayload = currentPayloadRef.current;
      if (!latestPayload || latestPayload.chartId !== id) return all;
      const latest = all[id] ?? chartEditorDraft(latestPayload);
      return { ...all, [id]: change(structuredClone(latest)) };
    });
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
    try {
      onCommand({ commandId: 'chart.update', params: { sheetId, chartId: id, payload: candidate, expectedPayload: draft.base } });
      setApplied(all => ({ ...all, [id]: candidate }));
      setMessage(null);
    } catch (error) { setMessage(error instanceof Error ? error.message : '图表更新失败'); }
  };
  const create = () => {
    try {
      const range = parseChartRange(createRange, sheetId, '数据区域');
      onInsertChart(createType, createSubtype, range, createTitle, chartStackingForSubtype(createSubtype) ?? 'none');
      setMessage(null);
    } catch (error) { setMessage(error instanceof Error ? error.message : '创建图表失败'); }
  };
  const activeChartElement = selectedChartElement?.chartId === id ? selectedChartElement : null;
  const selectedSeriesId = activeChartElement && 'seriesId' in activeChartElement ? activeChartElement.seriesId : undefined;
  return <Panel className="flex h-full min-h-0 flex-col overflow-hidden rounded-none border-0 bg-slate-50/70 shadow-none" onKeyDown={event => {
    if (event.key === 'Escape' && dirty) { event.preventDefault(); event.stopPropagation(); cancel(); }
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter' && dirty && !validationError) { event.preventDefault(); apply(); }
  }}>
    <PanelHeader className="shrink-0 border-b border-slate-200 bg-white px-4 py-3"><Stack gap="xs"><PanelTitle size="sm">{current ? '图表设计与格式' : '插入图表'}</PanelTitle><Text size="xs" tone="muted">{current ? '编辑草稿后应用 · Ctrl+Enter 应用 · Esc 取消' : '从选定的数据区域创建图表'}</Text></Stack>{onClose ? <Button icon="x" iconOnly aria-label="关闭图表面板" variant="ghost" size="sm" onClick={onClose} /> : null}</PanelHeader>
    <PanelBody className="min-h-0 flex-1 overflow-y-auto p-3"><Stack gap="sm">
      {activeChartElement ? <Box className="rounded-lg border border-emerald-200 bg-emerald-50 p-3"><Text size="xs" className="text-emerald-800">当前选择：{elementLabels[activeChartElement.kind]}{selectedSeriesId ? ' · ' + selectedSeriesId : ''}</Text></Box> : null}
      {conflict ? <Box role="alert" className="rounded-lg border border-amber-300 bg-amber-50 p-3 text-xs text-amber-900">图表已被其他操作修改。你的草稿仍保留；请核对当前图表，取消草稿后重新编辑。</Box> : null}
      {preserved ? <Box role="status" className="rounded-lg bg-amber-50 p-3 text-xs">该原生图表尚未支持编辑，原始内容将保留。</Box> : null}
      {validationError && dirty ? <Box role="alert" className="rounded-lg bg-red-50 p-3 text-xs text-red-700">{validationError}</Box> : null}
      {message ? <Box role="alert" className="rounded-lg bg-red-50 p-3 text-xs text-red-700">{message}</Box> : null}
      {payload && draft ? <fieldset disabled={preserved} className="min-w-0 space-y-3 border-0 p-0">
        <Group title="图表与数据" open>
          <Field label="图表标题"><TextInput aria-label="图表标题" value={payload.elements.title ?? ''} placeholder="输入标题" onChange={event => updateElements({ title: event.target.value })} /></Field>
          <TypeFields value={payload} onChange={(chartType, subtype) => edit(value => retargetChartDraft(value, chartType, subtype))} />
          {payload.chartType === 'map' ? <Group title="离线地图资源" open>
            <Text size="xs" tone="muted">地图只读取工作簿内已校验的 GeoJSON，不访问外部地图服务。</Text>
            {payload.mapOptions?.resource ? <Inline className="items-center justify-between rounded-md bg-slate-100 px-2 py-1.5" gap="xs"><Text size="xs">{payload.mapOptions.resource.resourceId} · {payload.mapOptions.resource.features.length} 个区域</Text><Button size="xs" variant="ghost" onClick={() => updatePayload({ mapOptions: { ...(payload.mapOptions ?? { geography: 'country-region', mapArea: 'automatic', labelLevel: 'best-fit', colorScale: 'sequential' }), resource: undefined } })}>移除</Button></Inline> : <Text size="xs" tone="muted">尚未导入 GeoJSON。没有资源时地图会明确显示不可用。</Text>}
            <FileButton accept=".geojson,application/geo+json,.json" icon="chart" size="sm" variant="secondary" onFile={(file) => {
              const targetChartId = id;
              const request = Symbol('map-import');
              mapImportRef.current = request;
              void file.text().then((text) => parseGeoJsonMapResource(text, file.name)).then((resource) => {
                if (!targetChartId || selectedChartIdRef.current !== targetChartId || mapImportRef.current !== request) return;
                edit(value => ({ ...value, value: { ...value.value, mapOptions: { ...(value.value.mapOptions ?? { geography: 'country-region', mapArea: 'automatic', labelLevel: 'best-fit', colorScale: 'sequential' }), resource } } }));
                mapImportRef.current = undefined;
                setMessage(null);
              }).catch((error) => { if (selectedChartIdRef.current === targetChartId && mapImportRef.current === request) { mapImportRef.current = undefined; setMessage(error instanceof Error ? error.message : 'GeoJSON 地图资源无效'); } });
            }}>导入 GeoJSON</FileButton>
          </Group> : null}
          {payload.source.kind === 'worksheet-ranges' ? <Field label="数据区域（多个区域用分号分隔）"><TextInput aria-label="图表数据区域" value={draft.sourceRanges} onChange={event => edit(value => ({ ...value, sourceRanges: event.target.value }))} /></Field>
            : payload.source.kind === 'report-range' ? <Field label="报表数据区域"><TextInput aria-label="图表数据区域" value={draft.sourceRanges} onChange={event => edit(value => ({ ...value, sourceRanges: event.target.value }))} /></Field>
              : <Text size="xs" tone="muted">数据绑定：{payload.source.kind === 'pivot' ? '透视结果' : 'Table'}，请在对应数据源中调整范围。</Text>}
          {payload.source.kind === 'worksheet-ranges' ? <><Field label="分类标签范围"><TextInput aria-label="分类标签范围" placeholder="自动，或 A2:A20" value={draft.categoryRange} onChange={event => edit(value => ({ ...value, categoryRange: event.target.value }))} /></Field>
          <Button size="sm" variant="secondary" onClick={() => updatePayload({ dataOrientation: payload.dataOrientation === 'rows' ? 'columns' : 'rows' })}>切换行／列（当前按{payload.dataOrientation === 'rows' ? '行' : '列'}）</Button></> : null}
        </Group>
        <Group title={'数据系列（' + (draft.series.length || '自动') + '）'} open={Boolean(selectedSeriesId)}>
          <ChartSeriesEditor chartType={payload.chartType} subtype={payload.subtype} series={draft.series} selectedSeriesId={selectedSeriesId} onChange={series => edit(value => ({ ...value, series }))} canAdd={payload.source.kind === 'worksheet-ranges'} allowEmpty={!['scatter', 'bubble', 'stock', 'combo'].includes(payload.chartType)} onMaterialize={payload.source.kind === 'worksheet-ranges' && !draft.series.length ? () => {
            try {
              const candidate = chartSourcePayloadFromDraft(draft, sheetId);
              if (candidate.source.kind !== 'worksheet-ranges') throw new Error('INVALID_CHART_SOURCE: 当前图表没有工作表数据区域');
              const explicit = candidate.source.ranges.flatMap((range, rangeIndex) => (buildExplicitChartSeries(candidate.chartType, candidate.subtype, range, candidate.dataOrientation ?? 'columns') ?? []).map((entry, seriesIndex) => ({
                ...entry,
                id: candidate.source.kind === 'worksheet-ranges' && candidate.source.ranges.length === 1 ? entry.id : `${entry.id ?? 'series'}:${rangeIndex + 1}:${seriesIndex + 1}`,
              })));
              const bindings = explicit.length ? { series: explicit } : resolveWorksheetChartRanges(candidate, range => readCellValue(range.sheetId, range.startRow, range.startColumn));
              const series = bindings.series.map(entry => chartSeriesDraft(entry));
              edit(value => ({ ...value, series }));
            } catch (error) { setMessage(error instanceof Error ? error.message : '无法生成可编辑系列'); }
          } : undefined} onAdd={() => {
            const existing = draft.series[draft.series.length - 1];
            if (!existing) { setMessage('请先生成可编辑系列'); return; }
            const series = structuredClone(existing);
            series.value.id = crypto.randomUUID();
            series.value.name = '系列 ' + (draft.series.length + 1);
            edit(value => ({ ...value, series: [...value.series, series] }));
          }} />
        </Group>
        <Group title="图例与标签" open={activeChartElement?.kind === 'legend' || activeChartElement?.kind === 'data-label'}>
          <Field label="图例位置"><Select aria-label="图例位置" value={payload.elements.legend?.visible ? payload.elements.legend.position : 'none'} onChange={event => updateElements({ legend: { ...payload.elements.legend, visible: event.target.value !== 'none', position: event.target.value === 'none' ? 'bottom' : event.target.value as NonNullable<ChartDrawingPayload['elements']['legend']>['position'] } })}><option value="none">隐藏</option><option value="top">上方</option><option value="bottom">下方</option><option value="left">左侧</option><option value="right">右侧</option><option value="top-right">右上方</option></Select></Field>
          {(['visible', 'showValue', 'showCategoryName', 'showSeriesName', 'showPercentage'] as const).map((key, index) => <CheckToggle key={key} label={['显示数据标签', '显示数值', '显示分类', '显示系列名称', '显示百分比'][index]!} checked={payload.elements.dataLabels?.[key] === true} onChange={event => {
            const checked = event.currentTarget.checked;
            updateElements({ dataLabels: { ...payload.elements.dataLabels, visible: key === 'visible' ? checked : checked ? true : payload.elements.dataLabels?.visible ?? false, [key]: checked } });
          }} />)}
          <TextInput aria-label="数据标签数字格式" placeholder="数字格式，例如 0.00" value={payload.elements.dataLabels?.numberFormat ?? ''} onChange={event => updateElements({ dataLabels: { visible: true, ...payload.elements.dataLabels, numberFormat: event.target.value || undefined } })} />
          <CheckToggle label="显示图表数据表" checked={payload.elements.dataTable?.visible === true} onChange={event => updateElements({ dataTable: { ...payload.elements.dataTable, visible: event.currentTarget.checked } })} />
        </Group>
        <Group title="坐标轴" open={activeChartElement?.kind === 'axis' || activeChartElement?.kind === 'axis-title'}>
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
        {entries.map(({ payload: item }) => <Box key={item.chartId} className="flex items-center justify-between gap-2"><Text size="xs">{item.elements.title || chartLabels[item.chartType]}{drafts[item.chartId] && !applied[item.chartId] ? ' · 草稿' : ''}</Text><Button size="xs" variant="ghost" icon="trash" iconOnly aria-label={'删除图表 ' + (item.elements.title || item.chartId)} onClick={() => { try { onCommand({ commandId: 'chart.remove', params: { sheetId, chartId: item.chartId } }); setMessage(null); } catch (error) { setMessage(error instanceof Error ? error.message : '删除图表失败'); } }} /></Box>)}
      </Group> : null}
    </Stack></PanelBody>
    <PanelFooter className="shrink-0 border-t border-slate-200 bg-white px-3 py-3"><Inline className="justify-between" gap="xs">
      <Text size="xs" tone="muted">{dirty ? '有未应用的草稿' : '修改后统一应用'}</Text>
      <Inline gap="xs">{current ? <><Button size="sm" variant="ghost" disabled={!storedDraft} onClick={cancel}>取消</Button><Button size="sm" variant="primary" disabled={!dirty || Boolean(validationError) || conflict || preserved} onClick={apply}>应用</Button></> : null}</Inline>
    </Inline></PanelFooter>
  </Panel>;
}
