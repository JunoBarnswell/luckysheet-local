import { Box, Button, CheckToggle, Inline, Select, Stack, Text, TextInput } from '@react-sheets/ui-system';
import { defaultChartSubtype, type ChartDrawingPayload } from '@react-sheets/core-model';
import { chartLabels } from './chart-labels';
import type { ChartSeriesDraft } from './chart-editor-state';

export interface ChartSeriesEditorProps {
  series: readonly ChartSeriesDraft[];
  chartType: ChartDrawingPayload['chartType'];
  selectedSeriesId?: string;
  onChange: (series: ChartSeriesDraft[]) => void;
  onAdd: () => void;
  canAdd: boolean;
}

export function ChartSeriesEditor({ series, chartType, selectedSeriesId, onChange, onAdd, canAdd }: ChartSeriesEditorProps) {
  const update = (index: number, change: (entry: ChartSeriesDraft) => ChartSeriesDraft) => onChange(series.map((entry, position) => position === index ? change(entry) : entry));
  const updateErrorType = (index: number, type: NonNullable<ChartSeriesDraft['value']['errorBars']>['type'] | 'none') => update(index, item => {
    if (type === 'none') return { ...item, value: { ...item.value, errorBars: undefined }, errorPlusRange: '', errorMinusRange: '' };
    const current = item.value.errorBars;
    return {
      ...item,
      value: {
        ...item.value,
        errorBars: {
          type,
          direction: current?.direction ?? 'vertical',
          endStyle: current?.endStyle ?? 'cap',
          ...(type === 'fixed' || type === 'percentage' ? { value: current?.value ?? 1 } : {}),
        },
      },
    };
  });
  const move = (index: number, offset: number) => {
    const next = [...series];
    const [entry] = next.splice(index, 1);
    if (entry) next.splice(index + offset, 0, entry);
    onChange(next);
  };
  return <Stack gap="sm">
    {!series.length ? <Text size="xs" tone="muted">系列由数据区域或透视结果自动生成。</Text> : null}
    {series.map((entry, index) => <Box key={entry.value.id ?? index} className={`rounded-lg border p-3 ${entry.value.id === selectedSeriesId ? 'border-emerald-500 bg-emerald-50/50' : 'border-slate-200 bg-white'}`}>
      <Stack gap="sm">
        <Inline className="justify-between"><Text size="xs" weight="semibold">系列 {index + 1}</Text><Inline gap="xs">
          <Button aria-label={`上移系列 ${index + 1}`} icon="arrow-up" iconOnly size="xs" variant="ghost" disabled={index === 0} onClick={() => move(index, -1)} />
          <Button aria-label={`下移系列 ${index + 1}`} icon="arrow-down" iconOnly size="xs" variant="ghost" disabled={index === series.length - 1} onClick={() => move(index, 1)} />
          <Button aria-label={`删除系列 ${index + 1}`} icon="trash" iconOnly size="xs" variant="ghost" disabled={series.length === 1} onClick={() => onChange(series.filter((_, position) => position !== index))} />
        </Inline></Inline>
        <TextInput aria-label={`系列 ${index + 1} 名称`} value={entry.value.name} onChange={event => update(index, item => ({ ...item, value: { ...item.value, name: event.target.value } }))} />
        <TextInput aria-label={`系列 ${index + 1} 范围`} value={entry.range} placeholder="数值范围，例如 B2:B20" onChange={event => update(index, item => ({ ...item, range: event.target.value }))} />
        {chartType === 'combo' ? <Select aria-label={`系列 ${index + 1} 图表类型`} value={entry.value.chartType ?? 'column'} onChange={event => {
          const type = event.target.value as 'column' | 'bar' | 'line' | 'area';
          update(index, item => ({ ...item, value: { ...item.value, chartType: type, subtype: defaultChartSubtype(type) } }));
        }}>{(['column', 'bar', 'line', 'area'] as const).map(type => <option key={type} value={type}>{chartLabels[type]}</option>)}</Select> : null}
        <Inline gap="xs"><Select aria-label={`系列 ${index + 1} 坐标轴`} value={entry.value.axis ?? 'primary'} onChange={event => update(index, item => ({ ...item, value: { ...item.value, axis: event.target.value as 'primary' | 'secondary' } }))}><option value="primary">主坐标轴</option><option value="secondary">次坐标轴</option></Select>
          <TextInput aria-label={`系列 ${index + 1} 颜色`} value={entry.value.color ?? ''} placeholder="自动颜色" onChange={event => update(index, item => ({ ...item, value: { ...item.value, color: event.target.value || undefined } }))} />
        </Inline>
        {chartType === 'scatter' || chartType === 'bubble' ? <Stack gap="xs">{(['xRange', 'yRange', ...(chartType === 'bubble' ? ['sizeRange' as const] : [])] as const).map(key => <TextInput key={key} aria-label={`系列 ${index + 1} ${key}`} value={entry[key]} placeholder={key === 'xRange' ? 'X 值范围' : key === 'yRange' ? 'Y 值范围' : '气泡大小范围'} onChange={event => update(index, item => ({ ...item, [key]: event.target.value }))} />)}</Stack> : null}
        <CheckToggle label="显示数据标记" checked={entry.value.marker?.enabled === true} onChange={event => update(index, item => ({ ...item, value: { ...item.value, marker: { ...item.value.marker, enabled: event.currentTarget.checked } } }))} />
        <Inline gap="xs"><Button size="xs" variant="secondary" onClick={() => update(index, item => ({ ...item, value: { ...item.value, trendlines: [...(item.value.trendlines ?? []), { id: crypto.randomUUID(), type: 'linear', displayEquation: true, displayRSquared: true }] } }))}>添加线性趋势线</Button></Inline>
        {entry.value.trendlines?.length ? <Button size="xs" variant="ghost" onClick={() => update(index, item => ({ ...item, value: { ...item.value, trendlines: [] } }))}>清除趋势线（{entry.value.trendlines.length}）</Button> : null}
        <Box className="space-y-2 rounded-md border border-slate-200 bg-slate-50 p-2"><Text size="xs" weight="medium">误差线</Text>
          <Inline gap="xs"><Select aria-label={`系列 ${index + 1} 误差线类型`} value={entry.value.errorBars?.type ?? 'none'} onChange={event => updateErrorType(index, event.target.value as NonNullable<ChartSeriesDraft['value']['errorBars']>['type'] | 'none')}><option value="none">无</option><option value="fixed">固定值</option><option value="percentage">百分比</option><option value="standard-deviation">标准偏差</option><option value="standard-error">标准误差</option><option value="custom">自定义</option></Select>
            {entry.value.errorBars ? <Select aria-label={`系列 ${index + 1} 误差线方向`} value={entry.value.errorBars.direction ?? 'vertical'} onChange={event => update(index, item => ({ ...item, value: { ...item.value, errorBars: item.value.errorBars ? { ...item.value.errorBars, direction: event.target.value as 'vertical' | 'horizontal' | 'both' } : undefined } }))}><option value="vertical">垂直</option><option value="horizontal">水平</option><option value="both">双向</option></Select> : null}</Inline>
          {entry.value.errorBars?.type === 'fixed' || entry.value.errorBars?.type === 'percentage' ? <TextInput aria-label={`系列 ${index + 1} 误差线值`} type="number" min="0" value={entry.value.errorBars.value ?? ''} placeholder={entry.value.errorBars.type === 'percentage' ? '百分比' : '数值'} onChange={event => update(index, item => ({ ...item, value: { ...item.value, errorBars: item.value.errorBars ? { ...item.value.errorBars, value: event.target.value === '' ? undefined : Number(event.target.value) } : undefined } }))} /> : null}
          {entry.value.errorBars?.type === 'custom' ? <Stack gap="xs"><TextInput aria-label={`系列 ${index + 1} 自定义正误差范围`} value={entry.errorPlusRange} placeholder="正误差范围，例如 D2:D20" onChange={event => update(index, item => ({ ...item, errorPlusRange: event.target.value }))} /><TextInput aria-label={`系列 ${index + 1} 自定义负误差范围`} value={entry.errorMinusRange} placeholder="负误差范围，例如 E2:E20" onChange={event => update(index, item => ({ ...item, errorMinusRange: event.target.value }))} /></Stack> : null}
          {entry.value.errorBars ? <CheckToggle label="显示端帽" checked={entry.value.errorBars.endStyle !== 'no-cap'} onChange={event => update(index, item => ({ ...item, value: { ...item.value, errorBars: item.value.errorBars ? { ...item.value.errorBars, endStyle: event.currentTarget.checked ? 'cap' : 'no-cap' } : undefined } }))} /> : null}
        </Box>
      </Stack>
    </Box>)}
    <Button size="sm" variant="secondary" icon="plus" disabled={!canAdd} onClick={onAdd}>添加数据系列</Button>
  </Stack>;
}
