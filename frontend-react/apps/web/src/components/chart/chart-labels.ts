import { CHART_SUBTYPES_BY_TYPE, type ChartDrawingPayload } from '@react-sheets/core-model';

export const chartLabels: Record<ChartDrawingPayload['chartType'], string> = {
  column: '柱形图', bar: '条形图', line: '折线图', area: '面积图', pie: '饼图', doughnut: '圆环图',
  scatter: '散点图', bubble: '气泡图', treemap: '矩形树图', sunburst: '旭日图', histogram: '直方图',
  pareto: '帕累托图', 'box-whisker': '箱线图', waterfall: '瀑布图', funnel: '漏斗图', stock: '股价图',
  surface: '曲面图', radar: '雷达图', map: '地图', combo: '组合图',
};
export const chartTypes = Object.keys(CHART_SUBTYPES_BY_TYPE) as ChartDrawingPayload['chartType'][];

export const chartSubtypeLabels: Record<ChartDrawingPayload['subtype'], string> = {
  clustered: '簇状', stacked: '堆积', 'percent-stacked': '百分比堆积',
  'three-dimensional': '三维', 'three-dimensional-stacked': '三维堆积', 'three-dimensional-percent-stacked': '三维百分比堆积',
  cone: '圆锥', 'cone-stacked': '堆积圆锥', 'cone-percent-stacked': '百分比圆锥',
  cylinder: '圆柱', 'cylinder-stacked': '堆积圆柱', 'cylinder-percent-stacked': '百分比圆柱',
  pyramid: '棱锥', 'pyramid-stacked': '堆积棱锥', 'pyramid-percent-stacked': '百分比棱锥',
  line: '折线', 'line-markers': '带数据标记', 'stacked-markers': '堆积带标记', 'percent-stacked-markers': '百分比堆积带标记',
  area: '面积', pie: '饼图', 'exploded-pie': '分离饼图', 'exploded-three-dimensional-pie': '三维分离饼图',
  'pie-of-pie': '复合饼图', 'bar-of-pie': '复合条饼图', doughnut: '圆环', 'exploded-doughnut': '分离圆环',
  'scatter-markers': '仅数据标记', 'scatter-smooth-lines-markers': '平滑线与标记', 'scatter-smooth-lines': '平滑线',
  'scatter-straight-lines-markers': '直线与标记', 'scatter-straight-lines': '直线',
  bubble: '气泡', 'bubble-three-dimensional': '三维气泡', treemap: '矩形树', sunburst: '旭日',
  histogram: '直方', pareto: '帕累托', 'box-whisker': '箱线', waterfall: '瀑布', funnel: '漏斗',
  'stock-high-low-close': '最高－最低－收盘', 'stock-open-high-low-close': '开盘－最高－最低－收盘',
  'stock-volume-high-low-close': '成交量－最高－最低－收盘', 'stock-volume-open-high-low-close': '成交量－开盘－最高－最低－收盘',
  'surface-three-dimensional': '三维曲面', 'surface-wireframe': '曲面线框', 'surface-contour': '等高线', 'surface-wireframe-contour': '等高线线框',
  radar: '雷达', 'radar-markers': '带标记雷达', 'radar-filled': '填充雷达',
  'clustered-column-line': '簇状柱形与折线', 'clustered-column-line-secondary': '柱形与次轴折线',
  'stacked-area-clustered-column': '堆积面积与柱形', 'custom-combo': '自定义组合',
  'filled-map': '填充地图', 'region-map': '区域地图',
};
