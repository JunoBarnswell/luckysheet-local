import type { ChartModel } from '@react-sheets/core-model';

export interface ChartDataSeries {
  name: string;
  values: number[];
  color: string;
}

export interface ChartRenderOptions {
  context: CanvasRenderingContext2D;
  chart: ChartModel;
  categories: string[];
  series: ChartDataSeries[];
}

const DEFAULT_CHART_PALETTE = [
  '#2563eb', // blue
  '#10b981', // emerald
  '#f59e0b', // amber
  '#ef4444', // red
  '#8b5cf6', // purple
  '#ec4899', // pink
  '#06b6d4', // cyan
];

export function drawChartOnCanvas(options: ChartRenderOptions): void {
  const { context, chart, categories, series } = options;
  const { x, y, width, height } = chart.bounds;

  context.save();
  context.translate(x, y);

  // Background and border card
  context.fillStyle = '#ffffff';
  context.fillRect(0, 0, width, height);
  context.strokeStyle = '#e2e8f0';
  context.lineWidth = 1;
  context.strokeRect(0, 0, width, height);

  // Chart Title
  if (chart.title) {
    context.fillStyle = '#1e293b';
    context.font = 'bold 14px Inter, sans-serif';
    context.textAlign = 'left';
    context.textBaseline = 'top';
    context.fillText(chart.title, 16, 12);
  }

  const plotTop = chart.title ? 40 : 20;
  const plotBottom = height - (chart.legendPosition === 'bottom' ? 40 : 24);
  const plotLeft = 48;
  const plotRight = width - (chart.legendPosition === 'right' ? 80 : 16);
  const plotWidth = Math.max(10, plotRight - plotLeft);
  const plotHeight = Math.max(10, plotBottom - plotTop);

  // Determine min and max
  let maxVal = 0;
  let minVal = 0;
  for (const s of series) {
    for (const v of s.values) {
      if (v > maxVal) maxVal = v;
      if (v < minVal) minVal = v;
    }
  }
  if (maxVal === 0 && minVal === 0) maxVal = 100;
  maxVal = Math.ceil(maxVal * 1.1);

  // Draw Grid lines
  const gridLines = 4;
  context.strokeStyle = '#f1f5f9';
  context.lineWidth = 1;
  context.font = '11px Inter, sans-serif';
  context.fillStyle = '#64748b';
  context.textAlign = 'right';
  context.textBaseline = 'middle';

  for (let i = 0; i <= gridLines; i++) {
    const gy = plotBottom - (i / gridLines) * plotHeight;
    const gVal = Math.round(minVal + (i / gridLines) * (maxVal - minVal));
    context.beginPath();
    context.moveTo(plotLeft, gy);
    context.lineTo(plotRight, gy);
    context.stroke();
    context.fillText(String(gVal), plotLeft - 8, gy);
  }

  // Draw chart series by type
  if (chart.type === 'column' || chart.type === 'bar') {
    drawColumnChart(context, categories, series, plotLeft, plotTop, plotWidth, plotHeight, maxVal);
  } else if (chart.type === 'line' || chart.type === 'area') {
    drawLineChart(context, categories, series, plotLeft, plotTop, plotWidth, plotHeight, maxVal, chart.type === 'area');
  } else if (chart.type === 'pie' || chart.type === 'doughnut') {
    drawPieChart(context, categories, series, plotLeft, plotTop, plotWidth, plotHeight, chart.type === 'doughnut');
  }

  // Draw Legend
  drawLegend(context, series, width, height, chart.legendPosition ?? 'top');

  context.restore();
}

function drawColumnChart(
  ctx: CanvasRenderingContext2D,
  categories: string[],
  series: ChartDataSeries[],
  left: number,
  top: number,
  width: number,
  height: number,
  maxVal: number,
): void {
  const catCount = Math.max(1, categories.length);
  const catWidth = width / catCount;
  const seriesCount = Math.max(1, series.length);
  const barWidth = Math.max(4, (catWidth * 0.7) / seriesCount);

  for (let c = 0; c < categories.length; c++) {
    const cx = left + c * catWidth + catWidth * 0.15;
    for (let s = 0; s < series.length; s++) {
      const val = series[s]?.values[c] ?? 0;
      const barH = (val / maxVal) * height;
      const bx = cx + s * barWidth;
      const by = top + height - barH;

      ctx.fillStyle = series[s]?.color ?? DEFAULT_CHART_PALETTE[s % DEFAULT_CHART_PALETTE.length]!;
      ctx.fillRect(bx, by, barWidth - 2, barH);
    }

    // Category Label
    ctx.fillStyle = '#64748b';
    ctx.font = '11px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(categories[c] ?? '', left + c * catWidth + catWidth / 2, top + height + 6);
  }
}

function drawLineChart(
  ctx: CanvasRenderingContext2D,
  categories: string[],
  series: ChartDataSeries[],
  left: number,
  top: number,
  width: number,
  height: number,
  maxVal: number,
  isArea: boolean,
): void {
  const catCount = Math.max(1, categories.length);
  const step = width / (catCount - 1 || 1);

  for (let s = 0; s < series.length; s++) {
    const currentSeries = series[s]!;
    const color = currentSeries.color ?? DEFAULT_CHART_PALETTE[s % DEFAULT_CHART_PALETTE.length]!;
    const points: Array<{ x: number; y: number }> = [];

    for (let c = 0; c < categories.length; c++) {
      const val = currentSeries.values[c] ?? 0;
      const px = left + c * step;
      const py = top + height - (val / maxVal) * height;
      points.push({ x: px, y: py });
    }

    if (isArea && points.length > 0) {
      ctx.save();
      ctx.fillStyle = color + '22'; // low opacity fill
      ctx.beginPath();
      ctx.moveTo(points[0]!.x, top + height);
      for (const p of points) ctx.lineTo(p.x, p.y);
      ctx.lineTo(points[points.length - 1]!.x, top + height);
      ctx.closePath();
      ctx.fill();
      ctx.restore();
    }

    // Draw line
    ctx.strokeStyle = color;
    ctx.lineWidth = 2.5;
    ctx.beginPath();
    for (let i = 0; i < points.length; i++) {
      if (i === 0) ctx.moveTo(points[i]!.x, points[i]!.y);
      else ctx.lineTo(points[i]!.x, points[i]!.y);
    }
    ctx.stroke();

    // Draw point markers
    ctx.fillStyle = '#ffffff';
    ctx.strokeStyle = color;
    ctx.lineWidth = 2;
    for (const p of points) {
      ctx.beginPath();
      ctx.arc(p.x, p.y, 4, 0, Math.PI * 2);
      ctx.fill();
      ctx.stroke();
    }
  }

  // Draw category labels
  for (let c = 0; c < categories.length; c++) {
    ctx.fillStyle = '#64748b';
    ctx.font = '11px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'top';
    ctx.fillText(categories[c] ?? '', left + c * step, top + height + 6);
  }
}

function drawPieChart(
  ctx: CanvasRenderingContext2D,
  categories: string[],
  series: ChartDataSeries[],
  left: number,
  top: number,
  width: number,
  height: number,
  isDoughnut: boolean,
): void {
  const firstSeries = series[0];
  if (!firstSeries) return;

  const total = firstSeries.values.reduce((acc, v) => acc + Math.max(0, v), 0);
  if (total <= 0) return;

  const centerX = left + width / 2;
  const centerY = top + height / 2;
  const radius = Math.min(width, height) / 2 - 10;
  let currentAngle = -Math.PI / 2;

  for (let i = 0; i < firstSeries.values.length; i++) {
    const val = Math.max(0, firstSeries.values[i] ?? 0);
    const sliceAngle = (val / total) * Math.PI * 2;
    const color = DEFAULT_CHART_PALETTE[i % DEFAULT_CHART_PALETTE.length]!;

    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, currentAngle, currentAngle + sliceAngle);
    ctx.closePath();
    ctx.fill();

    currentAngle += sliceAngle;
  }

  if (isDoughnut) {
    ctx.fillStyle = '#ffffff';
    ctx.beginPath();
    ctx.arc(centerX, centerY, radius * 0.55, 0, Math.PI * 2);
    ctx.fill();
  }
}

function drawLegend(
  ctx: CanvasRenderingContext2D,
  series: ChartDataSeries[],
  width: number,
  height: number,
  position: 'top' | 'bottom' | 'left' | 'right' | 'none',
): void {
  if (position === 'none') return;
  ctx.font = '11px Inter, sans-serif';
  ctx.textAlign = 'left';
  ctx.textBaseline = 'middle';

  let lx = 16;
  const ly = position === 'bottom' ? height - 16 : 24;

  for (let s = 0; s < series.length; s++) {
    const currentSeries = series[s]!;
    const color = currentSeries.color ?? DEFAULT_CHART_PALETTE[s % DEFAULT_CHART_PALETTE.length]!;

    ctx.fillStyle = color;
    ctx.fillRect(lx, ly - 4, 10, 10);

    ctx.fillStyle = '#475569';
    ctx.fillText(currentSeries.name, lx + 14, ly + 1);

    lx += ctx.measureText(currentSeries.name).width + 32;
    if (lx > width - 40) break;
  }
}
