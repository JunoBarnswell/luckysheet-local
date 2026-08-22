import type { SparklineModel } from '@react-sheets/core-model';

export interface SparklineRenderOptions {
  context: CanvasRenderingContext2D;
  sparkline: SparklineModel;
  values: number[];
  rect: { x: number; y: number; width: number; height: number };
}

export function drawSparklineOnCanvas(options: SparklineRenderOptions): void {
  const { context, sparkline, values, rect } = options;
  if (values.length === 0) return;

  context.save();
  const padding = 4;
  const drawX = rect.x + padding;
  const drawY = rect.y + padding;
  const drawWidth = Math.max(2, rect.width - padding * 2);
  const drawHeight = Math.max(2, rect.height - padding * 2);

  const maxVal = Math.max(...values);
  const minVal = Math.min(...values);
  const range = maxVal === minVal ? 1 : maxVal - minVal;

  if (sparkline.type === 'line') {
    const step = drawWidth / (values.length - 1 || 1);
    const points: Array<{ x: number; y: number }> = values.map((val, i) => ({
      x: drawX + i * step,
      y: drawY + drawHeight - ((val - minVal) / range) * drawHeight,
    }));

    context.strokeStyle = sparkline.color || '#2563eb';
    context.lineWidth = 1.5;
    context.beginPath();
    for (let i = 0; i < points.length; i++) {
      if (i === 0) context.moveTo(points[i]!.x, points[i]!.y);
      else context.lineTo(points[i]!.x, points[i]!.y);
    }
    context.stroke();

    // Min / Max markers
    if (sparkline.highlightMax) {
      const maxIdx = values.indexOf(maxVal);
      const p = points[maxIdx];
      if (p) {
        context.fillStyle = '#10b981'; // Green for max
        context.beginPath();
        context.arc(p.x, p.y, 2.5, 0, Math.PI * 2);
        context.fill();
      }
    }
    if (sparkline.highlightMin) {
      const minIdx = values.indexOf(minVal);
      const p = points[minIdx];
      if (p) {
        context.fillStyle = '#ef4444'; // Red for min
        context.beginPath();
        context.arc(p.x, p.y, 2.5, 0, Math.PI * 2);
        context.fill();
      }
    }
  } else if (sparkline.type === 'column') {
    const barWidth = Math.max(2, drawWidth / values.length - 2);
    const step = drawWidth / values.length;

    for (let i = 0; i < values.length; i++) {
      const val = values[i]!;
      const barH = Math.max(2, (Math.abs(val) / (Math.max(Math.abs(maxVal), Math.abs(minVal)) || 1)) * drawHeight);
      const bx = drawX + i * step + 1;
      const isNegative = val < 0;
      const by = isNegative ? drawY + drawHeight / 2 : drawY + drawHeight / 2 - barH;

      context.fillStyle = isNegative ? sparkline.negativeColor || '#ef4444' : sparkline.color || '#2563eb';
      context.fillRect(bx, by, barWidth, barH);
    }
  } else if (sparkline.type === 'win-loss') {
    const barWidth = Math.max(2, drawWidth / values.length - 2);
    const step = drawWidth / values.length;
    const halfH = drawHeight / 2;

    for (let i = 0; i < values.length; i++) {
      const val = values[i]!;
      const bx = drawX + i * step + 1;
      if (val > 0) {
        context.fillStyle = sparkline.color || '#2563eb';
        context.fillRect(bx, drawY, barWidth, halfH - 1);
      } else if (val < 0) {
        context.fillStyle = sparkline.negativeColor || '#ef4444';
        context.fillRect(bx, drawY + halfH + 1, barWidth, halfH - 1);
      } else {
        context.fillStyle = '#94a3b8';
        context.fillRect(bx, drawY + halfH - 1, barWidth, 2);
      }
    }
  }

  context.restore();
}
