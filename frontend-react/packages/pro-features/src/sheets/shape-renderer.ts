import type { ShapeModel } from '@react-sheets/core-model';

export interface ShapeRenderOptions {
  context: CanvasRenderingContext2D;
  shape: ShapeModel;
}

export function drawShapeOnCanvas(options: ShapeRenderOptions): void {
  const { context, shape } = options;
  const { x, y, width, height } = shape.bounds;

  context.save();
  context.translate(x + width / 2, y + height / 2);
  if (shape.rotation) {
    context.rotate((shape.rotation * Math.PI) / 180);
  }
  context.translate(-width / 2, -height / 2);

  context.fillStyle = shape.fill;
  context.strokeStyle = shape.stroke;
  context.lineWidth = shape.strokeWidth ?? 1.5;

  if (shape.type === 'rectangle') {
    context.fillRect(0, 0, width, height);
    context.strokeRect(0, 0, width, height);
  } else if (shape.type === 'rounded-rectangle') {
    const r = Math.min(8, width / 4, height / 4);
    context.beginPath();
    context.roundRect(0, 0, width, height, r);
    context.fill();
    context.stroke();
  } else if (shape.type === 'ellipse') {
    context.beginPath();
    context.ellipse(width / 2, height / 2, width / 2, height / 2, 0, 0, Math.PI * 2);
    context.fill();
    context.stroke();
  } else if (shape.type === 'line') {
    context.beginPath();
    context.moveTo(0, height / 2);
    context.lineTo(width, height / 2);
    context.stroke();
  } else if (shape.type === 'arrow') {
    const headLen = Math.min(16, width / 3);
    context.beginPath();
    context.moveTo(0, height / 2);
    context.lineTo(width - headLen, height / 2);
    context.stroke();

    context.fillStyle = shape.stroke;
    context.beginPath();
    context.moveTo(width, height / 2);
    context.lineTo(width - headLen, height / 2 - headLen / 2);
    context.lineTo(width - headLen, height / 2 + headLen / 2);
    context.closePath();
    context.fill();
  } else if (shape.type === 'star') {
    drawStar(context, width / 2, height / 2, 5, width / 2, width / 4);
    context.fill();
    context.stroke();
  } else if (shape.type === 'callout') {
    const r = 6;
    const bodyH = height - 12;
    context.beginPath();
    context.moveTo(r, 0);
    context.lineTo(width - r, 0);
    context.arcTo(width, 0, width, r, r);
    context.lineTo(width, bodyH - r);
    context.arcTo(width, bodyH, width - r, bodyH, r);
    context.lineTo(24, bodyH);
    context.lineTo(12, height);
    context.lineTo(16, bodyH);
    context.lineTo(r, bodyH);
    context.arcTo(0, bodyH, 0, bodyH - r, r);
    context.lineTo(0, r);
    context.arcTo(0, 0, r, 0, r);
    context.closePath();
    context.fill();
    context.stroke();
  }

  // Draw text inside shape if exists
  if (shape.text) {
    context.fillStyle = shape.textColor || '#1e293b';
    context.font = `${shape.fontSize || 13}px Inter, sans-serif`;
    context.textAlign = 'center';
    context.textBaseline = 'middle';
    context.fillText(shape.text, width / 2, height / 2, Math.max(10, width - 8));
  }

  context.restore();
}

function drawStar(
  ctx: CanvasRenderingContext2D,
  cx: number,
  cy: number,
  spikes: number,
  outerRadius: number,
  innerRadius: number,
): void {
  let rot = (Math.PI / 2) * 3;
  let x = cx;
  let y = cy;
  const step = Math.PI / spikes;

  ctx.beginPath();
  ctx.moveTo(cx, cy - outerRadius);
  for (let i = 0; i < spikes; i++) {
    x = cx + Math.cos(rot) * outerRadius;
    y = cy + Math.sin(rot) * outerRadius;
    ctx.lineTo(x, y);
    rot += step;

    x = cx + Math.cos(rot) * innerRadius;
    y = cy + Math.sin(rot) * innerRadius;
    ctx.lineTo(x, y);
    rot += step;
  }
  ctx.lineTo(cx, cy - outerRadius);
  ctx.closePath();
}
