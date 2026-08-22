import React, { useState } from 'react';
import { cn } from './cn';

export interface ColorPickerProps {
  color?: string;
  onChange: (color: string) => void;
  onClose?: () => void;
}

const PALETTE_COLORS = [
  ['#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff'],
  ['#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff'],
  ['#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc'],
  ['#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd'],
  ['#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0'],
  ['#a61c1c', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79'],
  ['#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1155cc', '#073763', '#20124d', '#4c1130'],
];

export function ColorPicker({ color, onChange, onClose }: ColorPickerProps) {
  const [customColor, setCustomColor] = useState(color || '#000000');

  return (
    <div className="flex flex-col gap-3 rounded-lg border border-slate-200 bg-white p-3 shadow-xl">
      <div className="text-xs font-semibold text-slate-600">Standard Palette</div>
      <div className="flex flex-col gap-1">
        {PALETTE_COLORS.map((row, rIdx) => (
          <div key={`row-${rIdx}`} className="flex gap-1">
            {row.map((c) => (
              <button
                key={c}
                type="button"
                onClick={() => {
                  onChange(c);
                  onClose?.();
                }}
                className={cn(
                  'h-4 w-4 rounded-xs border border-slate-300 transition-transform hover:scale-125 hover:z-10',
                  color === c && 'ring-2 ring-blue-500 ring-offset-1',
                )}
                style={{ backgroundColor: c }}
                title={c}
              />
            ))}
          </div>
        ))}
      </div>

      <div className="border-t border-slate-100 pt-2">
        <div className="mb-1 text-xs font-semibold text-slate-600">Custom Color</div>
        <div className="flex items-center gap-2">
          <input
            type="color"
            value={customColor}
            onChange={(e) => setCustomColor(e.target.value)}
            className="h-7 w-7 cursor-pointer rounded border border-slate-300 p-0"
          />
          <input
            type="text"
            value={customColor}
            onChange={(e) => setCustomColor(e.target.value)}
            className="h-7 flex-1 rounded border border-slate-200 px-2 text-xs font-mono"
            placeholder="#000000"
          />
          <button
            type="button"
            onClick={() => {
              onChange(customColor);
              onClose?.();
            }}
            className="rounded bg-blue-600 px-2.5 py-1 text-xs font-medium text-white hover:bg-blue-700"
          >
            Apply
          </button>
        </div>
      </div>
    </div>
  );
}
