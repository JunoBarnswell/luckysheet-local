export type PhoneticType = 'fullwidth-katakana' | 'halfwidth-katakana' | 'hiragana' | 'no-conversion';
export type PhoneticAlignment = 'left' | 'center' | 'distributed' | 'no-control';

export interface PhoneticRun {
  text: string;
  start: number;
  end: number;
}

export interface CellPhoneticMetadata {
  visible: boolean;
  type: PhoneticType;
  alignment: PhoneticAlignment;
  fontFamily?: string;
  fontSizePx?: number;
  runs: PhoneticRun[];
}

export function isCellPhoneticMetadata(value: unknown): value is CellPhoneticMetadata {
  if (!value || typeof value !== 'object') return false;
  const input = value as Record<string, unknown>;
  if (typeof input.visible !== 'boolean' || !['fullwidth-katakana', 'halfwidth-katakana', 'hiragana', 'no-conversion'].includes(String(input.type)) || !['left', 'center', 'distributed', 'no-control'].includes(String(input.alignment)) || !Array.isArray(input.runs)) return false;
  return input.runs.every((run) => {
    if (!run || typeof run !== 'object') return false;
    const entry = run as Record<string, unknown>;
    return typeof entry.text === 'string' && entry.text.length > 0 && Number.isSafeInteger(entry.start) && Number.isSafeInteger(entry.end) && Number(entry.start) >= 0 && Number(entry.end) > Number(entry.start);
  });
}
