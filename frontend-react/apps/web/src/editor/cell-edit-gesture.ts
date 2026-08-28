import { canonicalKeyGesture, type CanonicalKeyGesture } from '@react-sheets/spreadsheet-app';

export interface KeyboardGestureSource {
  key: string;
  code: string;
  altKey: boolean;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  repeat: boolean;
  nativeEvent: { isComposing?: boolean };
}
/** DOM normalization only. Editing semantics remain in CellEditDomain. */
export function toCanonicalKeyGesture(event: KeyboardGestureSource): CanonicalKeyGesture {
  return canonicalKeyGesture({
    key: event.key,
    code: event.code,
    altKey: event.altKey,
    ctrlKey: event.ctrlKey,
    metaKey: event.metaKey,
    shiftKey: event.shiftKey,
    repeat: event.repeat,
    composing: Boolean(event.nativeEvent.isComposing),
  });
}
