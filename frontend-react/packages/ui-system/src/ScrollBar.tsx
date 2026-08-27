import { useEffect, useRef, type KeyboardEvent, type PointerEvent } from 'react';
import { Box } from './layout';

export interface ScrollBarProps {
  orientation: 'horizontal' | 'vertical';
  viewportSize: number;
  contentSize: number;
  offset: number;
  onChange: (offset: number) => void;
}

export function ScrollBar({ orientation, viewportSize, contentSize, offset, onChange }: ScrollBarProps) {
  const dragging = useRef(false);
  const trackRef = useRef<HTMLDivElement | null>(null);
  const pendingOffset = useRef<number | null>(null);
  const frame = useRef<number | null>(null);
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  useEffect(() => () => {
    if (frame.current !== null && typeof window !== 'undefined') window.cancelAnimationFrame(frame.current);
  }, []);
  const maxOffset = Math.max(0, contentSize - viewportSize);
  if (maxOffset <= 0 || viewportSize <= 0 || contentSize <= 0) return null;
  const viewportRatio = Math.min(1, viewportSize / contentSize);
  const thumbRatio = Math.max(0.08, viewportRatio);
  const flushPendingOffset = () => {
    if (frame.current !== null && typeof window !== 'undefined') window.cancelAnimationFrame(frame.current);
    frame.current = null;
    const nextOffset = pendingOffset.current;
    pendingOffset.current = null;
    if (nextOffset !== null) onChangeRef.current(nextOffset);
  };
  const scheduleOffset = (nextOffset: number) => {
    pendingOffset.current = nextOffset;
    if (frame.current !== null) return;
    if (typeof window === 'undefined' || typeof window.requestAnimationFrame !== 'function') {
      flushPendingOffset();
      return;
    }
    frame.current = window.requestAnimationFrame(() => flushPendingOffset());
  };
  const updateFromPoint = (client: number) => {
    const track = trackRef.current;
    if (!track) return;
    const rect = track.getBoundingClientRect();
    const length = orientation === 'horizontal' ? rect.width : rect.height;
    const position = orientation === 'horizontal' ? client - rect.left : client - rect.top;
    const thumbLength = length * thumbRatio;
    const usable = Math.max(1, length - thumbLength);
    scheduleOffset(Math.max(0, Math.min(maxOffset, ((position - thumbLength / 2) / usable) * maxOffset)));
  };
  const handlePointerDown = (event: PointerEvent<HTMLDivElement>) => {
    dragging.current = true;
    event.currentTarget.setPointerCapture(event.pointerId);
    updateFromPoint(orientation === 'horizontal' ? event.clientX : event.clientY);
  };
  const handlePointerMove = (event: PointerEvent<HTMLDivElement>) => {
    if (dragging.current) updateFromPoint(orientation === 'horizontal' ? event.clientX : event.clientY);
  };
  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const page = viewportSize * 0.9;
    if (orientation === 'horizontal') {
      if (event.key === 'ArrowLeft') scheduleOffset(offset - 40);
      if (event.key === 'ArrowRight') scheduleOffset(offset + 40);
      if (event.key === 'PageUp') scheduleOffset(offset - page);
      if (event.key === 'PageDown') scheduleOffset(offset + page);
    } else {
      if (event.key === 'ArrowUp') scheduleOffset(offset - 40);
      if (event.key === 'ArrowDown') scheduleOffset(offset + 40);
      if (event.key === 'PageUp') scheduleOffset(offset - page);
      if (event.key === 'PageDown') scheduleOffset(offset + page);
    }
  };
  const thumbStyle = orientation === 'horizontal'
    ? { width: `${thumbRatio * 100}%`, left: `${(offset / maxOffset) * (100 - thumbRatio * 100)}%`, top: 1, bottom: 1 }
    : { height: `${thumbRatio * 100}%`, top: `${(offset / maxOffset) * (100 - thumbRatio * 100)}%`, left: 2, right: 2 };
  return (
    <Box
      aria-label={`${orientation} scrollbar`}
      aria-valuemax={maxOffset}
      aria-valuemin={0}
      aria-valuenow={offset}
      className={orientation === 'horizontal' ? 'absolute bottom-0 left-0 right-[15px] z-30 h-[10px] border-t border-slate-300 bg-[#f2f2f2]' : 'absolute bottom-[10px] right-0 top-0 z-30 w-[15px] border-l border-slate-300 bg-[#f2f2f2]'}
      onKeyDown={handleKeyDown}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={() => { dragging.current = false; flushPendingOffset(); }}
      onPointerCancel={() => { dragging.current = false; flushPendingOffset(); }}
      ref={trackRef}
      role="scrollbar"
      tabIndex={0}
    >
      <Box className="absolute rounded-sm border border-slate-400 bg-slate-300 hover:bg-slate-400" style={thumbStyle} />
    </Box>
  );
}
