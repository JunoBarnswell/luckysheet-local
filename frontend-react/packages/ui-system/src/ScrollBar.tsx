import { useEffect, useRef, type KeyboardEvent, type PointerEvent } from 'react';
import { Box } from './layout';
import {
  claimPointerGesture,
  ownsPointerGesture,
  releasePointerGesture,
  releasePointerGesturesForSurface,
  type PointerGestureOwner,
} from './pointer-gesture-owner';

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
  const activePointer = useRef<{ pointerId: number; owner: PointerGestureOwner } | null>(null);
  const pendingOffset = useRef<number | null>(null);
  const frame = useRef<number | null>(null);
  const onChangeRef = useRef(onChange);
  onChangeRef.current = onChange;
  useEffect(() => () => {
    if (frame.current !== null && typeof window !== 'undefined') window.cancelAnimationFrame(frame.current);
    const track = trackRef.current;
    if (track) releasePointerGesturesForSurface(track.ownerDocument, track);
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
    event.stopPropagation();
    if (activePointer.current) return;
    const owner: PointerGestureOwner = `scrollbar-${orientation}`;
    if (!claimPointerGesture(event.currentTarget.ownerDocument, event.pointerId, owner, event.currentTarget)) return;
    activePointer.current = { pointerId: event.pointerId, owner };
    dragging.current = true;
    event.currentTarget.setPointerCapture(event.pointerId);
    updateFromPoint(orientation === 'horizontal' ? event.clientX : event.clientY);
  };
  const handlePointerMove = (event: PointerEvent<HTMLDivElement>) => {
    event.stopPropagation();
    const active = activePointer.current;
    if (!active || active.pointerId !== event.pointerId) return;
    if (!ownsPointerGesture(event.currentTarget.ownerDocument, event.pointerId, active.owner, event.currentTarget)) return;
    if (dragging.current) updateFromPoint(orientation === 'horizontal' ? event.clientX : event.clientY);
  };
  const handleKeyDown = (event: KeyboardEvent<HTMLDivElement>) => {
    const page = viewportSize * 0.9;
    let handled = false;
    if (orientation === 'horizontal') {
      if (event.key === 'ArrowLeft') { scheduleOffset(offset - 40); handled = true; }
      if (event.key === 'ArrowRight') { scheduleOffset(offset + 40); handled = true; }
      if (event.key === 'PageUp') { scheduleOffset(offset - page); handled = true; }
      if (event.key === 'PageDown') { scheduleOffset(offset + page); handled = true; }
    } else {
      if (event.key === 'ArrowUp') { scheduleOffset(offset - 40); handled = true; }
      if (event.key === 'ArrowDown') { scheduleOffset(offset + 40); handled = true; }
      if (event.key === 'PageUp') { scheduleOffset(offset - page); handled = true; }
      if (event.key === 'PageDown') { scheduleOffset(offset + page); handled = true; }
    }
    if (!handled) return;
    event.preventDefault();
    event.stopPropagation();
  };
  const handlePointerEnd = (event: PointerEvent<HTMLDivElement>) => {
    event.stopPropagation();
    const active = activePointer.current;
    if (!active || active.pointerId !== event.pointerId) return;
    releasePointerGesture(event.currentTarget.ownerDocument, event.pointerId, active.owner, event.currentTarget);
    activePointer.current = null;
    dragging.current = false;
    flushPendingOffset();
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
      data-pointer-gesture-owner={`scrollbar-${orientation}`}
      onKeyDown={handleKeyDown}
      onPointerDown={handlePointerDown}
      onPointerMove={handlePointerMove}
      onPointerUp={handlePointerEnd}
      onPointerCancel={handlePointerEnd}
      onLostPointerCapture={handlePointerEnd}
      ref={trackRef}
      role="scrollbar"
      tabIndex={0}
    >
      <Box className="absolute rounded-sm border border-slate-400 bg-slate-300 hover:bg-slate-400" style={thumbStyle} />
    </Box>
  );
}
