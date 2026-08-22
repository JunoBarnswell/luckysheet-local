import { createElement, forwardRef, useImperativeHandle, useLayoutEffect, useRef } from 'react';
import type { CSSProperties } from 'react';
import { CanvasRenderEngine, type CanvasRenderEngineOptions } from './canvas-render-engine';

export interface CanvasRenderSurfaceProps {
  engine?: CanvasRenderEngine;
  options?: CanvasRenderEngineOptions;
  className?: string;
  style?: CSSProperties;
  onReady?: (engine: CanvasRenderEngine) => void;
}

export const CanvasRenderSurface = forwardRef<CanvasRenderEngine, CanvasRenderSurfaceProps>(
  function CanvasRenderSurface(props, ref) {
    const engineRef = useRef<CanvasRenderEngine | null>(null);
    const ownsEngineRef = useRef(false);
    if (!engineRef.current) {
      engineRef.current = props.engine ?? new CanvasRenderEngine(props.options);
      ownsEngineRef.current = props.engine === undefined;
    }
    const engine = engineRef.current;
    const hostRef = useRef<HTMLDivElement | null>(null);

    useImperativeHandle(ref, () => engine, [engine]);
    useLayoutEffect(() => {
      const host = hostRef.current;
      if (!host) return undefined;
      engine.mount(host);
      props.onReady?.(engine);
      const resize = () => {
        engine.resizeFromHost();
        engine.requestRender();
      };
      const observer = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(resize) : null;
      observer?.observe(host);
      const frame = typeof requestAnimationFrame === 'function' ? requestAnimationFrame(resize) : null;
      resize();
      return () => {
        observer?.disconnect();
        if (frame !== null) cancelAnimationFrame(frame);
        engine.unmount();
        if (ownsEngineRef.current) engine.dispose();
      };
    }, [engine, props.onReady]);

    return createElement('div', {
      ref: hostRef,
      className: props.className,
      style: {
        position: 'absolute',
        inset: 0,
        overflow: 'hidden',
        width: '100%',
        height: '100%',
        ...props.style,
      },
      'data-render-surface': 'canvas',
    });
  },
);
