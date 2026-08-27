import { Fragment, useEffect, useRef, type PointerEvent as ReactPointerEvent } from 'react';
import { Box } from '@react-sheets/ui-system';
import type { CanvasRenderEngine } from '@react-sheets/render-engine';
import { cellAddress, type CanvasSheetSnapshot, type CellEditController, type CellEditSession, type FormulaReferenceSelection } from '@react-sheets/spreadsheet-app';

const REFERENCE_COLORS = ['#2563eb', '#dc2626', '#16a34a', '#9333ea', '#ea580c', '#0891b2', '#c026d3', '#65a30d'] as const;

interface ReferenceDrag {
  pointerId: number;
  mode: 'move' | 'resize';
  reference: FormulaReferenceSelection;
  startCell: { row: number; column: number };
  element: HTMLElement;
}

export interface FormulaReferenceOverlayProps {
  cellEdit: CellEditController;
  engine: CanvasRenderEngine;
  host: HTMLElement;
  session: CellEditSession;
  sheet: CanvasSheetSnapshot;
}

function normalizeRange(startRow: number, startColumn: number, endRow: number, endColumn: number, sheetId: string) {
  return {
    sheetId,
    startRow: Math.min(startRow, endRow),
    endRow: Math.max(startRow, endRow),
    startColumn: Math.min(startColumn, endColumn),
    endColumn: Math.max(startColumn, endColumn),
  };
}

function referenceText(sheet: CanvasSheetSnapshot, session: CellEditSession, selection: FormulaReferenceSelection): string {
  const start = cellAddress(selection.range.startRow, selection.range.startColumn);
  const end = cellAddress(selection.range.endRow, selection.range.endColumn);
  const address = start === end ? start : `${start}:${end}`;
  if (sheet.id === session.target.display.sheetId) return address;
  return `'${sheet.name.replace(/'/g, "''")}'!${address}`;
}

export function FormulaReferenceOverlay({ cellEdit, engine, host, session, sheet }: FormulaReferenceOverlayProps) {
  const dragRef = useRef<ReferenceDrag | null>(null);
  const pendingCellRef = useRef<{ row: number; column: number } | null>(null);
  const frameRef = useRef<number | null>(null);

  useEffect(() => () => {
    if (frameRef.current !== null) cancelAnimationFrame(frameRef.current);
    const drag = dragRef.current;
    if (drag?.element.hasPointerCapture(drag.pointerId)) drag.element.releasePointerCapture(drag.pointerId);
  }, []);

  const applyPendingDrag = () => {
    frameRef.current = null;
    const drag = dragRef.current;
    const current = pendingCellRef.current;
    if (!drag || !current) return;
    const start = drag.reference.range;
    let range;
    if (drag.mode === 'resize') {
      range = normalizeRange(start.startRow, start.startColumn, current.row, current.column, sheet.id);
    } else {
      const rowSpan = start.endRow - start.startRow;
      const columnSpan = start.endColumn - start.startColumn;
      const requestedRow = start.startRow + current.row - drag.startCell.row;
      const requestedColumn = start.startColumn + current.column - drag.startCell.column;
      const nextStartRow = Math.max(0, Math.min(sheet.rowCount - rowSpan - 1, requestedRow));
      const nextStartColumn = Math.max(0, Math.min(sheet.columnCount - columnSpan - 1, requestedColumn));
      range = { ...start, startRow: nextStartRow, endRow: nextStartRow + rowSpan, startColumn: nextStartColumn, endColumn: nextStartColumn + columnSpan };
    }
    const liveSession = cellEdit.getSnapshot().session;
    const liveReference = liveSession?.referenceSelections.find((reference) => reference.id === drag.reference.id) ?? drag.reference;
    const selection: FormulaReferenceSelection = { ...liveReference, range, operation: drag.mode };
    cellEdit.dispatch({ type: 'caret.set', caret: liveReference.tokenSpan });
    cellEdit.dispatch({ type: 'reference.insert', referenceText: referenceText(sheet, liveSession ?? session, selection), selection });
  };

  const beginDrag = (event: ReactPointerEvent<HTMLElement>, reference: FormulaReferenceSelection, mode: ReferenceDrag['mode']) => {
    const bounds = host.getBoundingClientRect();
    const startCell = engine.cellAtLocalPoint({ x: event.clientX - bounds.left, y: event.clientY - bounds.top });
    if (!startCell) return;
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.setPointerCapture(event.pointerId);
    cellEdit.dispatch({ type: 'reference.gesture.begin' });
    cellEdit.dispatch({ type: 'reference.set', selections: session.referenceSelections, activeReferenceId: reference.id });
    dragRef.current = { pointerId: event.pointerId, mode, reference: structuredClone(reference), startCell, element: event.currentTarget };
  };

  const moveDrag = (event: ReactPointerEvent<HTMLElement>) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    const bounds = host.getBoundingClientRect();
    pendingCellRef.current = engine.cellAtLocalPoint({ x: event.clientX - bounds.left, y: event.clientY - bounds.top });
    if (!pendingCellRef.current || frameRef.current !== null) return;
    frameRef.current = requestAnimationFrame(applyPendingDrag);
  };

  const endDrag = (event: ReactPointerEvent<HTMLElement>) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (frameRef.current !== null) {
      cancelAnimationFrame(frameRef.current);
      applyPendingDrag();
    }
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
    dragRef.current = null;
    pendingCellRef.current = null;
    cellEdit.dispatch({ type: 'reference.gesture.end' });
  };

  const cancelDrag = (event: ReactPointerEvent<HTMLElement>) => {
    const drag = dragRef.current;
    if (!drag || drag.pointerId !== event.pointerId) return;
    if (frameRef.current !== null) cancelAnimationFrame(frameRef.current);
    frameRef.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
    dragRef.current = null;
    pendingCellRef.current = null;
    cellEdit.dispatch({ type: 'reference.gesture.cancel' });
  };

  return (
    <>
      {session.referenceSelections.filter((reference) => reference.sheetId === sheet.id).flatMap((reference) => {
        const color = REFERENCE_COLORS[reference.colorIndex % REFERENCE_COLORS.length]!;
        return engine.contentRangeToScreenRects(reference.range).map((rect, rectIndex) => (
          <Fragment key={`${reference.id}:${rectIndex}:${rect.x}:${rect.y}`}>
            <Box className="pointer-events-none absolute z-[18] border-2" style={{ left: rect.x, top: rect.y, width: rect.width, height: rect.height, borderColor: color }} />
            <Box data-pointer-gesture-owner="cell-editor" className="absolute z-[19] h-2 cursor-move" style={{ left: rect.x, top: rect.y - 3, width: rect.width }} onPointerDown={(event) => beginDrag(event, reference, 'move')} onPointerMove={moveDrag} onPointerUp={endDrag} onPointerCancel={cancelDrag} />
            <Box data-pointer-gesture-owner="cell-editor" className="absolute z-[19] w-2 cursor-move" style={{ left: rect.x - 3, top: rect.y, height: rect.height }} onPointerDown={(event) => beginDrag(event, reference, 'move')} onPointerMove={moveDrag} onPointerUp={endDrag} onPointerCancel={cancelDrag} />
            <Box data-pointer-gesture-owner="cell-editor" className="absolute z-[19] h-2 cursor-move" style={{ left: rect.x, top: rect.y + rect.height - 4, width: rect.width }} onPointerDown={(event) => beginDrag(event, reference, 'move')} onPointerMove={moveDrag} onPointerUp={endDrag} onPointerCancel={cancelDrag} />
            <Box data-pointer-gesture-owner="cell-editor" className="absolute z-[19] w-2 cursor-move" style={{ left: rect.x + rect.width - 4, top: rect.y, height: rect.height }} onPointerDown={(event) => beginDrag(event, reference, 'move')} onPointerMove={moveDrag} onPointerUp={endDrag} onPointerCancel={cancelDrag} />
            <Box data-pointer-gesture-owner="cell-editor" className="absolute z-[20] h-2.5 w-2.5 cursor-se-resize border border-white" style={{ left: rect.x + rect.width - 5, top: rect.y + rect.height - 5, backgroundColor: color }} onPointerDown={(event) => beginDrag(event, reference, 'resize')} onPointerMove={moveDrag} onPointerUp={endDrag} onPointerCancel={cancelDrag} />
          </Fragment>
        ));
      })}
    </>
  );
}
