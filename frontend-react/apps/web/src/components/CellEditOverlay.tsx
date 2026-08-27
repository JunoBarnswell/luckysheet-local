import { useMemo } from 'react';
import { Box, Button, Dialog, Inline, Text } from '@react-sheets/ui-system';
import type { CanvasRenderEngine } from '@react-sheets/render-engine';
import { useCellEdit, type CanvasSheetSnapshot, type CellEditController } from '@react-sheets/spreadsheet-app';
import { CellEditor } from './CellEditor';
import { FormulaReferenceOverlay } from './FormulaReferenceOverlay';

export interface CellEditOverlayProps {
  cellEdit: CellEditController;
  engine: CanvasRenderEngine | null;
  host: HTMLElement | null;
  scrollTick: number;
  sheet: CanvasSheetSnapshot;
}

interface EditorRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

function hasVisibleValue(sheet: CanvasSheetSnapshot, row: number, column: number): boolean {
  const cell = sheet.getCell(row, column);
  return Boolean(cell && (cell.formula || cell.value !== null && cell.value !== undefined && String(cell.displayValue ?? cell.value).length > 0));
}

function editorRect(
  engine: CanvasRenderEngine,
  sheet: CanvasSheetSnapshot,
  row: number,
  column: number,
  mergedRange: CellEditOverlayProps['sheet']['merges'][number]['range'] | undefined,
  draft: string,
): EditorRect | null {
  const range = mergedRange ?? { sheetId: sheet.id, startRow: row, endRow: row, startColumn: column, endColumn: column };
  const screen = engine.contentRangeToScreenRects(range)[0];
  if (!screen) return null;
  const style = sheet.getCell(row, column)?.style;
  const viewport = engine.viewport.getSnapshot();
  const fontSize = style?.fontSizePx ?? 13;
  const lines = draft.split('\n');
  const multiline = lines.length > 1 || style?.wrapText === true;
  let x = screen.x;
  let width = screen.width;
  let height = screen.height;

  if (multiline) {
    const estimatedLines = style?.wrapText
      ? Math.max(lines.length, Math.ceil(Math.max(...lines.map((line) => line.length), 1) * fontSize * 0.56 / Math.max(1, screen.width - 8)))
      : lines.length;
    height = Math.max(height, estimatedLines * Math.ceil(fontSize * 1.35) + 6);
  } else {
    const desiredWidth = Math.max(screen.width, draft.length * fontSize * 0.56 + 10);
    if (style?.horizontalAlignment === 'right') {
      let candidate = range.startColumn - 1;
      while (candidate >= 0 && width < desiredWidth && !hasVisibleValue(sheet, row, candidate)) {
        const nextWidth = engine.skeleton.getColumnWidth(candidate);
        width += nextWidth;
        x -= nextWidth;
        candidate -= 1;
      }
    } else {
      let candidate = range.endColumn + 1;
      while (candidate < sheet.columnCount && width < desiredWidth && !hasVisibleValue(sheet, row, candidate)) {
        width += engine.skeleton.getColumnWidth(candidate);
        candidate += 1;
      }
    }
  }

  x = Math.max(0, x);
  width = Math.max(screen.width, Math.min(width, viewport.width - x));
  height = Math.max(screen.height, Math.min(height, viewport.height - screen.y));
  return { x, y: screen.y, width, height };
}

export function CellEditOverlay({ cellEdit, engine, host, scrollTick, sheet }: CellEditOverlayProps) {
  const edit = useCellEdit(cellEdit);
  const session = edit.session;
  const displayTarget = session?.target.display;
  const mergedRange = session?.target.mergedRange;
  const draftText = session?.draft.text ?? '';
  const rect = useMemo(() => {
    void scrollTick;
    if (!engine || !displayTarget || displayTarget.sheetId !== sheet.id) return null;
    return editorRect(engine, sheet, displayTarget.row, displayTarget.column, mergedRange, draftText);
  }, [displayTarget, draftText, engine, mergedRange, scrollTick, sheet]);

  if (!session) return null;
  const cellStyle = session.target.display.sheetId === sheet.id ? sheet.getCell(session.target.display.row, session.target.display.column)?.style : undefined;
  const confirmation = session.overlay.kind === 'validation-confirmation' ? session.overlay : null;
  const editorList = session.overlay.kind === 'editor-list' ? session.overlay : null;
  const autocomplete = session.surface === 'grid' && session.overlay.kind === 'autocomplete' ? session.overlay : null;
  const functionHint = session.surface === 'grid' && session.overlay.kind === 'function-hint' ? session.overlay : null;
  const blockingError = session.validation.kind === 'blocking-error' ? session.validation : null;
  const inputMessage = session.overlay.kind === 'input-message' ? session.overlay : null;
  const popupBelow = !engine || !rect || rect.y + rect.height + 240 <= engine.viewport.height;
  const popupAnchorClass = popupBelow ? 'top-full mt-px' : 'bottom-full mb-px';

  return (
    <>
      {engine && host ? <FormulaReferenceOverlay cellEdit={cellEdit} engine={engine} host={host} session={session} sheet={sheet} /> : null}
      {rect ? (
        <Box
          className="absolute z-20 overflow-visible rounded-none border border-[#5292f7] bg-white"
          style={{ left: rect.x, top: rect.y, width: rect.width, height: rect.height }}
        >
          <CellEditor editorSurface={session.editorSurface} cellEdit={cellEdit} cellStyle={cellStyle} draft={session.draft} caret={session.caret} />
        {editorList ? (
          <Box className={`absolute left-0 z-30 max-h-56 min-w-full overflow-y-auto rounded border border-[#9ba8b6] bg-white py-1 shadow-lg ${popupAnchorClass}`}>
            {editorList.items.map((item, index) => (
              <Button
                key={`${index}:${item.text}`}
                size="sm"
                variant={index === editorList.activeIndex ? 'soft' : 'ghost'}
                className="w-full justify-start rounded-none"
                onPointerDown={(event) => event.preventDefault()}
                onClick={() => cellEdit.dispatch({ type: 'editor-list.accept', index })}
              >
                {item.label}
              </Button>
            ))}
          </Box>
        ) : null}
        {autocomplete ? (
          <Box className={`absolute left-0 z-30 max-h-64 min-w-72 overflow-y-auto rounded border border-[#9ba8b6] bg-white py-1 shadow-lg ${popupAnchorClass}`}>
            {autocomplete.candidates.map((candidate, index) => (
              <Button
                key={candidate.id}
                size="sm"
                variant={index === autocomplete.activeIndex ? 'soft' : 'ghost'}
                className="w-full justify-between rounded-none"
                onPointerDown={(event) => event.preventDefault()}
                onDoubleClick={() => cellEdit.dispatch({ type: 'autocomplete.accept' })}
                onMouseEnter={() => cellEdit.dispatch({ type: 'autocomplete.move', delta: index - autocomplete.activeIndex })}
              >
                <Text size="sm">{candidate.label}</Text>
                <Text size="xs" tone="subtle">{candidate.detail ?? candidate.kind}</Text>
              </Button>
            ))}
          </Box>
        ) : null}
        {functionHint ? (
          <Box className={`absolute left-0 z-30 rounded border border-[#9ba8b6] bg-amber-50 px-2 py-1 shadow-lg ${popupAnchorClass}`}>
            <Text size="xs"><Text size="xs" weight="semibold">{functionHint.functionName}</Text> · argument {functionHint.argumentIndex}</Text>
          </Box>
        ) : null}
          {blockingError ? (
          <Box className={`absolute left-0 z-30 min-w-64 max-w-96 rounded border border-rose-300 bg-white px-2 py-1 shadow-lg ${popupAnchorClass}`}>
            <Text size="xs" className="text-rose-700">{blockingError.message}</Text>
          </Box>
          ) : null}
          {inputMessage ? (
            <Box className={`absolute left-0 z-30 min-w-64 max-w-96 rounded border border-sky-200 bg-white px-2 py-1 shadow-lg ${popupAnchorClass}`}>
              {inputMessage.title ? <Text size="xs" weight="semibold">{inputMessage.title}</Text> : null}
              <Text size="xs" tone="muted">{inputMessage.message}</Text>
            </Box>
          ) : null}
        </Box>
      ) : null}
      <Dialog open={Boolean(confirmation)} title={confirmation?.title ?? '数据验证'} onClose={() => cellEdit.dispatch({ type: 'validation.reject' })} maxWidth="sm">
        <Text size="sm">{confirmation?.message ?? ''}</Text>
        <Inline gap="sm" className="mt-4 justify-end">
          <Button size="sm" variant="ghost" onClick={() => cellEdit.dispatch({ type: 'validation.reject' })}>返回编辑</Button>
          <Button size="sm" variant="primary" onClick={() => cellEdit.dispatch({ type: 'validation.confirm' })}>仍然提交</Button>
        </Inline>
      </Dialog>
    </>
  );
}
