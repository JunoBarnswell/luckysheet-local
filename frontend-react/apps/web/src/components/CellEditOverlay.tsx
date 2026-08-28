import { useMemo } from 'react';
import { Box, Button, Dialog, Inline, Text } from '@react-sheets/ui-system';
import type { CanvasRenderEngine, CellContentLayoutResult, CellRenderData } from '@react-sheets/render-engine';
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

export function CellEditOverlay({ cellEdit, engine, host, scrollTick, sheet }: CellEditOverlayProps) {
  const edit = useCellEdit(cellEdit);
  const session = edit.session;
  const displayTarget = session?.target.display;
  const mergedRange = session?.target.mergedRange;
  const draftText = session?.draft.text ?? '';
  const cellStyle = session && displayTarget?.sheetId === sheet.id ? sheet.getCell(displayTarget.row, displayTarget.column)?.style : undefined;
  const layout = useMemo<CellContentLayoutResult | null>(() => {
    void scrollTick;
    if (!engine || !displayTarget || displayTarget.sheetId !== sheet.id) return null;
    const source: CellRenderData = {
      value: session?.originalCell?.value ?? null,
      style: cellStyle,
      ...(session?.originalCell?.formula ? { formula: session.originalCell.formula } : {}),
      ...(session?.originalCell?.richText ? { richText: session.originalCell.richText } : {}),
      ...(session?.originalCell?.phonetic ? { phonetic: session.originalCell.phonetic } : {}),
    };
    const layout = engine.cellContentLayoutAtScreen(displayTarget, draftText, 'edit', {
      ...(mergedRange ? { range: mergedRange } : {}),
      cell: source,
      ...(session ? { caret: session.caret } : {}),
    });
    return layout;
  }, [cellStyle, displayTarget, draftText, engine, mergedRange, scrollTick, session, sheet.id]);
  const rect = layout?.editRect ?? null;

  if (!session) return null;
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
          <CellEditor editorSurface={session.editorSurface} cellEdit={cellEdit} cellStyle={cellStyle} draft={session.draft} caret={session.caret} layout={layout} />
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
