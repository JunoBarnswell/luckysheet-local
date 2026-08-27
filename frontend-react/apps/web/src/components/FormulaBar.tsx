import { useLayoutEffect, useRef, useState } from 'react';
import { Box, Button, FormulaBar as FormulaBarShell, Text, type FormulaBarProps as ShellFormulaBarProps } from '@react-sheets/ui-system';
import { useCellEdit, type AppPhase, type CellEditController } from '@react-sheets/spreadsheet-app';
import { formulaBarLabels, type Locale } from '../i18n';
import { toCanonicalKeyGesture } from '../editor/cell-edit-gesture';

export interface FormulaBarProps {
  cellName: string;
  cellEdit: CellEditController;
  disabled: boolean;
  readyFormula: string;
  locale: Locale;
  onCommitReady: () => void;
  onNameBoxCommit?: (value: string) => void;
  onOpenNameManager?: () => void;
  onOpenWizard?: () => void;
  phase: AppPhase;
}

export function FormulaBar({ cellEdit, locale, phase, readyFormula, onCommitReady, ...props }: FormulaBarProps) {
  const edit = useCellEdit(cellEdit);
  const hostRef = useRef<HTMLElement | null>(null);
  const [autocompleteAnchor, setAutocompleteAnchor] = useState<{ left: number; top: number; width: number } | null>(null);
  const formula = edit.session?.draft.text ?? readyFormula;
  const commit = () => {
    if (edit.session) cellEdit.dispatch({ type: 'commit', moveAfter: 'down' });
    else onCommitReady();
  };
  const shellProps: ShellFormulaBarProps = {
    ...props,
    formula,
    ...(edit.session ? { caret: edit.session.caret } : {}),
    onCancel: () => cellEdit.dispatch({ type: 'cancel' }),
    onChange: (value) => cellEdit.dispatch({ type: 'text.replace', text: value, caret: { start: value.length, end: value.length } }),
    onCommit: commit,
    onCaretChange: (start, end) => cellEdit.dispatch({ type: 'caret.set', caret: { start, end } }),
    onCompositionStart: () => cellEdit.dispatch({ type: 'composition.start' }),
    onCompositionUpdate: (value) => cellEdit.dispatch({ type: 'composition.update', text: value }),
    onCompositionEnd: (value, start, end) => cellEdit.dispatch({ type: 'composition.end', text: value, caret: { start, end } }),
    onFocusFormula: () => {
      if (edit.session) {
        cellEdit.dispatch({ type: 'surface.focus', surface: 'formula-bar' });
        if (edit.status === 'enter') cellEdit.dispatch({ type: 'status.toggle' });
      } else cellEdit.dispatch({ type: 'begin.request', source: 'formula-bar', surface: 'formula-bar' });
    },
    onFormulaKeyDown: (event) => {
      const result = cellEdit.dispatch({ type: 'keyboard', gesture: toCanonicalKeyGesture(event) });
      if (!result.handled) return;
      if (result.preventDefault) event.preventDefault();
      event.stopPropagation();
    },
    labels: formulaBarLabels(locale, phase),
  };
  const autocomplete = edit.session?.surface === 'formula-bar' && edit.session.overlay.kind === 'autocomplete'
    ? edit.session.overlay
    : null;
  const functionHint = edit.session?.surface === 'formula-bar' && edit.session.overlay.kind === 'function-hint'
    ? edit.session.overlay
    : null;
  const formulaOverlayOpen = autocomplete !== null || functionHint !== null;
  useLayoutEffect(() => {
    const host = hostRef.current;
    const input = host?.querySelector<HTMLElement>('[data-testid="formula-input"]');
    if (!formulaOverlayOpen || !host || !input) {
      setAutocompleteAnchor(null);
      return;
    }
    const hostRect = host.getBoundingClientRect();
    const inputRect = input.getBoundingClientRect();
    setAutocompleteAnchor({ left: inputRect.left - hostRect.left, top: inputRect.bottom - hostRect.top + 1, width: inputRect.width });
  }, [formulaOverlayOpen]);
  return (
    <Box ref={hostRef} className="relative">
      <FormulaBarShell {...shellProps} />
      {autocomplete && autocompleteAnchor ? (
        <Box className="absolute z-50 max-h-64 min-w-72 overflow-y-auto rounded border border-[#9ba8b6] bg-white py-1 shadow-lg" style={{ left: autocompleteAnchor.left, top: autocompleteAnchor.top, width: Math.max(288, autocompleteAnchor.width) }}>
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
      {functionHint && autocompleteAnchor ? (
        <Box className="absolute z-50 rounded border border-[#9ba8b6] bg-amber-50 px-2 py-1 shadow-lg" style={{ left: autocompleteAnchor.left, top: autocompleteAnchor.top }}>
          <Text size="xs"><Text size="xs" weight="semibold">{functionHint.functionName}</Text> · argument {functionHint.argumentIndex}</Text>
        </Box>
      ) : null}
    </Box>
  );
}
