import React, { useEffect, useId, useState } from 'react';
import { CANONICAL_FONT_FAMILIES, normalizeFontFamily } from '@react-sheets/core-model';
import { TextInput } from '@react-sheets/ui-system';

export interface FontFamilyControlProps {
  value?: string;
  fallbackValue?: string;
  mixed?: boolean;
  disabled?: boolean;
  label: string;
  mixedPlaceholder?: string;
  className?: string;
  testId?: string;
  onCommit: (fontFamily: string) => void;
}

/** Returns the canonical value that a font-family editor may commit. */
export function commitFontFamilyValue(value: unknown): string | undefined {
  if (typeof value !== 'string' || value.trim().length === 0) return undefined;
  return normalizeFontFamily(value);
}

/**
 * Shared editable font-family editor used by HOME and Format Cells.  The
 * datalist is only a suggestion surface: imported or locally installed fonts
 * remain editable instead of being discarded by a closed select list.
 */
export function FontFamilyControl({
  value,
  fallbackValue,
  mixed = false,
  disabled = false,
  label,
  mixedPlaceholder,
  className,
  testId,
  onCommit,
}: FontFamilyControlProps): React.ReactElement {
  const listId = `font-family-options-${useId().replace(/:/g, '')}`;
  const displayedValue = mixed ? '' : value ?? fallbackValue ?? '';
  const [draft, setDraft] = useState(displayedValue);
  const [invalid, setInvalid] = useState(false);

  useEffect(() => {
    setDraft(displayedValue);
    setInvalid(false);
  }, [displayedValue]);

  const commit = () => {
    const canonical = commitFontFamilyValue(draft);
    if (canonical === undefined) {
      setInvalid(true);
      return;
    }
    setInvalid(false);
    setDraft(canonical);
    onCommit(canonical);
  };

  return (
    <>
      <TextInput
        aria-label={label}
        aria-invalid={invalid || undefined}
        className={className}
        data-testid={testId}
        disabled={disabled}
        error={invalid}
        list={listId}
        placeholder={mixed ? mixedPlaceholder : undefined}
        value={draft}
        onChange={(event) => {
          setDraft(event.target.value);
          if (invalid) setInvalid(false);
        }}
        onBlur={commit}
        onKeyDown={(event) => {
          if (event.key === 'Enter') {
            event.preventDefault();
            commit();
          }
        }}
      />
      <datalist id={listId}>
        {CANONICAL_FONT_FAMILIES.map((fontFamily) => <option key={fontFamily} value={fontFamily} />)}
      </datalist>
    </>
  );
}
