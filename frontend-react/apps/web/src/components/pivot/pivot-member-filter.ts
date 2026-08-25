import { pivotMemberKey, pivotMemberKeyEquals, type PivotMemberKey } from '@react-sheets/core-model';
import type { PivotManualFilterState } from './pivot-contract';

/**
 * Evaluate a member against the canonical manual filter without expanding the
 * full member domain.  This is important for high-cardinality fields: the
 * visible list is only a presentation window, while the filter state remains
 * a compact include/exclude delta.
 */
export function pivotManualMemberSelected(state: PivotManualFilterState, member: PivotMemberKey): boolean {
  if (state.mode === 'all') return true;
  const listed = state.memberKeys.some((candidate) => pivotMemberKeyEquals(candidate, member));
  return state.mode === 'include' ? listed : !listed;
}

function uniqueKeys(keys: readonly PivotMemberKey[]): PivotMemberKey[] {
  const seen = new Set<string>();
  return keys.filter((key) => {
    const identity = pivotMemberKey(key);
    if (seen.has(identity)) return false;
    seen.add(identity);
    return true;
  });
}

/**
 * Apply a checkbox delta to the canonical filter state.  It intentionally
 * never materializes all members, so toggling a 10,001st member cannot erase
 * members outside the bounded display window.
 */
export function applyPivotManualMemberDelta(
  state: PivotManualFilterState,
  targets: readonly PivotMemberKey[],
  checked: boolean,
): PivotManualFilterState {
  const changed = uniqueKeys(targets);
  if (changed.length === 0) return { mode: state.mode, memberKeys: [...state.memberKeys] };
  const contains = (key: PivotMemberKey) => changed.some((target) => pivotMemberKeyEquals(target, key));

  if (state.mode === 'all') {
    return checked ? { mode: 'all', memberKeys: [] } : { mode: 'exclude', memberKeys: changed };
  }
  if (state.mode === 'include') {
    return {
      mode: 'include',
      memberKeys: uniqueKeys(checked ? [...state.memberKeys, ...changed] : state.memberKeys.filter((key) => !contains(key))),
    };
  }
  return {
    mode: 'exclude',
    memberKeys: uniqueKeys(checked ? state.memberKeys.filter((key) => !contains(key)) : [...state.memberKeys, ...changed]),
  };
}

/**
 * Convert the explicit mode selector while retaining typed identities.  The
 * common `all -> exclude` transition is represented by an empty exclusion
 * delta instead of allocating the complete domain.
 */
export function convertPivotManualFilterMode(
  state: PivotManualFilterState,
  mode: PivotManualFilterState['mode'],
  domain: readonly PivotMemberKey[],
): PivotManualFilterState {
  if (mode === 'all') return { mode: 'all', memberKeys: [] };
  if (mode === state.mode) return { mode, memberKeys: [...state.memberKeys] };
  if (mode === 'exclude' && state.mode === 'all') return { mode: 'exclude', memberKeys: [] };
  const selected = domain.filter((member) => pivotManualMemberSelected(state, member));
  if (mode === 'include') return { mode, memberKeys: selected };
  return { mode, memberKeys: domain.filter((member) => !selected.some((candidate) => pivotMemberKeyEquals(candidate, member))) };
}
