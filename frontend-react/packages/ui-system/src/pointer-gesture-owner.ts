export const POINTER_GESTURE_OWNER_ATTRIBUTE = 'data-pointer-gesture-owner';

export type PointerGestureOwner =
  | 'worksheet'
  | 'scrollbar-horizontal'
  | 'scrollbar-vertical'
  | 'cell-editor'
  | 'floating-object'
  | 'popover-menu'
  | 'other-ui-control';

interface PointerGestureSession {
  owner: PointerGestureOwner;
  surface: EventTarget;
}

const sessionsByDocument = new WeakMap<Document, Map<number, PointerGestureSession>>();

function sessionsFor(ownerDocument: Document): Map<number, PointerGestureSession> {
  const existing = sessionsByDocument.get(ownerDocument);
  if (existing) return existing;
  const sessions = new Map<number, PointerGestureSession>();
  sessionsByDocument.set(ownerDocument, sessions);
  return sessions;
}

/**
 * Claims one pointer sequence for one interaction state machine. A competing
 * owner fails closed instead of allowing two consumers to mutate state from
 * the same pointerId.
 */
export function claimPointerGesture(
  ownerDocument: Document,
  pointerId: number,
  owner: PointerGestureOwner,
  surface: EventTarget,
): boolean {
  const sessions = sessionsFor(ownerDocument);
  const current = sessions.get(pointerId);
  if (current) return current.owner === owner && current.surface === surface;
  sessions.set(pointerId, { owner, surface });
  return true;
}

export function ownsPointerGesture(
  ownerDocument: Document,
  pointerId: number,
  owner: PointerGestureOwner,
  surface: EventTarget,
): boolean {
  const current = sessionsByDocument.get(ownerDocument)?.get(pointerId);
  return current?.owner === owner && current.surface === surface;
}

export function releasePointerGesture(
  ownerDocument: Document,
  pointerId: number,
  owner: PointerGestureOwner,
  surface: EventTarget,
): boolean {
  if (!ownsPointerGesture(ownerDocument, pointerId, owner, surface)) return false;
  sessionsByDocument.get(ownerDocument)?.delete(pointerId);
  return true;
}

export function releasePointerGesturesForSurface(ownerDocument: Document, surface: EventTarget): void {
  const sessions = sessionsByDocument.get(ownerDocument);
  if (!sessions) return;
  for (const [pointerId, session] of sessions) {
    if (session.surface === surface) sessions.delete(pointerId);
  }
}

/** Returns the nearest declarative interaction boundary for a DOM target. */
export function resolvePointerGestureOwner(target: EventTarget | null): string | null {
  const candidate = target as { closest?: (selector: string) => Element | null } | null;
  const boundary = candidate?.closest?.(`[${POINTER_GESTURE_OWNER_ATTRIBUTE}]`);
  return boundary?.getAttribute(POINTER_GESTURE_OWNER_ATTRIBUTE) ?? null;
}
