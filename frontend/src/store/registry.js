import { createWorkbookContext } from "./context";

const units = new Map();
let focusedId = null;
const idle = createWorkbookContext({ instanceId: "__idle__" });

export function createUnit(ctx) {
    if (!ctx || !ctx.instanceId) {
        throw new Error("createUnit requires instanceId");
    }
    units.set(ctx.instanceId, ctx);
    focusedId = ctx.instanceId;
    return ctx;
}

export function disposeUnit(id) {
    if (!id || !units.has(id)) {
        return false;
    }
    units.delete(id);
    if (focusedId === id) {
        const rest = listUnits();
        focusedId = rest.length ? rest[rest.length - 1] : null;
    }
    return true;
}

export function getUnit(id) {
    return units.get(id) || null;
}

export function listUnits() {
    return Array.from(units.keys());
}

export function getFocusedId() {
    return focusedId;
}

export function getFocusedContext() {
    if (focusedId && units.has(focusedId)) {
        return units.get(focusedId);
    }
    return idle;
}

export function focusUnit(id) {
    if (!id || !units.has(id)) {
        return false;
    }
    if (focusedId === id) {
        return true;
    }
    focusedId = id;
    return true;
}

export function withInstance(id, fn) {
    const prev = focusedId;
    if (id && id !== prev) {
        focusUnit(id);
    }
    try {
        return fn(getFocusedContext());
    } finally {
        if (prev && prev !== id && units.has(prev)) {
            focusUnit(prev);
        }
    }
}
