const units = new Map();
let focusedId = null;

export function createUnit(ctx) {
    if (!ctx || !ctx.instanceId) {
        throw new Error("createUnit requires instanceId");
    }
    ctx.disposed = false;
    units.set(ctx.instanceId, ctx);
    focusedId = ctx.instanceId;
    return ctx;
}

export function disposeUnit(id) {
    if (!id || !units.has(id)) {
        return false;
    }
    const ctx = units.get(id);
    ctx.disposed = true;
    if (ctx.disposers) {
        ctx.disposers.splice(0).forEach(function (dispose) {
            try { dispose(); } catch (e) { /* best-effort teardown */ }
        });
    }
    if (ctx.timers) {
        ctx.timers.forEach(function (timer) {
            clearTimeout(timer);
            clearInterval(timer);
        });
        ctx.timers.clear();
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
    return null;
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
    if (!id || !units.has(id)) {
        throw new Error("Unknown LuckySheet instance: " + id);
    }
    if (id !== prev) {
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

export function requireFocusedContext() {
    const context = getFocusedContext();
    if (!context) {
        throw new Error("No focused LuckySheet instance");
    }
    return context;
}
