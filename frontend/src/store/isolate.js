import defaultConfig from "./defaults";
import { focusUnit, getFocusedId, getFocusedContext, getUnit } from "./registry";

const boxes = {
    formula: null,
    server: null,
    sheet: null,
    pivot: null,
    image: null,
    dv: null,
    freezen: null,
    config: null,
    getScrollFlags: function () { return {}; },
    setScrollFlags: function () {},
};

const CONFIG_DEFAULTS = {};

export function bindIsolateModules(modules) {
    Object.assign(boxes, modules || {});
    if (boxes.config && !Object.keys(CONFIG_DEFAULTS).length) {
        Object.assign(CONFIG_DEFAULTS, boxes.config);
    }
}

function snapshotData(obj) {
    const out = {};
    if (!obj) {
        return out;
    }
    Object.keys(obj).forEach(function (key) {
        if (typeof obj[key] !== "function") {
            out[key] = obj[key];
        }
    });
    return out;
}

function restoreData(obj, snap) {
    if (!obj || !snap) {
        return;
    }
    Object.keys(snap).forEach(function (key) {
        obj[key] = snap[key];
    });
}

function applyDefaults(obj, defaults) {
    if (!obj || !defaults) {
        return;
    }
    const cloned = typeof window !== "undefined" && window.jQuery
        ? window.jQuery.extend(true, {}, defaults)
        : JSON.parse(JSON.stringify(defaults));
    Object.keys(cloned).forEach(function (key) {
        obj[key] = cloned[key];
    });
}

export function snapshotModules(ctx) {
    if (!ctx) {
        return;
    }
    ctx.moduleSnap = {
        formula: snapshotData(boxes.formula),
        server: snapshotData(boxes.server),
        sheet: snapshotData(boxes.sheet),
        pivot: snapshotData(boxes.pivot),
        image: snapshotData(boxes.image),
        dv: snapshotData(boxes.dv),
        freezen: snapshotData(boxes.freezen),
        config: snapshotData(boxes.config),
        scroll: boxes.getScrollFlags(),
    };
}

export function restoreModules(ctx) {
    if (!ctx || !ctx.moduleSnap) {
        return;
    }
    const snap = ctx.moduleSnap;
    restoreData(boxes.formula, snap.formula);
    restoreData(boxes.server, snap.server);
    restoreData(boxes.sheet, snap.sheet);
    restoreData(boxes.pivot, snap.pivot);
    restoreData(boxes.image, snap.image);
    restoreData(boxes.dv, snap.dv);
    restoreData(boxes.freezen, snap.freezen);
    restoreData(boxes.config, snap.config);
    if (snap.scroll) {
        boxes.setScrollFlags(snap.scroll);
    }
}

export function resetModulesForNewInstance() {
    applyDefaults(boxes.formula, defaultConfig.defaultFormula);
    applyDefaults(boxes.sheet, defaultConfig.defaultSheet);
    applyDefaults(boxes.pivot, defaultConfig.defaultPivotTable);
    applyDefaults(boxes.image, defaultConfig.defaultImage);
    applyDefaults(boxes.dv, defaultConfig.defaultDataVerification);
    applyDefaults(boxes.freezen, defaultConfig.defaultFreezen);
    const server = boxes.server;
    const serverDefaults = defaultConfig.defaultServer;
    if (server) {
        Object.keys(serverDefaults).forEach(function (key) {
            if (key === "websocket") {
                server.websocket = null;
                return;
            }
            if (key === "cellClock") {
                server.cellClock = {};
                return;
            }
            if (key === "outboundQueue") {
                server.outboundQueue = [];
                return;
            }
            server[key] = serverDefaults[key];
        });
    }
    if (boxes.config) {
        Object.keys(CONFIG_DEFAULTS).forEach(function (key) {
            boxes.config[key] = CONFIG_DEFAULTS[key];
        });
    }
    boxes.setScrollFlags({
        scrollRequestAnimationFrameIni: true,
        scrollRequestAnimationFrame: false,
        scrollTimeOutCancel: null,
    });
}

export function withWorkbook(id, fn) {
    const prev = getFocusedId();
    if (id && id !== prev) {
        focusWorkbook(id);
    }
    try {
        return fn(getFocusedContext());
    } finally {
        if (prev && prev !== id && getUnit(prev)) {
            focusWorkbook(prev);
        }
    }
}

export function focusWorkbook(id) {
    if (!id || !getUnit(id)) {
        return false;
    }
    if (getFocusedId() === id) {
        return true;
    }
    const prev = getFocusedId();
    if (prev && getUnit(prev)) {
        snapshotModules(getFocusedContext());
    }
    focusUnit(id);
    restoreModules(getUnit(id));
    return true;
}

export function closeServerSocket() {
    const server = boxes.server;
    if (!server || server.websocket == null) {
        return;
    }
    try {
        server.intentionalClose = true;
        server.websocket.onopen = null;
        server.websocket.onmessage = null;
        server.websocket.onerror = null;
        server.websocket.onclose = null;
        if (server.websocket.readyState === 0 || server.websocket.readyState === 1) {
            server.websocket.close(1000, "instance destroy");
        }
    } catch (e) { /* ignore */ }
    server.websocket = null;
}
