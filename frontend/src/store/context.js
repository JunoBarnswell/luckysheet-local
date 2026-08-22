import { defaultConfig } from "./defaults";

function deepClone(value) {
    if (value == null || typeof value !== "object") {
        return value;
    }
    if (typeof window !== "undefined" && typeof window.jQuery === "function") {
        return window.jQuery.extend(true, Array.isArray(value) ? [] : {}, value);
    }
    return JSON.parse(JSON.stringify(value));
}

export function createStoreFields() {
    const store = deepClone(defaultConfig.defaultStore);
    store.loadingObj = {};
    store.definedNames = [];
    store.defaultFontSize = 10;
    store.cooperativeEdit = {
        usernameTimeout: {},
        changeCollaborationSize: [],
        allDataColumnlen: [],
        merge_range: {},
        checkoutData: [],
    };
    store.asyncLoad = ["core"];
    store.defaultCell = {
        bg: null,
        bl: 0,
        ct: { fa: "General", t: "n" },
        fc: "rgb(51, 51, 51)",
        ff: 0,
        fs: 11,
        ht: 1,
        it: 0,
        vt: 1,
        m: "",
        v: "",
    };
    store.conditionFormatCells = {};
    store.toJsonOptions = {};
    store.plugins = [];
    store.luckysheetPrint = null;
    store.allowEdit = true;
    store.limitSheetNameLength = true;
    store.defaultSheetNameMaxLength = 31;
    store.instanceId = null;
    store.domPrefix = "";
    store.multi = false;
    store.modules = {};
    store.disposers = [];
    store.portals = new Set();
    store.timers = new Set();
    store.runtime = {
        formula: {
            functions: {},
            current: { row: null, column: null, index: null, formula: null },
        },
        scroll: { requestAnimationFrameIni: true, requestId: false, timeoutId: null },
        refresh: { timeoutId: null, dirtyCells: [] },
        listener: { undoTimer: null, redoTimer: null },
        keyboard: { shiftDown: false },
        sheetBar: { initialized: false, currentItem: null, doubleClickTimer: null, oldSheetName: "" },
        resize: { gridW: 0, gridH: 0 },
    };
    return store;
}

export function createWorkbookContext(overrides) {
    const ctx = createStoreFields();
    if (overrides) {
        Object.assign(ctx, overrides);
    }
    if (!ctx.jfundo) {
        ctx.jfundo = [];
    }
    if (!ctx.jfredo) {
        ctx.jfredo = [];
    }
    ctx.disposed = false;
    return ctx;
}

export function generateInstanceId() {
    return "lsu_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
}
