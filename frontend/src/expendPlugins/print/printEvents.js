import luckysheetConfigsetting from "../../controllers/luckysheetConfigsetting";

const HOOK_MAP = {
    BeforeSheetPrintOpen: "beforeSheetPrintOpen",
    AfterSheetPrintOpen: "afterSheetPrintOpen",
    BeforeSheetPrintConfirm: "beforeSheetPrintConfirm",
    AfterSheetPrintConfirm: "afterSheetPrintConfirm",
    BeforeSheetPrintCanceled: "beforeSheetPrintCanceled",
};

function getHook(name) {
    const hook = luckysheetConfigsetting && luckysheetConfigsetting.hook;
    if (!hook) {
        return null;
    }
    return hook[name] || hook[HOOK_MAP[name]];
}

export function emitPrintEvent(eventName, payload) {
    const fn = getHook(HOOK_MAP[eventName] || eventName);
    if (typeof fn !== "function") {
        return true;
    }
    try {
        const result = fn(payload);
        return result !== false;
    } catch (e) {
        console.error("[luckysheet-print]", eventName, e);
        return true;
    }
}

export function emitBeforeSheetPrintOpen(payload) {
    return emitPrintEvent("BeforeSheetPrintOpen", payload);
}

export function emitAfterSheetPrintOpen(payload) {
    emitPrintEvent("AfterSheetPrintOpen", payload);
}

export function emitBeforeSheetPrintConfirm(payload) {
    return emitPrintEvent("BeforeSheetPrintConfirm", payload);
}

export function emitAfterSheetPrintConfirm(payload) {
    emitPrintEvent("AfterSheetPrintConfirm", payload);
}

export function emitBeforeSheetPrintCanceled(payload) {
    emitPrintEvent("BeforeSheetPrintCanceled", payload);
}
