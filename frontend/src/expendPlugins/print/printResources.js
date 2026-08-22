import Store from "../../store";

export const RESOURCE_WAIT_MS = 10000;

const imageCache = new Map();

export function clearPrintResourceCache() {
    imageCache.clear();
}

export function waitPrintResources(draft) {
    if (draft || typeof document === "undefined") {
        return Promise.resolve();
    }
    const root = Store.container ? document.getElementById(Store.container) : document;
    if (!root) {
        return Promise.resolve();
    }
    const canvases = root.querySelectorAll(
        ".luckysheet-data-visualization-chart canvas, .luckysheet-modal-dialog-slider canvas"
    );
    const pending = [];
    canvases.forEach(function (cv) {
        if (cv.width > 0 && cv.height > 0) {
            return;
        }
        pending.push(
            new Promise(function (resolve) {
                const start = Date.now();
                const tick = function () {
                    if ((cv.width > 0 && cv.height > 0) || Date.now() - start > RESOURCE_WAIT_MS) {
                        resolve();
                        return;
                    }
                    setTimeout(tick, 50);
                };
                tick();
            })
        );
    });
    if (!pending.length) {
        return Promise.resolve();
    }
    return Promise.race([
        Promise.all(pending),
        new Promise(function (resolve) {
            setTimeout(resolve, RESOURCE_WAIT_MS);
        }),
    ]);
}

export function loadImage(src) {
    if (!src) {
        return Promise.resolve(null);
    }
    if (imageCache.has(src)) {
        return Promise.resolve(imageCache.get(src));
    }
    return new Promise(function (resolve) {
        const img = new Image();
        img.crossOrigin = "anonymous";
        img.onload = function () {
            imageCache.set(src, img);
            resolve(img);
        };
        img.onerror = function () {
            resolve(null);
        };
        img.src = src;
    });
}

export function collectChartCanvases() {
    if (typeof document === "undefined") {
        return [];
    }
    const root = Store.container ? document.getElementById(Store.container) : document;
    if (!root) {
        return [];
    }
    const list = [];
    root.querySelectorAll(".luckysheet-data-visualization-chart").forEach(function (node) {
        const canvas = node.querySelector("canvas");
        if (!canvas || canvas.width <= 0) {
            return;
        }
        const rect = node.getBoundingClientRect();
        list.push({ canvas: canvas, left: rect.left, top: rect.top, width: rect.width, height: rect.height, node: node });
    });
    return list;
}
