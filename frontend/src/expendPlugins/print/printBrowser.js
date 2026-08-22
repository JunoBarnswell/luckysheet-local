import { PrintDirection } from "./printLayout";

export const PRINT_CONTAINER_CLASS = "printing-canvas-container";
export const PRINT_CANVAS_CLASS = "printing-canvas";

export function createPrintStyle(w, h, direction, sessionId) {
    const pageW = w;
    const pageH = h;
    const css = [
        "." + PRINT_CANVAS_CLASS + " {",
        "  page-break-after: always!important;",
        "  height: " + pageH + "px;",
        "  width: " + pageW + "px;",
        "  position: relative;",
        "}",
        "@media print {",
        "  html { width: fit-content; }",
        "  body { overflow: auto!important; width: fit-content; }",
        "  @page { size: " + pageW + "px " + pageH + "px; margin: 0; visibility: hidden; }",
        "  body > * { display: none!important; }",
        "  ." + PRINT_CONTAINER_CLASS + '[data-print-session="' + sessionId + '"], .' + PRINT_CONTAINER_CLASS + '[data-print-session="' + sessionId + '"] * {',
        "    display: block!important;",
        "    height: fit-content;",
        "    overflow: visible;",
        "    top: 0;",
        "    width: fit-content;",
        "  }",
        "  ." + PRINT_CANVAS_CLASS + " {",
        "    height: " + pageH + "px;",
        "    width: " + pageW + "px;",
        "    position: relative;",
        "  }",
        "}",
    ].join("\n");
    const style = document.createElement("style");
    style.innerHTML = css;
    style.className = "offline-printing-css";
    return style;
}

export function ensurePrintStyleTag(extraRules) {
    if (typeof document === "undefined") {
        return;
    }
    if (document.getElementById("luckysheet-print-inline-style")) {
        return;
    }
    const style = document.createElement("style");
    style.id = "luckysheet-print-inline-style";
    style.textContent = [
        ".luckysheet-print-box canvas{display:block;margin:8px auto;box-shadow:0 1px 4px rgba(0,0,0,.12);background:#fff;}",
        ".luckysheet-print-panel{display:flex;gap:12px;min-height:360px;}",
        ".luckysheet-print-settings{flex:0 0 280px;max-height:480px;overflow:auto;padding-right:8px;border-right:1px solid #e5e7eb;}",
        ".luckysheet-print-preview-pane{flex:1;overflow:auto;max-height:480px;}",
        ".luckysheet-print-section{margin:12px 0;padding-top:8px;border-top:1px solid #f3f4f6;}",
        ".luckysheet-print-section-title{font-weight:600;font-size:13px;margin-bottom:6px;}",
        ".luckysheet-print-row{display:flex;align-items:center;margin:6px 0;gap:8px;flex-wrap:wrap;}",
        ".luckysheet-print-row label{min-width:72px;font-size:12px;}",
        ".luckysheet-print-row input[type=text],.luckysheet-print-row input[type=number],.luckysheet-print-row select{flex:1;min-width:120px;height:28px;font-size:12px;}",
        ".luckysheet-print-radio{display:flex;gap:12px;flex-wrap:wrap;}",
        ".luckysheet-print-preview{position:fixed;inset:0;background:#fff;z-index:100010;overflow:auto;}",
        extraRules || "",
    ].join("");
    document.head.appendChild(style);
}
