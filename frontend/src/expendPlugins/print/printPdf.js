import { jsPDF } from "jspdf";
import { PrintDirection } from "./printLayout";

export function exportPdfFromCanvases(canvases, layout, filename) {
    if (!canvases || !canvases.length) {
        return Promise.resolve({ pageCount: 0 });
    }
    const paper = canvases[0];
    const w = paper.width / (window.devicePixelRatio || 1);
    const h = paper.height / (window.devicePixelRatio || 1);
    const orientation = layout && layout.direction === PrintDirection.Landscape ? "landscape" : "portrait";

    const pdf = new jsPDF({
        orientation: orientation,
        unit: "px",
        format: [w, h],
        hotfixes: ["px_scaling"],
    });
    canvases.forEach(function (canvas, index) {
        if (index > 0) {
            pdf.addPage([w, h], orientation);
        }
        const dataUrl = canvas.toDataURL("image/jpeg", 0.92);
        pdf.addImage(dataUrl, "JPEG", 0, 0, w, h);
    });
    const name = filename || "luckysheet-print.pdf";
    pdf.save(name);
    return Promise.resolve({ pageCount: canvases.length, filename: name });
}

export function exportPdf(pack, layout, render, drawPageFn, filename) {
    const entries = pack.workbookPages || (pack.pages || []).map(function (page) {
        return { file: pack.file, page: page, pack: pack };
    });
    const canvases = entries.map(function (entry, index) {
        const meta = {
            pageIndex: index,
            pageTotal: entries.length,
            sheetPage: pack.pages.indexOf(entry.page) + 1,
            sheetPageTotal: entry.pack.pages.length,
        };
        return drawPageFn(entry.page, entry.file, layout, render, entry.pack, meta);
    });
    return exportPdfFromCanvases(canvases, layout, filename);
}
