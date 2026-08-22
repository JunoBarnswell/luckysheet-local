
import locale from '../../locale/locale';
import { modelHTML } from "../../controllers/constant";
import { arrayRemoveItem, replaceHtml } from '../../utils/util';
import tooltip from '../../global/tooltip';
import { getSheetIndex } from '../../methods/get';
import Store from '../../store';

let selectedOption = 'allSheets';

function exportXlsx(options, config, isDemo) {
    arrayRemoveItem(Store.asyncLoad,'exportXlsx')
}

function downloadXlsx(data, filename) {
    const blob = new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
}

function exportFail(code, message, cause) {
    return { code, message, cause: cause || null };
}

function localeExport() {
    const pack = locale();
    return pack.exportXlsx || {};
}

function failMessage(error) {
    const texts = localeExport();
    if (!error) {
        return texts.serverError || '';
    }
    if (error.code === 'EMPTY_URL') {
        return texts.emptyUrl || texts.notice || texts.serverError;
    }
    if (error.code === 'NETWORK') {
        return texts.networkError || texts.serverError;
    }
    if (error.code === 'HTTP' || error.code === 'INVALID_RESPONSE') {
        return texts.invalidResponse || texts.serverError;
    }
    return error.message || texts.serverError;
}

function isXlsxBlob(blob) {
    if (!blob) {
        return false;
    }
    const type = (blob.type || '').toLowerCase();
    return (
        type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
        type === 'application/octet-stream' ||
        type === 'application/zip' ||
        type === ''
    );
}

function collectChartMap() {
    return new Promise((resolve) => {
        try {
            if (typeof luckysheet.getAllChartsBase64 !== 'function') {
                resolve({});
                return;
            }
            let settled = false;
            const timer = setTimeout(() => {
                if (!settled) {
                    settled = true;
                    resolve({});
                }
            }, 4000);
            luckysheet.getAllChartsBase64((chartMap) => {
                if (settled) {
                    return;
                }
                settled = true;
                clearTimeout(timer);
                resolve(chartMap || {});
            });
        } catch (err) {
            resolve({});
        }
    });
}

function fetchAndDownloadXlsx({url,order}, success, fail) {
    const notifyFail = (error) => {
        if (typeof fail === 'function') {
            fail(error);
        }
    };
    const notifySuccess = () => {
        if (typeof success === 'function') {
            success();
        }
    };

    if (url == null || String(url).trim() === '') {
        notifyFail(exportFail('EMPTY_URL', failMessage({ code: 'EMPTY_URL' })));
        return;
    }

    const endpoint = String(url).trim();
    const texts = localeExport();

    collectChartMap().then((chartMap) => {
        let luckyJson;
        try {
            luckyJson = luckysheet.toJson();
        } catch (err) {
            notifyFail(exportFail('SERIALIZE', err && err.message, err));
            return;
        }

        luckyJson.chartMap = chartMap;
        luckyJson.devicePixelRatio = window.devicePixelRatio;
        luckyJson.exportXlsx = { order };

        return fetch(endpoint, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(luckyJson)
        }).then(async (response) => {
            if (!response.ok) {
                let detail = '';
                try {
                    detail = await response.text();
                } catch (err) {
                    detail = '';
                }
                throw exportFail(
                    'HTTP',
                    (texts.httpError || texts.serverError || '') + ' HTTP ' + response.status,
                    { status: response.status, body: detail }
                );
            }
            const blob = await response.blob();
            if (!isXlsxBlob(blob) || blob.size === 0) {
                throw exportFail('INVALID_RESPONSE', texts.invalidResponse || texts.serverError, { type: blob.type, size: blob.size });
            }
            const filename = (luckyJson.title || 'luckysheet') + '.xlsx';
            downloadXlsx(blob, filename);
            notifySuccess();
        }).catch((error) => {
            if (error && error.code) {
                notifyFail(error);
                return;
            }
            console.error('fetch error:', error);
            notifyFail(exportFail('NETWORK', texts.networkError || texts.serverError, error));
        });
    }).catch((error) => {
        notifyFail(exportFail('CHART', error && error.message, error));
    });
}

function createExportDialog(url) {
    const texts = localeExport();
    if (url == null || String(url).trim() === '') {
        tooltip.info(texts.emptyUrl || texts.notice || texts.serverError, "");
        return;
    }

    $("#luckysheet-modal-dialog-mask").hide();
    var xlsxContainer = $("#luckysheet-export-xlsx");

    if (xlsxContainer.length === 0) {

        const _locale = locale();
        const locale_exportXlsx = _locale.exportXlsx;
        const locale_button = _locale.button;

        let content = `<div class="luckysheet-export-xlsx-content" style="padding: 10px 10px 10px 0;">
                <span>${locale_exportXlsx.range}</span>
                <select class="luckysheet-export-xlsx-select-area">
                    <option value="allSheets" selected="selected">${locale_exportXlsx.allSheets}</option>
                    <option value="currentSheet">${locale_exportXlsx.currentSheet}</option>
                </select>
        </div>`;

        $("body").append(
            replaceHtml(modelHTML, {
                id: "luckysheet-export-xlsx",
                addclass: "luckysheet-export-xlsx",
                title: locale_exportXlsx.title,
                content: content,
                botton: `<button class="btn btn-primary luckysheet-model-confirm-btn">${locale_button.confirm}</button><button class="btn btn-default luckysheet-model-close-btn">${locale_button.close}</button>`,
                style: "z-index:991",
                close: locale_button.close,
            }),
        );

        selectedOption = 'allSheets'

        $("#luckysheet-export-xlsx .luckysheet-model-confirm-btn").on('click',()=>{
            luckysheet.showLoadingProgress()

            var order = 'all'
            if(selectedOption === 'currentSheet'){
                order = getSheetIndex(Store.currentSheetIndex)
            }
            fetchAndDownloadXlsx({url: String(url).trim(), order},()=>{
                luckysheet.hideLoadingProgress()
            },(error)=>{
                luckysheet.hideLoadingProgress()
                tooltip.info(failMessage(error), "");
            })
            $("#luckysheet-export-xlsx").hide()
        })

        $("#luckysheet-export-xlsx .luckysheet-export-xlsx-select-area").change(function() {
            selectedOption = $(this).val();
          });

    }



    let $t = $("#luckysheet-export-xlsx").find(".luckysheet-modal-dialog-content").css("min-width", 350).end(),
        myh = $t.outerHeight(),
        myw = $t.outerWidth();
    let winw = $(window).width(),
        winh = $(window).height();
    let scrollLeft = $(document).scrollLeft(),
        scrollTop = $(document).scrollTop();
    $("#luckysheet-export-xlsx")
        .css({ left: (winw + scrollLeft - myw) / 2, top: (winh + scrollTop - myh) / 3 })
        .show();

}

export { exportXlsx, downloadXlsx, fetchAndDownloadXlsx, createExportDialog, exportFail }
