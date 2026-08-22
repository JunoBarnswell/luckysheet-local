const fs = require('fs');
const path = require('path');
const LuckyExcel = require('luckyexcel');

const sampleFile = path.join(__dirname, '..', 'House cleaning checklist.xlsx');

function transformExcelBuffer(buffer) {
    return new Promise((resolve, reject) => {
        LuckyExcel.transformExcelToLucky(buffer, (exportJson) => {
            resolve(exportJson);
        }, (err) => {
            reject(err || new Error('Failed to parse xlsx'));
        });
    });
}

const fn_luckyexcel = async (ctx) => {
    if (!fs.existsSync(sampleFile)) {
        ctx.status = 404;
        ctx.body = { error: 'Sample xlsx not found. Place a .xlsx file at project root or use POST /luckyexcel/upload.' };
        return;
    }

    try {
        const data = fs.readFileSync(sampleFile);
        ctx.body = await transformExcelBuffer(data);
    } catch (err) {
        ctx.status = 500;
        ctx.body = { error: err.message };
    }
};

const fn_upload = async (ctx) => {
    const file = ctx.request.files && ctx.request.files.file;
    if (!file) {
        ctx.status = 400;
        ctx.body = { error: 'Missing file field "file"' };
        return;
    }

    const filePath = file.filepath || file.path;
    try {
        const data = fs.readFileSync(filePath);
        ctx.body = await transformExcelBuffer(data);
    } catch (err) {
        ctx.status = 400;
        ctx.body = { error: err.message || 'Invalid xlsx file' };
    } finally {
        if (filePath && fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
        }
    }
};

module.exports = {
    'GET /luckyexcel': fn_luckyexcel,
    'POST /luckyexcel/upload': fn_upload,
};
