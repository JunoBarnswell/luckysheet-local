const Koa = require('koa');
const cors = require('@koa/cors');
const { koaBody } = require('koa-body');
const controller = require('./controller');

const PORT = Number(process.env.PORT || 3002);
const app = new Koa();

app.use(cors({ origin: '*' }));
app.use(koaBody({ jsonLimit: '50mb', multipart: true, formidable: { maxFileSize: 50 * 1024 * 1024 } }));
app.use(async (ctx, next) => {
    console.log(`${new Date().toISOString()} ${ctx.method} ${ctx.url}`);
    await next();
});
app.use(controller());

app.listen(PORT, '127.0.0.1', () => {
    console.log(`Luckyexcel-node listening at http://127.0.0.1:${PORT}`);
    console.log(`  Import demo : GET  /luckyexcel`);
    console.log(`  Import file : POST /luckyexcel/upload`);
    console.log(`  Export xlsx : POST /luckyToXlsx`);
});
