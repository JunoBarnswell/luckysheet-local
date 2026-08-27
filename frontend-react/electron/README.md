# Windows desktop packaging

Run `npm run desktop:dist` from `frontend-react` to build the Vite frontend and produce the Windows NSIS installer under `frontend-react/dist/desktop`.

The packaged app serves `dist/web` through Electron's secure `app://` protocol and proxies `/api` to `REACT_SHEETS_API_ORIGIN` (default `http://127.0.0.1:8082`). Collaboration connects directly to the matching `ws://` or `wss://` endpoint. The installer does not bundle the Java backend, so that backend must be running before a remote workbook can open.
