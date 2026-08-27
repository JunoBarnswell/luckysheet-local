# Windows desktop packaging

Run `npm run desktop:dist` from `frontend-react` to build the Vite frontend and produce the Windows NSIS installer under `frontend-react/dist/desktop`.

The packaged app serves the generated `dist/web` directory through Electron's secure `app://` protocol. It does not include or start the backend API.
