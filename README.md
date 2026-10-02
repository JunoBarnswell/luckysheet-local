# React Sheets

React Sheets is a browser-based spreadsheet with a React and TypeScript Canvas editor and a Java 21 workbook service. It supports local workbook editing and XLSX import/export, plus server-backed workbooks, access control, sharing, and collaboration.

## Architecture

- `frontend-react/` contains the web application and the canonical spreadsheet model, formula, rendering, command, persistence, and OOXML packages.
- `backend/` contains the Spring Boot workbook service. It owns durable workbook operations, revisions, access control, and collaboration events.
- `contracts/` contains shared protocol and API contracts.
- `docs/` contains product and acceptance documentation.

The editor sends mutations through the command runtime and operation API. The backend validates and commits workbook operations before broadcasting committed revision events. The browser uses REST and WebSocket endpoints at the same public origin; local Vite development proxies `/api` and `/ws` to the backend.

## Requirements

- Node.js 24.x and npm 11 or newer
- Java 21 and Maven

## Run locally

Start the backend from one terminal:

```powershell
cd backend
mvn spring-boot:run
```

The backend listens on `http://127.0.0.1:8082` and uses the local H2 profile by default.

Start the web application from another terminal:

```powershell
cd frontend-react
npm install
npm run dev
```

Open `http://127.0.0.1:4180/`. The Vite proxy targets `http://127.0.0.1:8082` by default; set `REACT_SHEETS_API_ORIGIN` to use a different local backend. Configure the browser OIDC public-client values shown in [`frontend-react/.env.example`](frontend-react/.env.example) for the identity provider used by your environment. Do not put a client secret in browser configuration.

## Build and checks

From `frontend-react/`:

```powershell
npm run build
```

From `backend/`:

```powershell
mvn test
```

## Documentation

- [Web application development and architecture](frontend-react/README.md)
- [Backend configuration and API behavior](backend/README.md)
- [Workbook hub requirements](frontend-react/docs/workbook-hub-prd.md)
- [Acceptance guide](frontend-react/docs/acceptance.md)
- [OpenAPI contract](contracts/openapi.yaml)
