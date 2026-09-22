# Workbook Hub and Cloud File Center

## Product boundary

This document is the executable product baseline for the React Sheets Web application. It supersedes the legacy local-only catalog and the in-session XLSX replacement behavior. The user supplied reference PRD and `ChatGPT Image 2026年8月24日 02_52_51.png` remain the visual source of truth; this document resolves their ambiguous implementation suggestions.

```text
Workbook Hub manages many workbook resources
        ↓
/workbooks/:unitId owns exactly one WorkbookSession
        ↓
one workbook contains many sheets
```

The application never keeps `Map<unitId, WorkbookSession>`. Opening another workbook disposes the previous session; opening in a new browser tab creates an independent session.

## Decisions

- Web runtime: desktop-style chrome is visual only. Help, Settings and the file-center exit are real controls; OS window glyphs are decorative and non-interactive.
- Identity: OIDC Authorization Code + PKCE. The browser obtains bearer tokens from `AuthSession`; no token is stored in a workbook, URL, or page-session memory record.
- Storage: server state is authoritative for remote workbooks. Local-only workbooks, pending-operation journals, source blocks, overlays, assets, and native artifacts exist only in the current page-session memory and are cleared on reload or page close.
- File domain: personal/team spaces and nested folders are first-class. Shared-with-me is a virtual view and never copies a workbook.
- Permissions: Owner controls sharing, cross-space moves, trash and purge. Editor may rename and move inside a permitted space. Commenter/Viewer are read-only. Every accessible role can export a copy.
- XLSX: importing always creates a new canonical workbook identity. The original XLSX package is stored locally and, for remote workbooks, in a server-owned LOB record. Export uses the latest snapshot plus that artifact.
- Persistence: client operations commit through the REST operation endpoint. WebSocket only carries committed broadcasts and presence/cursor events.
- Lifecycle: delete is a soft move to trash. Artifact, history, blocks and local mirror data are deleted only by an Owner purge; there is no automatic retention purge.

## Routes

| Route | Owner | Behavior |
|---|---|---|
| `/` | application router | replace-navigate to `/workbooks` |
| `/workbooks` | Workbook Hub | Catalog, create, import, local/remote/space/trash views |
| `/workbooks/:unitId` | Spreadsheet editor | One keyed `WorkbookSession` |
| `/auth/callback` | AuthSession | OIDC callback then return to the original route |
| `/auth/silent-renew` | AuthSession | OIDC silent renewal callback |

## Pixel baseline

At Chrome 100% zoom and DPR=1, the target viewport is 1672 × 941. The image pixels override conflicting prose values: top bar 58px, left rail 168px, primary content x=218, right content edge approximately x=1635, template cards y=193–374, information banner y=400–468, action bar height 40px, and file rows approximately 45px. The tolerance is 1px for core geometry, 2px for icon bounds, and 0.5% screenshot diff excluding font antialiasing only.

## Acceptance matrix

`home-default`, `home-local-files`, `home-shared`, `home-empty`, `home-offline`, `home-syncing`, `home-search`, `home-row-menu`, and `backstage-active-workbook` all require a real data path, not a mocked UI branch.
