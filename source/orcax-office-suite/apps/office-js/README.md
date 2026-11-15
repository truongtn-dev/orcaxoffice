# Office.js adapters (E06 / E07)

Word and Excel task panes for modern Office hosts. Document/workbook bodies stay in the host. The API only receives bootstrap data and run metadata (summary counts / print settings), never file bodies.

## Layout

- `manifests/` — Word and Excel sideload manifests (HTTPS origin `https://localhost:3002` by default)
- `public/` — task pane HTML/CSS/JS
- `src/word` / `src/excel` — deterministic validate/print planners (Node-tested) and browser cores
- `dist/` — static output from `npm run build -w @orcax/office-js`

## Local sideload

1. Run API (`npm run dev:api`) and apply migrations.
2. Serve the add-in static files over HTTPS on port 3002 (Office requires HTTPS for sideload).
3. Sideload `manifests/word-manifest.xml` or `excel-manifest.xml` in Office.
4. Sign in from the task pane with the same account as the portal; activate a license that includes `word.core` / `excel.core`.

## Word (E06)

Ribbon opens the task pane. User selects an approved Word profile, validates locally, previews high-risk formatting changes, confirms apply, then posts a validation-run summary to the API.

## Excel (E07)

Task pane loads approved Excel print profiles, previews A4 settings, confirms apply, and attempts PDF export through host APIs when available.

## VSTO

Windows desktop adapter lives in `apps/office-vsto`. It consumes the same HTTP contracts; it does not share a .NET binary with this package.
