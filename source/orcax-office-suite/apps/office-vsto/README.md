# VSTO adapters (E06 / E07)

Windows desktop adapter for legacy Office hosts. Office Object Model calls must run on the Office UI thread.

## Boundaries

- Share the **JSON shape** of rule results with Office.js via `@orcax/api-contracts` / `OrcaX.Office.Contracts.RuleResult`. Do not compile a .NET assembly for the Office.js adapter to load.
- Keep document and workbook bodies on the machine. API calls use the same endpoints as Office.js: `/add-in/word/*` and `/add-in/excel/*`.
- Degrade with a visible explanation when a host cannot perform a command.

## Current delivery

- `OrcaX.Office.Contracts` — C# records/interfaces mirroring HTTP contracts (scaffold for a future VSTO Ribbon/task pane project).
- Full Visual Studio VSTO designer project, installer, and COM host automation are intentionally not shipped until a Windows build agent is available. Office.js is the runnable MVP host adapter.

## Next implementation step

Create Word/Excel VSTO projects that:

1. Call `IAddInApi` over HTTPS with the user access token.
2. Build local snapshots and apply the same deterministic planners documented under `apps/office-js/src`.
3. Never upload document bodies on ordinary validate/apply/print flows.
