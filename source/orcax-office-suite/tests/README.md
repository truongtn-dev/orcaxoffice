# Cross-application tests

Unit tests live next to the package they cover. This directory is reserved for later suites:

- `integration/` for flows that cross module boundaries once those modules exist
- `e2e/` for portal and API journeys
- `compatibility/` for real Word and Excel hosts

E01 covers the audit schema with `apps/api/test/audit-schema.integration-spec.ts`. That test runs when PostgreSQL is available, including in CI.
