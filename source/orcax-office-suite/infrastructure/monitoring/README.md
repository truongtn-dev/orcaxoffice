# Monitoring

E01 writes structured JSON logs to stdout and stamps every API request with a correlation id. The header name is `x-correlation-id`.

Metrics, tracing, and alerting products are not selected in this epic. Adding one later must go through an adapter so the monolith does not take a hard dependency on a single vendor.

Logs must not contain passwords, access tokens, payment credentials, or Office document bodies. Request bodies and query strings are omitted from the request log.
