# Security model / unresolved risks

**Treat every paired `exec` as arbitrary code execution in an Office add-in webview.** There is no read-only classifier, JS sandbox, per-call consent or undo guarantee in v0.1. Any CLI process holding the local runtime state token can read/alter document content using APIs available to the paired add-in. A malicious script could also use network APIs exposed to that webview. Only pair and execute on a trusted personal machine.

Controls implemented: localhost-only bind; Host/Origin validation to reduce DNS rebinding/cross-origin attacks; CLI bearer held in a user-private state file (POSIX permissions only); one-time short-lived pairing; failed-pair throttling; per-target random bearer; bounded request/code/result sizes; FIFO execution; target expiry; disconnect on execution timeout.

These are **not** a sandbox. Particularly on Windows, default ACLs of the user's profile directory govern access to the token file and must be checked; 0600 on Windows isn't a full ACL guarantee. Do not expose the runtime via port forwarding, reverse proxy or public interface. Do not commit `.livedocument/runtime.json`, tokens, or TLS keys. Avoid executing downloaded code, unreviewed Agent-generated scripts or untrusted content.

If an execution times out, the host can still complete its changes. Reopen/re-pair and inspect state before any retry. On failure, do not assume idempotence or rollback.
