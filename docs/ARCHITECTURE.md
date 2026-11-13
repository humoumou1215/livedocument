# Runtime and protocol

The runtime exists independently of CLI processes. The CLI reads `runtime.json` (PID, port, HTTPS flag, 256-bit CLI token) from a user-private state directory. It makes authenticated requests only to `127.0.0.1`. Browser bridges load scripts from the runtime's own origin and acquire a per-target 256-bit token after redeeming a one-time 8-digit code valid 120 seconds. Only one code redemption is accepted.

## Operations

- `GET /v1/status` — CLI bearer; process + health.
- `POST /v1/pair-code` — CLI bearer; creates temporary pairing code.
- `POST /v1/pair` — code, `host`, `kind`, `name`; gives `targetId`, bridge bearer.
- `GET /v1/targets` — CLI bearer; ephemeral execution contexts.
- `POST /v1/execute` — CLI bearer; target, JS body, timeout; long-running HTTP response.
- `POST /v1/bridge/next` — bridge bearer; long-polls for one job.
- `POST /v1/bridge/result` — bridge bearer; completes exact job ID.
- `POST /v1/bridge/bye` — bridge bearer; unregisters target.
- `POST /v1/stop` — CLI bearer; graceful shutdown.

An origin-checked browser may access only the same-origin local runtime. All state-changing API routes require `application/json` content; unauthorized requests get an error. The master CLI bearer is never sent to an add-in page.

## Invariants

- Target IDs are ephemeral, unpredictable and distinct from document paths.
- Each target can have only **one active execution** at a time, with FIFO queued work.
- A job result must match the active job ID and the paired bridge bearer.
- A target that times out during execution is disconnected, because host JS cannot be force-stopped. Never replay an uncertain write.
- A target that stops polling expires after 45 seconds (pruned on CLI operations).
- No remote binding, no webhook, no Agent state, no MCP surface.

## Limitations / future hardening

Live target metadata is captured at pair time; renaming or switching the active document within the same Office add-in can invalidate the human-readable name. A production version should monitor and revalidate the host's real document identity before each execution. Other missing items: trusted code signing, host-level write confirmation, per-document user scope, browser compatibility (legacy WPS), installer, automatic pairing revocation on pane unload, TLS provisioning, and ability to stop a host script. Do not claim the current preview has these protections.
