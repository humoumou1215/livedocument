# Live Document

**CLI-first local bridge to the Office document you are working on.** A zero-dependency Node.js runtime plus thin WPS / Microsoft Office task-pane adapters. It does **not** embed an AI agent, implement an alternative document model, or expose MCP.

> **Status: v0.1.0 engineering preview.** CLI, HTTP runtime, pairing, routing, mock WPS and mock Office.js execution have automated tests. **Real WPS and Microsoft Office installation / live editing have not been verified on actual desktop hosts.** This is not a production-ready Office add-in or general-purpose JavaScript sandbox.

## Architecture

```text
Agent / terminal / shell script
           |
     livedoc CLI (short-lived)
           | authenticated local HTTP
     Runtime (persistent, 127.0.0.1)
       | target registry / FIFO dispatcher / timeout
       | pairing, per-target session token
       | long-poll requests (host-initiated)
       +---- WPS JS add-in pane ---> Application (WPS JS API)
       +---- Office task pane ----> Excel/Word/PowerPoint.run (Office.js)
```

The browser-side bridge initiates outbound requests to localhost; the CLI does not attempt to call into a closed Office host. Every target is an **execution context**, not a permanent document ID. Only a paired, currently alive bridge can execute work.

## Requirements

Node.js **20 or newer**. No npm dependencies, network access, model subscription, or MCP server is required to run the CLI and runtime. Office.js itself is loaded from Microsoft's hosted script by the Office bridge, so **Microsoft Office integration is not fully offline out of the box**.

```sh
git clone https://github.com/humoumou1215/livedocument.git
cd livedocument
npm test
npm link           # optional: makes `livedoc` available in PATH
livedoc start      # or `node src/cli.js start`
livedoc status
livedoc pair       # prints 8-digit code valid for 120 seconds
```

For WPS, run `livedoc wps install --dry-run` to inspect the changes, then `livedoc wps install` and restart WPS. The installer preserves other entries and backs up any existing `publish.xml`; `livedoc wps uninstall` removes only the entries it added.

Then open the appropriate host bridge **from within the Office application**, enter the code, and click Connect. See [WPS setup](docs/WPS.md) and [Office setup](docs/OFFICE.md). Merely opening the page in a normal browser will **not** give it WPS or Office API access.

```sh
livedoc targets
# {"ok":true,"targets":[{"id":"...","host":"wps","kind":"spreadsheet","name":"book.xlsx", ...}]}

# WPS JS API: target is obtained from livedoc targets
livedoc exec --target TARGET_ID --code 'return Application.Selection.Address'

# File based script avoids shell escaping; uses async function body
livedoc exec --target TARGET_ID --file ./samples/wps-selection.js

# Microsoft Office / Excel.js code uses `context` instead of `Application`:
livedoc exec --target OFFICE_TARGET_ID --code 'const r=context.workbook.getSelectedRange(); r.load("address"); await context.sync(); return r.address;'

livedoc stop
```

Use `--stdin` for piping JS from shell, and `--timeout 30000` to control execution time (1–120 seconds). All commands produce one JSON object on stdout and use nonzero exit status on errors. `livedoc serve` runs in the foreground for debugging; `livedoc start` starts a detached runtime.

## Execution contract

Scripts are the **body of an async function** executed inside the paired host. Use `return ...` to send a JSON-serializable result. WPS scripts receive `Application` (its native object model); Microsoft Office scripts receive `context` inside `Excel.run`, `Word.run`, or `PowerPoint.run`. Office.js reads may need `load` + `await context.sync()` before accessing properties. The Office wrapper calls `context.sync()` again on successful return. This is not a unified Office object model.

**Do not execute scripts from untrusted sources.** Dynamic execution is explicitly not isolated from the add-in environment, and current build does not classify reads vs writes. The 8-digit pairing code is a consent step, not an execution sandbox. A timed-out script can still finish in the host; the runtime disconnects its target rather than risking overlapping subsequent writes. See [security](docs/SECURITY.md).

## What's actually implemented

- Local HTTP server bound to `127.0.0.1` only; runtime auth token in a local state file; host/origin checking; code size and response size limits.
- Single-use 8-digit pairing codes, expiry and failed-pair rate limiting; a unique token per paired bridge.
- Live target discovery; one running command per target with FIFO scheduling; result transport, disconnect and timeout handling.
- CLI start, serve, stop, status, pair, targets, exec; background launch and JSON output.
- WPS ribbon stub and task pane; Office sideload manifest and task pane; host JS executors.
- Node-level unit and end-to-end HTTP/CLI tests with mocked Office applications.

## Not yet production complete

- No verified WPS 2016 / WPS 2026 / UOS / Microsoft Office WebView end-to-end host test. Some older embedded browser engines may not support JavaScript modules, async functions, or `fetch`.
- No signed binaries / auto-update; WPS `publish.xml` registration is automated but **not yet verified in real WPS hosts**. The install command edits the WPS configuration on the local machine (backing up any prior file).
- Office's manifest expects a locally trusted HTTPS certificate and a working WebView localhost loopback; this is manual (see docs). Office.js CDN is not mirrored.
- No hard cancellation of host-side JavaScript, per-command confirmation, security sandbox, transaction/undo semantics, or remote-network access.
- Multiple simultaneously active documents within one Office add-in context require independent pairing; metadata identifies the context on connection, not a permanent file identity.

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `LIVEDOC_PORT` | `18765` | Stable localhost port; WPS sample URL and Office manifest must match |
| `LIVEDOC_HOME` | `~/.livedocument` (or `$XDG_RUNTIME_DIR/livedocument`) | state directory |
| `LIVEDOC_TLS_CERT` | unset | Path to trusted HTTPS certificate (PEM) |
| `LIVEDOC_TLS_KEY` | unset | Corresponding HTTPS private key (PEM) |
| `LIVEDOC_WPS_ADDONS_DIR` | platform WPS path | Override WPS `publish.xml` location for custom builds/tests |

For Windows, protect the state directory with your user account's filesystem permissions; POSIX builds set `0700` on the directory and `0600` on the state file. The runtime state token permits full execution requests; do not share it.

## Development

```sh
npm run check
npm test
```

Zero external dependencies. CI runs Node's built-in test runner on Linux/Windows/macOS. The [agent guide](skills/livedocument/SKILL.md) has compact command usage. [Architecture](docs/ARCHITECTURE.md) documents transport and invariants.

License: [MIT](LICENSE).
