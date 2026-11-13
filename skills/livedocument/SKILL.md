---
name: livedocument
summary: Operate a user's currently open WPS or Office document via local CLI.
---

# Live Document CLI

Use only when the user explicitly authorizes document access. The user must launch WPS/Office bridge, run `livedoc pair` and enter its code.

1. `livedoc status` (or `livedoc start` if not running). On first WPS setup, run `livedoc wps install --dry-run`, then `livedoc wps install` and restart WPS (this edits/backups publish.xml).
2. `livedoc targets` to find an available active execution context. Never reuse stale IDs.
3. Identify `host` and `kind`; for WPS use native `Application` JS API; for Microsoft Office use `context` in the applicable `*.run` and `await context.sync()` for loaded properties.
4. Prefer `livedoc exec --target ID --file /path/to/script.js` to avoid shell quoting, or `--stdin`.
5. All code is an async function body. Use `return JSON_SERIALIZABLE_VALUE` to retrieve results.
6. Confirm with the user before any unexpected write. Never assume a timed-out script was cancelled; inspect the document before retrying.
7. Error JSON / nonzero process exit: report directly. A lost target requires new pairing.

Security: this is **not** a sandbox. Only execute trusted scripts in the document the user selected.
