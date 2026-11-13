# Microsoft Office host setup (engineering preview)

The sample `adapters/office/manifest.xml` declares Excel, Word, and PowerPoint task-pane contexts. It points at **https://localhost:18765**. A trusted local HTTPS certificate is required. This is not configured automatically.

1. Install `mkcert` (external developer utility), run `mkcert -install`, then issue a cert covering both `localhost` and `127.0.0.1`: `mkcert -cert-file local-cert.pem -key-file local-key.pem localhost 127.0.0.1`. Keep the private key private.
2. Set `LIVEDOC_TLS_CERT` and `LIVEDOC_TLS_KEY` to their full file paths; `livedoc serve` or `livedoc start` starts an HTTPS runtime. The CLI validates the local HTTPS certificate chain, so it must be trusted by Node/OS.
3. In Microsoft Office's supported add-in sideload flow, load `adapters/office/manifest.xml`; open its **Live Document CLI Bridge** pane. Some Office/WebView versions may require a localhost loopback exemption and may not support eval dynamically. This page loads `office.js` from the Microsoft CDN and therefore requires access to that script.
4. Run `livedoc pair`; enter the one-time code in the Office pane, click Connect. `livedoc targets` lists the new Office execution context.
5. Excel example:

```sh
livedoc exec --target TARGET --code 'const r=context.workbook.getSelectedRange();r.load("address");await context.sync();return r.address'
```

Word and PowerPoint run adapters are included but **are unverified in real Office environments**. The bridge executes a JS function body inside the native `Excel.run`, `Word.run`, or `PowerPoint.run` callback, and automatically invokes `context.sync()` after it returns. There is no cross-host API compatibility layer.

The manifest and runtime API serve as development scaffolding, not a signed or published Microsoft Marketplace add-in. If your Office runtime's CSP forbids `new Function`, dynamic code execution is unsupported there and requires a new host-specific, structured-operation adapter rather than weakening system security settings.

References:
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/add-in-manifests
- https://learn.microsoft.com/en-us/office/dev/add-ins/develop/application-specific-api-model
