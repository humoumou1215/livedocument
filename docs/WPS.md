# WPS host setup (engineering preview)

1. Install Node.js 20+, run `livedoc start`; browse to `http://127.0.0.1:18765/wps-addon/ribbon.xml` to verify the runtime serves the WPS add-in assets.
2. The development add-in is exposed at `http://127.0.0.1:18765/wps-addon/` (or the matching HTTPS URL when TLS is enabled). The sample `ribbon.xml` adds a **Live Document → Connect CLI** ribbon control, and `main.js` opens the `/bridge/wps.html` task pane.
3. Register this URL as an **online WPS JS add-in** using your WPS-supported development/publishing flow (e.g. `wpsjs` and official `publish.xml` install workflow). **No one-click installer is included yet.** Different WPS builds support different installation modes. The legacy `jsplugins.xml` mode is restricted on newer builds; prefer `publish.xml` where available.
4. In WPS, open the desired document and click **Connect CLI**. In the terminal run `livedoc pair`, enter the returned code in the WPS task pane, and click Connect.
5. Run `livedoc targets`, then a simple read: `livedoc exec --target TARGET --code 'return Application.Selection.Address'` (spreadsheet only).

If the ribbon does not appear, use the WPS add-in debugger and verify the main.js/ribbon.xml resources. If the panel reports `Application unavailable`, it was not opened in a supported WPS add-in execution context. **The WPS installation and real host execution path have not been verified on WPS desktop in CI**. Older WPS 2016 JS add-in browser environments may lack the ESM/async/fetch features this preview uses.

Official documentation:
- https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wps-addin-development-instructions
- https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/addin-api/Application/obj
