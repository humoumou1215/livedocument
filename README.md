# Live Document

**CLI-first local bridge to the Office document you are working on.** A zero-dependency Node.js runtime plus thin WPS / Microsoft Office task-pane adapters. It does **not** embed an AI agent, implement an alternative document model, or expose MCP.

> **Status: v0.1.0 engineering preview.** CLI, HTTP runtime, pairing, routing, mock WPS and mock Office.js execution have automated tests. **Real WPS and Microsoft Office installation / live editing have not been verified on actual desktop hosts.** This is not a production-ready Office add-in or general-purpose JavaScript sandbox.

## Development status / 开发进度

> **Last updated: 2026-11-13 · Current version: v0.1.0 · Stage: engineering preview (not production-ready).**
> 当前结论：**CLI + Runtime + Bridge 代码闭环已经建立，自动化验证覆盖模拟宿主；真实 WPS / Microsoft Office 读写尚未通过实机验收。** 表中的“已实现”指代码存在，不代表所有桌面宿主已可用。

### Progress checklist / 功能完成情况

状态说明：**✅ 已实现并有自动化验证**；**🟡 已提供实现/样例，但待真实宿主验证**；**⬜ 尚未实现或尚未达到发布要求**。

| 范围 | 状态 | 当前证据 / 尚需完成 |
|---|---|---|
| CLI（`start/serve/stop/status/pair/targets/exec`，代码/文件/stdin 输入） | ✅ | 已有 CLI 子进程与 HTTP 集成测试；JSON 输出、错误退出码 |
| 本地常驻 Runtime、Target Registry、路由、每 Target FIFO 执行 | ✅ | Node 测试覆盖配对、调度和断线/超时等核心路径；宿主环境仍需实测 |
| 一次性配对码、Runtime/Bridge Token、localhost 限制 | ✅ | 已实现基础认证和限制；**不等于**安全沙箱或完整写操作授权 |
| WPS Bridge（Writer / Spreadsheets / Presentation） | 🟡 | JS API 执行适配器与加载项面板已写好；真实 WPS 中能否加载、读写仍待验证 |
| `livedoc wps install/uninstall` | 🟡 | `publish.xml` 保留原条目、幂等、备份、dry-run 有测试；不同 WPS 安装版的真实注册流程未验收 |
| Microsoft Office Bridge（Excel / Word / PowerPoint） | 🟡 | Office.js 适配器与 sideload manifest 已提交；证书、WebView/CSP 与真实读写尚未验证 |
| 跨平台自动化检查 | ✅ | GitHub Actions 配置为 Windows / macOS / Linux 上运行 `npm run check`、`npm test`；**不包含**真实 Office 桌面验收 |
| 文档身份与误写保护 | ⬜ | 目前 Target 是配对时的执行上下文；文档切换/重命名后的执行前身份校验待实现 |
| 安全执行与可恢复性 | ⬜ | 缺少可靠的读/写区分、宿主侧写操作确认、强制终止、事务回滚/撤销；超时后不得盲目重试 |
| 可分发安装包与正式发布 | ⬜ | 暂无签名安装包、自动更新和经过实机测试的完整安装/卸载流程 |

**验证口径：** 通过 `npm run check && npm test` 代表静态语法检查和模拟执行路径通过；**不代表 WPS 或 Office 的生产可用性、安装兼容性或安全审核通过**。最新自动化状态以 [GitHub Actions](https://github.com/humoumou1215/livedocument/actions) 为准。

### Current blockers / 当前主要风险

1. **WPS 实机兼容：** Windows WPS 2016、较新 Windows WPS、UOS WPS、macOS WPS 的加载项注册、WebView JavaScript 特性和 JS API 行为可能不同；旧环境可能缺少 ESM / `async` / `fetch`。
2. **Office.js 动态代码执行：** Office WebView 的 CSP 可能禁止 `new Function`；不能以放宽安全设置代替适配，应先实测，再决定是否改为受控操作协议。
3. **文档指向正确性：** 配对时采集的文档名称可能过期。必须防止切换文档后把 Agent 命令写进错误文档。
4. **写入安全和超时：** 当前调用脚本没有沙箱；Runtime 超时并不保证宿主脚本已停止，也没有自动回滚。
5. **安装交付：** Microsoft Office 当前需要手动准备可信本地 HTTPS 证书和 sideload；WPS 安装器还需要各平台真实验证。

## Roadmap / 开发计划

计划按**验收门槛**推进，不承诺未经估算的发布日期；后一阶段以对应前置验收完成为条件。

| 优先级 / 里程碑 | 目标 | 交付事项 | 完成判定（Definition of Done） |
|---|---|---|---|
| **P0 · v0.1.x** | **WPS 端到端可用** | 实测 Windows WPS（优先较新版和 WPS 2016）、UOS WPS；修复加载项安装和旧 WebView 兼容；增加可重复的手工验收记录 | 在真实 WPS 表格中完成 **配对 → `targets` → 读取 → 修改 → 读取核验**；断开/重连、关闭文档和 Target 过期无误写 |
| **P1 · v0.2.x** | **Microsoft Office 可用** | 完成可信 HTTPS 与 sideload 指引；首先闭环 Excel，再验证 Word / PowerPoint；处理 CSP 和 `Office.run/sync` 限制 | 支持的真实 Office 宿主分别给出可复现的读取/写入验收结果；不支持的宿主明确标注且安全失败 |
| **P1 · v0.2.x** | **文档安全与正确性** | 执行前重新校验目标文档；写入授权/确认机制；完善取消与超时状态、错误分类、审计日志及敏感信息保护 | 文档切换不会静默误写；危险写操作需要明确授权；超时/断线结果呈现“未知是否完成”，不自动重放 |
| **P2 · v0.3.x** | **安装与可维护性** | 优化跨平台安装、配置诊断、卸载与清理；兼容性矩阵；离线环境说明；版本发布流程 | 干净环境可按说明安装/卸载、重启后恢复；Windows/macOS/Linux 的支持边界清晰可测 |
| **Later / 待评估** | **扩展能力** | 更多脚本示例、针对具体 Office API 的能力检测、Agent Skill 使用体验 | 不改变 CLI-first、Runtime 独立的核心设计；需求评审后再排期 |

### Next work items / 下一轮建议执行顺序

- [ ] **先做 WPS 实机 smoke test**：安装检查 → Ribbon/TaskPane → 配对 → 列出 Target → 读取选区 → 写入测试单元格 → 再读取核对，记录 WPS 版本、系统与失败日志。
- [ ] **补兼容性矩阵**：Windows WPS 2016 / 较新版 WPS / UOS WPS / macOS WPS；分别标记 *未测、通过、失败、暂不支持*，不得以 CI 代替实机结果。
- [ ] **修复发现的 Bridge 兼容问题**；在真实宿主验证文档关闭、切换、重命名、双开、多 Target。
- [ ] **再推进 Office.js Excel**：可信本地证书、加载项 sideload、只读测试、写入测试、CSP 失败路径。
- [ ] **最后完成发布前安全门槛**：写操作确认、目标身份复核、超时语义、日志脱敏与安装回滚。

### Project boundaries / 不在当前范围

- **不提供 MCP Server**；CLI 是唯一对 Agent/终端公开的操作入口。
- 不嵌入模型、Agent 会话管理、聊天 UI 或独立 Office 文档模型。
- 不支持未授权的远程调用、公共网络暴露；暂不承诺跨主机操作。
- 不宣称支持在不受信任脚本上安全执行；请仅在已授权的文档运行可信脚本。

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

## Implementation and known limitations

The status table above is the maintained progress tracker; the detailed execution and deployment limitations are described in [Architecture](docs/ARCHITECTURE.md), [Security](docs/SECURITY.md), [WPS setup](docs/WPS.md), and [Office setup](docs/OFFICE.md). In particular, neither host-side JavaScript cancellation nor real Office installation has been proven in this preview.

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
