# putinha-faz-tudo (formerly Windows Process Bridge MCP)

> A local MCP (Model Context Protocol) server that gives an AI client like Claude Desktop real access to Windows processes, services, files, clipboard and an upload server — without controlling the mouse or screen, and without a VM sandbox.

Most AI desktop assistants either run in a sandboxed cloud VM (no real OS access) or fall back to screen/mouse control to check something as simple as "is this app running". This server exposes that directly, as proper tool calls, with a safety-first design: read operations are unrestricted, state-changing operations require an explicit `confirm=True` plus the client's own permission prompt, and arbitrary PowerShell is blocked by an allowlist of read-only verbs (`Get-`, `Measure-`, `Select-`, `Sort-`, `Where-`, `Format-`, `Group-`). Any attempt to run `Stop-`, `Remove-`, `Set-`, `Start-`, `Invoke-`, `New-`, or to redirect into a file/exe, is rejected before it ever executes.

## 🧰 What it gives your AI client

**Processes & system** (original batch)
- `list_processes` / `check_process(name)` / `open_task_manager` — basic process visibility, no mouse needed
- `run_powershell(cmd)` / `run_powershell_readonly(cmd)` — any read-only PowerShell one-liner or pipeline, allowlist-enforced
- `get_process_stats`, `list_processes_grouped`, `system_snapshot`, `gpu_status` — CPU/RAM/VRAM/disk at a glance, via `psutil` and `nvidia-smi`
- `start_process(target)` / `kill_process(pid, name, confirm=True)` — launch or terminate, with destructive actions gated behind an explicit confirm flag
- `list_services`, `startup_apps`, `disk_usage` — Windows services, startup programs and largest folders

**Files** (save straight to disk, no browser "Save As" popup)
- `save_url_to_file(url, dest_path, overwrite=False)` — downloads http(s) straight to disk
- `save_base64_to_file(data_b64, dest_path, overwrite=False)` — writes base64 (e.g. extracted from a `blob:` URL via page JS)
- `move_file(src, dest, overwrite=False)` / `copy_file(src, dest, overwrite=False)`
- `ensure_dir(path)` / `list_dir(path, pattern="", top=50)` / `list_recent_downloads(minutes=10, pattern="")`
- `delete_file(path, confirm=True)` — sends to the Recycle Bin via `SHFileOperationW` (reversible, not a permanent delete)
- `file_info(path)` — size, dates, sha256
- `open_in_explorer(path)` — opens Explorer with the item pre-selected

**Clipboard**
- `save_clipboard_image(dest_path, overwrite=False)` — saves as `.bmp` whatever image is on the clipboard (e.g. after copying an image in the browser)
- `get_clipboard_text()` / `set_clipboard_text(text)`

**Upload server** (for large files/video, without the content going through chat)
- `start_upload_server(dest_dir, allowed_origins="https://labs.google,https://flow.google.com", port=0, max_mb=500)` — starts an `http.server.ThreadingHTTPServer` on `127.0.0.1` (ephemeral port), CORS restricted to an origin allowlist plus a random token required in the `X-Upload-Token` header. The page does `fetch(blob)` POST straight to `/upload?name=...`; the bytes go to disk without passing through the AI client's chat context.
- `stop_upload_server()` / `upload_server_status()`

## 🚫 Known limitations (tested live, not guesses)

### CDP / `--remote-debugging-port` on Claude Desktop: blocked on purpose

Claude Desktop is a full-trust MSIX app (Desktop Bridge, not an AppContainer/UWP sandbox — user data lives in a normal folder, not the virtualized `LocalState`). But the main process **refuses to start** if the command line contains any of 16 debug/network-override switches (`--remote-debugging-port`, `--inspect`, etc.):

```
Claude: refusing to start — a debugging or network-override switch is present on the command line.
```

It only unlocks with a signed `CLAUDE_CDP_AUTH` token (ed25519, key compiled into the app) — only Anthropic holds the private key. Independently confirmed by others:
- [`bigfnj/desktop-ai-companion`](https://github.com/bigfnj/desktop-ai-companion/blob/master/docs/IDEAS.md) — same discovery via `app.asar` reverse engineering.
- [`RemyMachado/claude-desktop-theme`](https://github.com/RemyMachado/claude-desktop-theme) — lists the 16 switches.
- [`anthropics/claude-code#89160`](https://github.com/anthropics/claude-code/issues/89160) and [`#92044`](https://github.com/anthropics/claude-code/issues/92044) — devs debugging other issues hit the same block.

This isn't a bug, it's a deliberate protection against session/token extraction via DevTools. No future app update will "unlock" this — not worth trying again.

### Claude Desktop's embedded browser pane has its own quirks

- `save_url_to_file` only works with a public http(s) URL that doesn't require a session cookie — fails with HTTP 400 on URLs that need auth (e.g. Google Flow).
- `save_base64_to_file` works, but the base64 goes through the chat context — only viable for small files (tested OK at ~200 KB).
- `save_clipboard_image` **does not work** inside Claude Desktop's embedded browser pane: the page gets `Write permission denied` when trying to write to the clipboard, and right-click doesn't open a native menu there (unlike a real Chrome window).
- `start_upload_server` works well for a full-size **image** (tested: 204 KB, HTTP 200, no popup at all). For **video** it fails: `fetch` on the `<video>` element gives "Failed to fetch", `captureStream` gives `SecurityError: cross-origin data` — the video host (e.g. Flow) serves it through a path the page's JS can't read. That's a limit on the site's side, not the MCP's; the remaining path is a native download (accepting the popup) plus `list_recent_downloads` + `move_file` to organize it afterward.

## 🚀 Tech Stack

- Python 3.11+
- [`mcp`](https://pypi.org/project/mcp/) (official Model Context Protocol SDK, `FastMCP`)
- [`psutil`](https://pypi.org/project/psutil/) for cross-process stats
- [`pywin32`](https://pypi.org/project/pywin32/) for clipboard and Recycle Bin (`win32clipboard`, `SHFileOperationW`)
- Stdlib `subprocess`, `http.server`, `urllib.request`

## 🛠️ Installation & Local Development

### Prerequisites
- Windows 10/11
- Python 3.11+

### Steps

```
git clone <repository-url>
cd win-process-bridge-mcp

python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python server.py   # runs as a stdio MCP server, waiting for a client
```

### Register it in Claude Desktop

Edit `claude_desktop_config.json` (Settings → Developer → Edit Config) and add:

```json
{
  "mcpServers": {
    "putinha-faz-tudo": {
      "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\win-process-bridge-mcp\\server.py"]
    }
  }
}
```

Restart Claude Desktop completely (quit from the system tray, not just close the window) for the new MCP server to load. On first use of any tool, the client shows its own permission prompt before running anything.

## 📄 Project Structure

```
server.py            All 30 tools in a single FastMCP server (stdio transport)
requirements.txt      mcp + psutil + pywin32
```

## 🔒 Security Model

| Aspect | Design |
|---|---|
| Read operations | Unrestricted (`list_processes`, `system_snapshot`, `gpu_status`, etc.) |
| PowerShell | Allowlist of read-only verbs only; blocklist regex as a second layer |
| Destructive ops | `kill_process` and `delete_file` require `confirm=True` in addition to the client's own permission UI |
| Upload server | Only listens on `127.0.0.1` (never `0.0.0.0`), CORS restricted to an explicit origin allowlist, and requires a random per-session token in the header — even an allowed origin can't upload without the right token |
| Scope | Only processes/services/files explicitly named by the calling tool call; no persistent background daemon besides the optional upload server, which the client itself starts and stops |
| Transport | stdio as a child process of the AI client; the only network listener is the optional upload server, and even that is ephemeral and local-only |

This project does not attempt to sandbox PowerShell execution itself (no AppContainer, no restricted token) — the allowlist is a cooperative guard against accidental or careless use by an AI client that is already supposed to behave, not a security boundary against a malicious one. Don't expose this server to an untrusted caller.

## License

MIT. See [LICENSE](LICENSE).

---

© 2026 Nauiter Master | AI Strategist, Digital Artist & Automation
