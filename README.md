# Windows Process Bridge MCP

> A local MCP (Model Context Protocol) server that gives AI clients like Claude Desktop real, read-first access to Windows processes, services, resources and PowerShell, without controlling the mouse or screen.

Most AI desktop assistants either run in a sandboxed cloud VM (no real OS access) or fall back to screen/mouse control to check something as simple as "is this app running". This server exposes that information directly, as proper tool calls, with a safety-first design: read operations are unrestricted, state-changing operations require an explicit `confirm=True` plus the client's own permission prompt, and arbitrary PowerShell is blocked by an allowlist of read-only verbs (`Get-`, `Measure-`, `Select-`, `Sort-`, `Where-`, `Format-`, `Group-`). Any attempt to run `Stop-`, `Remove-`, `Set-`, `Start-`, `Invoke-`, `New-`, or to pipe into a file/exe, is rejected before it ever executes.

## 🧰 What it gives your AI client

- `list_processes` / `check_process(name)` / `open_task_manager` — basic process visibility, no mouse needed
- `run_powershell(cmd)` / `run_powershell_readonly(cmd)` — any read-only PowerShell one-liner or pipeline, allowlist-enforced
- `get_process_stats`, `list_processes_grouped`, `system_snapshot`, `gpu_status` — CPU/RAM/VRAM/disk at a glance, via `psutil` and `nvidia-smi`
- `start_process(target)` / `kill_process(pid, name, confirm=True)` — launch or terminate, with destructive actions gated behind an explicit confirm flag
- `list_services`, `startup_apps`, `disk_usage` — Windows services, startup programs and largest folders

## 🚀 Tech Stack

- Python 3.11+
- [`mcp`](https://pypi.org/project/mcp/) (official Model Context Protocol SDK, `FastMCP`)
- [`psutil`](https://pypi.org/project/psutil/) for cross-process stats
- Stdlib `subprocess` for `tasklist`/`taskkill`/`powershell`/`nvidia-smi`

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
    "win-process-bridge": {
      "command": "C:\\path\\to\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\win-process-bridge-mcp\\server.py"]
    }
  }
}
```

Restart Claude Desktop completely (quit from the system tray, not just close the window) for the new MCP server to load. On first use of any tool, the client shows its own permission prompt before running anything.

## 📄 Project Structure

```
server.py            All 13 tools in a single FastMCP server (stdio transport)
requirements.txt      mcp + psutil
```

## 🔒 Security Model

| Aspect | Design |
|---|---|
| Read operations | Unrestricted (`list_processes`, `system_snapshot`, `gpu_status`, etc.) |
| PowerShell | Allowlist of read-only verbs only; blocklist regex as a second layer |
| Destructive ops | `kill_process` requires `confirm=True` in addition to the client's own permission UI |
| Scope | Only processes/services explicitly named by the calling tool call; no background daemons, no network listener |
| Transport | stdio only — runs as a child process of the AI client, no open port |

This project does not attempt to sandbox PowerShell execution itself (no AppContainer, no restricted token) — the allowlist is a cooperative guard against accidental or careless use by an AI client that is already supposed to behave, not a security boundary against a malicious one. Don't expose this server to an untrusted caller.

## License

MIT. See [LICENSE](LICENSE).

---

© 2026 Nauiter Master | AI Strategist, Digital Artist & Automation
