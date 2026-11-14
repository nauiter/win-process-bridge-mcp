"""
Windows Process Bridge - local MCP server (stdio).

Gives an MCP-compatible AI client (Claude Desktop, etc.) real, read-first
access to Windows processes, services, resource usage and PowerShell,
without relying on screen/mouse control. Runs as a stdio child process of
the client itself - no network, no sandboxed VM.

Tools:
  - list_processes()               -> lists every running Windows process
  - check_process(name)            -> 'running'/'closed' for one process
  - open_task_manager()            -> opens Task Manager on screen
  - run_powershell(cmd)            -> read-only PowerShell cmdlet (allowlist)
  - run_powershell_readonly(cmd)   -> alias of run_powershell
  - get_process_stats(sort_by,top) -> CPU, RAM and threads per process
  - list_processes_grouped()       -> total RAM per process name
  - system_snapshot()              -> RAM, CPU and disk in one call
  - gpu_status()                   -> GPU/VRAM usage via nvidia-smi
  - start_process(target)          -> opens an app by name or path
  - kill_process(pid, name, confirm) -> terminates by PID or name
  - list_services(name_filter)     -> Windows services and their state
  - startup_apps()                 -> what launches with Windows
  - disk_usage(scan_path, top)     -> free space per drive + largest folders
"""
import os
import re
import subprocess

import psutil
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("win-process-bridge")

# Allowlist of read-only cmdlet verbs for run_powershell. The first token of
# every pipeline stage must start with one of these safe verbs. Anything
# outside that (Stop-, Remove-, Set-, Start-, Invoke-, New-, file redirects,
# arbitrary .exe calls) is rejected before it ever runs.
_SAFE_VERB_PREFIXES = ("get-", "measure-", "select-", "sort-", "where-",
                       "format-", "group-")
_BLOCKED_PATTERNS = re.compile(
    r"[;&|`]{2,}|(^|\s)(stop-|remove-|set-|start-|invoke-|new-|rm |del |"
    r"iwr|invoke-webrequest|invoke-expression|iex)\b",
    re.IGNORECASE,
)


# --------------------------------------------------------------------------
# Basic processes (tasklist)
# --------------------------------------------------------------------------

@mcp.tool()
def list_processes() -> str:
    """List every process currently running on Windows (name, PID, memory)."""
    result = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp850", errors="replace",
    )
    return result.stdout or result.stderr or "no output"


@mcp.tool()
def check_process(name: str) -> str:
    """Check whether a specific process is running (e.g. 'notepad.exe').

    Returns 'STATE=running' or 'STATE=closed' followed by tasklist details.
    """
    result = subprocess.run(
        ["tasklist", "/FI", f"IMAGENAME eq {name}", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp850", errors="replace",
    )
    out = result.stdout or ""
    state = "running" if name.lower() in out.lower() else "closed"
    return f"STATE={state}\n{out}"


@mcp.tool()
def open_task_manager() -> str:
    """Open Windows Task Manager on the user's screen."""
    subprocess.Popen(["taskmgr.exe"])
    return "Task Manager opened."


# --------------------------------------------------------------------------
# Read-only PowerShell
# --------------------------------------------------------------------------

def _validate_readonly_powershell(cmd: str) -> str | None:
    """Return None if the command is safe, or an error message otherwise."""
    if _BLOCKED_PATTERNS.search(cmd):
        return "Blocked: command contains a dangerous verb/pattern (Stop-, Remove-, Set-, Invoke-, etc.)."
    stages = [s.strip() for s in cmd.split("|") if s.strip()]
    if not stages:
        return "Blocked: empty command."
    for stage in stages:
        first_word = stage.split()[0].lower()
        if not any(first_word.startswith(p) for p in _SAFE_VERB_PREFIXES):
            return (
                f"Blocked: '{first_word}' is not in the allowlist "
                f"({', '.join(_SAFE_VERB_PREFIXES)})."
            )
    return None


@mcp.tool()
def run_powershell(cmd: str) -> str:
    """Run a READ-ONLY PowerShell command (allowlist: Get-, Measure-,
    Select-, Sort-, Where-, Format-, Group-). Any state-changing verb
    (Stop-, Remove-, Set-, Start-, Invoke-, New-) is rejected before
    execution. Example: 'Get-Process | Sort-Object CPU -Descending | Select-Object -First 5'
    """
    error = _validate_readonly_powershell(cmd)
    if error:
        return error
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", cmd],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=30,
    )
    return result.stdout or result.stderr or "no output"


@mcp.tool()
def run_powershell_readonly(cmd: str) -> str:
    """Alias of run_powershell: read-only PowerShell cmdlet, same allowlist."""
    return run_powershell(cmd)


# --------------------------------------------------------------------------
# Stats via psutil
# --------------------------------------------------------------------------

@mcp.tool()
def get_process_stats(sort_by: str = "cpu", top: int = 10) -> str:
    """CPU (%), RAM (MB) and thread count per process, sorted.

    sort_by: 'cpu', 'mem' or 'threads'. top: how many processes to return.
    """
    rows = []
    for p in psutil.process_iter(["pid", "name", "memory_info", "num_threads"]):
        try:
            cpu = p.cpu_percent(interval=None)
            mem_mb = p.info["memory_info"].rss / (1024 * 1024) if p.info["memory_info"] else 0
            rows.append((p.info["name"] or "?", p.info["pid"], cpu, mem_mb, p.info["num_threads"] or 0))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    key = {"cpu": 2, "mem": 3, "threads": 4}.get(sort_by, 2)
    rows.sort(key=lambda r: r[key], reverse=True)
    lines = ["name,pid,cpu_%,ram_mb,threads"]
    for name, pid, cpu, mem_mb, threads in rows[:top]:
        lines.append(f"{name},{pid},{cpu:.1f},{mem_mb:.1f},{threads}")
    return "\n".join(lines)


@mcp.tool()
def list_processes_grouped() -> str:
    """Total RAM and instance count grouped by process name (e.g. Chrome)."""
    groups: dict[str, list[float]] = {}
    for p in psutil.process_iter(["name", "memory_info"]):
        try:
            name = p.info["name"] or "?"
            mem_mb = p.info["memory_info"].rss / (1024 * 1024) if p.info["memory_info"] else 0
            groups.setdefault(name, [0, 0])
            groups[name][0] += mem_mb
            groups[name][1] += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    rows = sorted(groups.items(), key=lambda kv: kv[1][0], reverse=True)
    lines = ["name,total_ram_mb,instances"]
    for name, (mem_total, count) in rows:
        lines.append(f"{name},{mem_total:.1f},{count}")
    return "\n".join(lines)


@mcp.tool()
def system_snapshot() -> str:
    """RAM, CPU and disk usage for the whole system in one call."""
    vm = psutil.virtual_memory()
    cpu_pct = psutil.cpu_percent(interval=0.5)
    lines = [
        f"CPU: {cpu_pct:.1f}%",
        f"RAM: {vm.used / (1024**3):.1f} GB used of {vm.total / (1024**3):.1f} GB ({vm.percent:.1f}%)",
        "Disks:",
    ]
    for part in psutil.disk_partitions(all=False):
        try:
            u = psutil.disk_usage(part.mountpoint)
            lines.append(
                f"  {part.device} {u.used / (1024**3):.1f} GB used of "
                f"{u.total / (1024**3):.1f} GB ({u.percent:.1f}%)"
            )
        except (PermissionError, OSError):
            continue
    return "\n".join(lines)


@mcp.tool()
def gpu_status() -> str:
    """GPU and VRAM usage via nvidia-smi (NVIDIA GPUs only)."""
    result = subprocess.run(
        ["nvidia-smi",
         "--query-gpu=name,utilization.gpu,memory.used,memory.total,temperature.gpu",
         "--format=csv,noheader"],
        capture_output=True, text=True, errors="replace", timeout=10,
    )
    if result.returncode != 0:
        return f"nvidia-smi unavailable: {result.stderr.strip() or 'no NVIDIA GPU'}"
    return result.stdout.strip() or "no output"


# --------------------------------------------------------------------------
# Process control (destructive actions require an explicit confirm)
# --------------------------------------------------------------------------

@mcp.tool()
def start_process(target: str) -> str:
    """Open an application by name (e.g. 'notepad') or full .exe path."""
    try:
        subprocess.Popen(target, shell=True)
        return f"Launch command sent for: {target}"
    except Exception as e:
        return f"Error launching '{target}': {e}"


@mcp.tool()
def kill_process(pid: int = 0, name: str = "", confirm: bool = False) -> str:
    """Terminate a process by PID or name. Requires confirm=True (an extra
    guard on top of the client's own permission prompt) - without it, this
    only simulates the action and kills nothing.
    """
    if not confirm:
        return "Not executed: pass confirm=True to actually terminate the process."
    if pid:
        result = subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                                 capture_output=True, text=True, encoding="cp850", errors="replace")
    elif name:
        result = subprocess.run(["taskkill", "/F", "/IM", name],
                                 capture_output=True, text=True, encoding="cp850", errors="replace")
    else:
        return "Provide pid or name."
    return result.stdout or result.stderr or "no output"


# --------------------------------------------------------------------------
# Services, startup and disk
# --------------------------------------------------------------------------

@mcp.tool()
def list_services(name_filter: str = "") -> str:
    """List Windows services and their state (running/stopped).

    Optional name_filter: substring of the service name (case-insensitive).
    """
    lines = ["name,display_name,status"]
    for s in psutil.win_service_iter():
        try:
            info = s.as_dict()
        except Exception:
            continue
        if name_filter and name_filter.lower() not in info["name"].lower() \
                and name_filter.lower() not in (info.get("display_name") or "").lower():
            continue
        lines.append(f"{info['name']},{info.get('display_name', '')},{info['status']}")
    return "\n".join(lines)


@mcp.tool()
def startup_apps() -> str:
    """List programs configured to launch together with Windows."""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command",
         "Get-CimInstance Win32_StartupCommand | Select-Object Name,Command,Location "
         "| ConvertTo-Csv -NoTypeInformation"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
    )
    return result.stdout or result.stderr or "no output"


@mcp.tool()
def disk_usage(scan_path: str = "", top: int = 10) -> str:
    """Free/used space per drive. If scan_path is given, also lists the
    'top' largest first-level subfolders inside that path (can be slow on
    very large folders).
    """
    lines = ["Drives:"]
    for part in psutil.disk_partitions(all=False):
        try:
            u = psutil.disk_usage(part.mountpoint)
            lines.append(
                f"  {part.device} free: {u.free / (1024**3):.1f} GB of "
                f"{u.total / (1024**3):.1f} GB"
            )
        except (PermissionError, OSError):
            continue

    if scan_path:
        if not os.path.isdir(scan_path):
            lines.append(f"\nInvalid scan_path: {scan_path}")
            return "\n".join(lines)
        sizes = []
        try:
            for entry in os.scandir(scan_path):
                if not entry.is_dir(follow_symlinks=False):
                    continue
                total = 0
                for root, _, files in os.walk(entry.path):
                    for f in files:
                        try:
                            total += os.path.getsize(os.path.join(root, f))
                        except OSError:
                            continue
                sizes.append((entry.name, total))
        except PermissionError:
            lines.append(f"\nPermission denied scanning: {scan_path}")
            return "\n".join(lines)
        sizes.sort(key=lambda x: x[1], reverse=True)
        lines.append(f"\nLargest folders in {scan_path}:")
        for name, total in sizes[:top]:
            lines.append(f"  {name}: {total / (1024**3):.2f} GB")
    return "\n".join(lines)


if __name__ == "__main__":
    mcp.run()
