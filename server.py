"""
Windows Process Bridge - local MCP server (stdio).

Used by Claude Desktop's mcpServers (Settings > Developer > Local MCP
Servers). Runs as a child process of the client itself - no network,
no sandboxed VM (except the optional local upload server below, which
only listens on 127.0.0.1).

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
  - save_url_to_file(url, dest)    -> downloads http(s) straight to disk, no popup
  - save_base64_to_file(b64, dest) -> writes base64 (e.g. extracted from a blob:
                                       URL via page JS) to disk
  - move_file(src, dest)           -> move/rename (e.g. Downloads -> project folder)
  - copy_file(src, dest)           -> copy while keeping the original
  - list_recent_downloads(min)     -> find what just landed in Downloads
  - ensure_dir(path)               -> create a folder (and parents) if missing
  - list_dir(path, pattern, top)   -> list any folder, not just Downloads
  - delete_file(path, confirm)     -> sends to the Recycle Bin (reversible), needs confirm
  - file_info(path)                -> size, dates and sha256 of a file
  - open_in_explorer(path)         -> opens Explorer with the file pre-selected
  - save_clipboard_image(dest)     -> saves an image copied (Ctrl+C) from the clipboard
  - get_clipboard_text()           -> reads the current Windows clipboard text
  - set_clipboard_text(text)       -> writes text to the Windows clipboard
  - start_upload_server(dest_dir)  -> local server so a web page can send bytes via
                                       fetch POST, without the content going through chat
  - stop_upload_server()           -> tears down the upload server
  - upload_server_status()         -> status and uploads received this session
"""
import base64
import ctypes
import hashlib
import http.server
import json
import os
import re
import secrets
import shutil
import struct
import subprocess
import threading
import time
import urllib.request
from urllib.parse import parse_qs, urlparse

import psutil
import win32clipboard
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("putinha-faz-tudo")

# Allowlist of read-only cmdlets for run_powershell. The first token of each
# pipeline stage must start with one of these safe verbs. Anything else
# (Stop-, Remove-, Set-, Start-, Invoke-, New-, redirecting to a file,
# calling an arbitrary .exe) is rejected.
_SAFE_VERB_PREFIXES = ("get-", "measure-", "select-", "sort-", "where-",
                       "format-", "group-")
_BLOCKED_PATTERNS = re.compile(
    r"[;&|`]{2,}|(^|\s)(stop-|remove-|set-|start-|invoke-|new-|rm |del |"
    r"iwr|invoke-webrequest|invoke-expression|iex)\b",
    re.IGNORECASE,
)

# Extension: allows, and only this, the ccusage pipeline.
#  - 1st stage can be exactly "npx ccusage@latest [--flags]"
#  - a "ConvertFrom-Json" stage with no arguments
#  - a "ForEach-Object" stage with only { } blocks containing variables,
#    numbers and arithmetic operators (no commands, calls or strings)
_NPX_RE = re.compile(r"^npx\s+ccusage@latest(\s+-{0,2}[A-Za-z][\w\-]*)*$", re.IGNORECASE)
_CCUSAGE_CMD = r"D:\hermes-work\tools\ccusage\node_modules\.bin\ccusage.cmd"
_CCUSAGE_RE = re.compile(r"^ccusage(\s+-{0,2}[A-Za-z][\w\-]*)*$", re.IGNORECASE)
_FE_BODY_RE = re.compile(r"\{([^{}]*)\}")
_FE_FULL_RE = re.compile(
    r"(?:\s+(?:-(?:begin|process|end))?\s*\{[^{}]*\})+\s*$", re.IGNORECASE)
_FE_TOKEN_RE = re.compile(r"\$\w+(?:\.\w+)*|\d+(?:\.\d+)?|[\s+\-*/=;]")


def _foreach_ok(stage: str) -> bool:
    rest = stage[len("foreach-object"):]
    if not _FE_FULL_RE.match(rest):
        return False
    return all(_FE_TOKEN_RE.sub("", body) == ""
               for body in _FE_BODY_RE.findall(rest))


# --------------------------------------------------------------------------
# Basic processes (tasklist)
# --------------------------------------------------------------------------

@mcp.tool()
def list_processes() -> str:
    """List every Windows process running right now (name, PID, memory)."""
    result = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp850", errors="replace",
    )
    return result.stdout or result.stderr or "no output"


@mcp.tool()
def check_process(name: str) -> str:
    """Check whether a specific process is running (e.g. 'CapCut.exe').

    Returns 'STATE=running' or 'STATE=closed' followed by tasklist's details.
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
    """Opens the Windows Task Manager on the user's screen."""
    subprocess.Popen(["taskmgr.exe"])
    return "Task Manager opened."


# --------------------------------------------------------------------------
# Read-only PowerShell
# --------------------------------------------------------------------------

def _validate_readonly_powershell(cmd: str) -> str | None:
    """Returns None if the command is safe, or an error message."""
    if _BLOCKED_PATTERNS.search(cmd):
        return "Blocked: command contains a dangerous verb/pattern (Stop-, Remove-, Set-, Invoke-, etc.)."
    stages = [s.strip() for s in cmd.split("|") if s.strip()]
    if not stages:
        return "Blocked: empty command."
    for i, stage in enumerate(stages):
        first_word = stage.split()[0].lower()
        if i == 0 and (_NPX_RE.match(stage) or _CCUSAGE_RE.match(stage)):
            continue
        if first_word == "convertfrom-json" and len(stage.split()) == 1:
            continue
        if first_word == "foreach-object":
            if _foreach_ok(stage):
                continue
            return ("Blocked: ForEach-Object only accepts blocks with variables, "
                    "numbers and arithmetic operators.")
        if not any(first_word.startswith(p) for p in _SAFE_VERB_PREFIXES):
            return (
                f"Blocked: '{first_word}' is not in the allowlist "
                f"({', '.join(_SAFE_VERB_PREFIXES)})."
            )
    return None


@mcp.tool()
def run_powershell(cmd: str) -> str:
    """Runs a READ-ONLY PowerShell command (allowlist: Get-, Measure-,
    Select-, Sort-, Where-, Format-, Group-). Any verb that changes state
    (Stop-, Remove-, Set-, Start-, Invoke-, New-) is rejected before it
    runs. Ex: 'Get-Process | Sort-Object CPU -Descending | Select-Object -First 5'
    """
    error = _validate_readonly_powershell(cmd)
    if error:
        return error
    # npx.ps1 is blocked by ExecutionPolicy; the npx.cmd shim isn't a PowerShell script.
    exec_cmd = re.sub(r"^\s*npx(\s)", r"npx.cmd\1", cmd, count=1, flags=re.IGNORECASE)
    exec_cmd = re.sub(r"^\s*ccusage(\s|$)", lambda m: _CCUSAGE_CMD + m.group(1), exec_cmd,
                      count=1, flags=re.IGNORECASE)
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", exec_cmd],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120 if (_NPX_RE.match(cmd.split("|")[0].strip()) or _CCUSAGE_RE.match(cmd.split("|")[0].strip())) else 30,
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
    """CPU (%), RAM (MB) and threads per process, sorted.

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
    """Total RAM and instance count per process name (e.g. Chrome)."""
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
    """RAM, CPU and disk of the system in a single call."""
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
    """GPU and VRAM usage via nvidia-smi (NVIDIA only)."""
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
# Process control (requires explicit confirm for the destructive part)
# --------------------------------------------------------------------------

@mcp.tool()
def start_process(target: str) -> str:
    """Opens an application by name (e.g. 'notepad') or full .exe path."""
    try:
        subprocess.Popen(target, shell=True)
        return f"Launch command sent for: {target}"
    except Exception as e:
        return f"Error opening '{target}': {e}"


@mcp.tool()
def kill_process(pid: int = 0, name: str = "", confirm: bool = False) -> str:
    """Terminates a process by PID or name. Requires confirm=True (extra
    defense beyond Claude's own permission screen) - without it, it only
    simulates and kills nothing.
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
    """Lists Windows services and their state (running/stopped).

    name_filter is optional: substring of the service name (case-insensitive).
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
    """Lists programs configured to launch together with Windows."""
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
    'top' largest first-level subfolders inside that path (can take a
    while on very large folders).
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
            lines.append(f"\ninvalid scan_path: {scan_path}")
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
            lines.append(f"\nNo permission to scan: {scan_path}")
            return "\n".join(lines)
        sizes.sort(key=lambda x: x[1], reverse=True)
        lines.append(f"\nLargest folders in {scan_path}:")
        for name, total in sizes[:top]:
            lines.append(f"  {name}: {total / (1024**3):.2f} GB")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# Files: save straight to disk without the browser's "Save As" popup
# --------------------------------------------------------------------------

def _safe_dest(dest_path: str, overwrite: bool) -> str | None:
    """Returns an error message, or None if dest_path can be written to."""
    parent = os.path.dirname(dest_path) or "."
    if not os.path.isdir(parent):
        return f"Folder does not exist: {parent}"
    if os.path.exists(dest_path) and not overwrite:
        return f"Already exists: {dest_path} (pass overwrite=True to replace)."
    return None


@mcp.tool()
def save_url_to_file(url: str, dest_path: str, overwrite: bool = False) -> str:
    """Downloads an http(s) URL straight to dest_path, without opening the
    browser's download dialog.

    Does not work with blob: URLs (those only exist inside the browser tab;
    use save_base64_to_file for that case, extracting the base64 via page JS).
    """
    if not re.match(r"^https?://", url, re.IGNORECASE):
        return "Blocked: only http:// or https:// URLs accepted (blob: can't be downloaded here)."
    error = _safe_dest(dest_path, overwrite)
    if error:
        return error
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest_path, "wb") as f:
            f.write(resp.read())
        size = os.path.getsize(dest_path)
        return f"Saved: {dest_path} ({size / 1024:.1f} KB)"
    except Exception as e:
        return f"Error downloading '{url}': {e}"


@mcp.tool()
def save_base64_to_file(data_b64: str, dest_path: str, overwrite: bool = False) -> str:
    """Decodes base64 (e.g. extracted from a blob: URL via page JS, without the
    data:URI prefix) and writes it to dest_path. Use this when the source is
    a blob: URL in the browser.
    """
    error = _safe_dest(dest_path, overwrite)
    if error:
        return error
    if data_b64.startswith("data:"):
        data_b64 = data_b64.split(",", 1)[-1]
    try:
        raw = base64.b64decode(data_b64)
        with open(dest_path, "wb") as f:
            f.write(raw)
        return f"Saved: {dest_path} ({len(raw) / 1024:.1f} KB)"
    except Exception as e:
        return f"Error decoding/writing '{dest_path}': {e}"


@mcp.tool()
def move_file(src: str, dest: str, overwrite: bool = False) -> str:
    """Moves/renames a file (e.g. from what just landed in Downloads to the
    right project folder). Fails if dest already exists and overwrite=False.
    """
    if not os.path.isfile(src):
        return f"Not found: {src}"
    error = _safe_dest(dest, overwrite)
    if error:
        return error
    try:
        if overwrite and os.path.exists(dest):
            os.remove(dest)
        os.replace(src, dest)
        return f"Moved: {src} -> {dest}"
    except Exception as e:
        return f"Error moving '{src}' -> '{dest}': {e}"


@mcp.tool()
def copy_file(src: str, dest: str, overwrite: bool = False) -> str:
    """Copies a file while keeping the original (e.g. duplicate before editing)."""
    if not os.path.isfile(src):
        return f"Not found: {src}"
    error = _safe_dest(dest, overwrite)
    if error:
        return error
    try:
        shutil.copy2(src, dest)
        return f"Copied: {src} -> {dest}"
    except Exception as e:
        return f"Error copying '{src}' -> '{dest}': {e}"


@mcp.tool()
def ensure_dir(path: str) -> str:
    """Creates the folder (and any needed parent folders) if it doesn't exist yet."""
    try:
        os.makedirs(path, exist_ok=True)
        return f"Folder ready: {path}"
    except Exception as e:
        return f"Error creating '{path}': {e}"


@mcp.tool()
def list_recent_downloads(minutes: int = 10, pattern: str = "") -> str:
    """Lists files created/modified in Downloads within the last N minutes
    (to find what the browser's download popup just saved, so it can be
    moved with move_file). pattern is optional: substring of the name
    (case-insensitive).
    """
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    if not os.path.isdir(downloads):
        return f"Folder not found: {downloads}"
    cutoff = time.time() - minutes * 60
    rows = []
    for name in os.listdir(downloads):
        path = os.path.join(downloads, name)
        if not os.path.isfile(path):
            continue
        if pattern and pattern.lower() not in name.lower():
            continue
        mtime = os.path.getmtime(path)
        if mtime < cutoff:
            continue
        rows.append((mtime, name, path, os.path.getsize(path)))
    if not rows:
        return f"Nothing in {downloads} in the last {minutes} min (filter: '{pattern}')."
    rows.sort(reverse=True)
    lines = [f"Recent downloads (last {minutes} min):"]
    for mtime, name, path, size in rows:
        lines.append(f"  {name}  ({size / 1024:.1f} KB)  {path}")
    return "\n".join(lines)


@mcp.tool()
def list_dir(path: str, pattern: str = "", top: int = 50) -> str:
    """Lists files and folders inside 'path' (non-recursive).

    pattern is optional: substring of the name (case-insensitive). top: line limit.
    """
    if not os.path.isdir(path):
        return f"Folder not found: {path}"
    rows = []
    try:
        for entry in os.scandir(path):
            if pattern and pattern.lower() not in entry.name.lower():
                continue
            try:
                st = entry.stat()
                kind = "DIR" if entry.is_dir() else "FILE"
                rows.append((st.st_mtime, kind, entry.name, st.st_size))
            except OSError:
                continue
    except PermissionError:
        return f"No permission to list: {path}"
    rows.sort(reverse=True)
    lines = [f"{path} ({len(rows)} items, most recent first):"]
    for mtime, kind, name, size in rows[:top]:
        size_s = f"{size / 1024:.1f} KB" if kind == "FILE" else "-"
        lines.append(f"  [{kind}] {name}  {size_s}")
    return "\n".join(lines)


@mcp.tool()
def delete_file(path: str, confirm: bool = False) -> str:
    """Sends a file to the Windows Recycle Bin (reversible, not a permanent delete).

    Requires confirm=True (extra defense beyond Claude's own permission screen).
    """
    if not confirm:
        return "Not executed: pass confirm=True to actually send it to the Recycle Bin."
    if not os.path.isfile(path):
        return f"Not found: {path}"
    try:
        FO_DELETE = 0x0003
        FOF_ALLOWUNDO = 0x0040
        FOF_NOCONFIRMATION = 0x0010
        FOF_SILENT = 0x0004

        class SHFILEOPSTRUCTW(ctypes.Structure):
            _fields_ = [
                ("hwnd", ctypes.c_void_p),
                ("wFunc", ctypes.c_uint),
                ("pFrom", ctypes.c_wchar_p),
                ("pTo", ctypes.c_wchar_p),
                ("fFlags", ctypes.c_uint16),
                ("fAnyOperationsAborted", ctypes.c_int),
                ("hNameMappings", ctypes.c_void_p),
                ("lpszProgressTitle", ctypes.c_wchar_p),
            ]

        op = SHFILEOPSTRUCTW()
        op.wFunc = FO_DELETE
        op.pFrom = os.path.abspath(path) + "\0\0"
        op.fFlags = FOF_ALLOWUNDO | FOF_NOCONFIRMATION | FOF_SILENT
        ret = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(op))
        if ret != 0 or op.fAnyOperationsAborted:
            return f"Failed to send to Recycle Bin (code {ret})."
        return f"Sent to Recycle Bin: {path}"
    except Exception as e:
        return f"Error deleting '{path}': {e}"


@mcp.tool()
def file_info(path: str) -> str:
    """Size, creation/modification date and sha256 of a file."""
    if not os.path.isfile(path):
        return f"Not found: {path}"
    try:
        st = os.stat(path)
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return (
            f"Path: {path}\n"
            f"Size: {st.st_size / 1024:.1f} KB\n"
            f"Modified: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.st_mtime))}\n"
            f"Created: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.st_ctime))}\n"
            f"SHA256: {h.hexdigest()}"
        )
    except Exception as e:
        return f"Error reading '{path}': {e}"


@mcp.tool()
def open_in_explorer(path: str) -> str:
    """Opens Windows Explorer with the file or folder already selected on screen."""
    if not os.path.exists(path):
        return f"Not found: {path}"
    try:
        subprocess.Popen(["explorer", "/select,", os.path.abspath(path)])
        return f"Explorer opened at: {path}"
    except Exception as e:
        return f"Error opening Explorer: {e}"


# --------------------------------------------------------------------------
# Clipboard: save copied image/text without any download at all
# --------------------------------------------------------------------------

def _bitmapfileheader(dib_size: int) -> bytes:
    # BITMAPFILEHEADER (14 bytes): 'BM', total size, reserved, pixel data offset.
    # A DIB coming from the clipboard always starts with a BITMAPINFOHEADER (40 bytes).
    offset = 14 + 40
    return struct.pack("<2sIHHI", b"BM", 14 + dib_size, 0, 0, offset)


@mcp.tool()
def save_clipboard_image(dest_path: str, overwrite: bool = False) -> str:
    """Saves as .bmp whatever image is currently on the clipboard (e.g. after
    Ctrl+C on an image in the browser, without needing the 'Save As' dialog).
    dest_path must end in .bmp.
    """
    error = _safe_dest(dest_path, overwrite)
    if error:
        return error
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                return "Clipboard has no image (CF_DIB format missing). Copy the image (Ctrl+C) first."
            dib = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        finally:
            win32clipboard.CloseClipboard()
        header = _bitmapfileheader(len(dib))
        with open(dest_path, "wb") as f:
            f.write(header)
            f.write(dib)
        return f"Saved: {dest_path} ({len(dib) / 1024:.1f} KB, .bmp)"
    except Exception as e:
        return f"Error saving clipboard image: {e}"


@mcp.tool()
def get_clipboard_text() -> str:
    """Reads the current Windows clipboard text."""
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                return "Clipboard has no text right now."
            return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        return f"Error reading clipboard: {e}"


@mcp.tool()
def set_clipboard_text(text: str) -> str:
    """Writes text to the Windows clipboard (to paste manually later)."""
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
        return f"Clipboard updated ({len(text)} characters)."
    except Exception as e:
        return f"Error writing to clipboard: {e}"


# --------------------------------------------------------------------------
# Upload server: a page sends bytes via fetch POST straight to disk, without
# the base64 going through the chat context (solves large image/video on Flow)
# --------------------------------------------------------------------------

_upload_lock = threading.Lock()
_upload_state = {
    "httpd": None,
    "thread": None,
    "port": None,
    "token": None,
    "dest_dir": None,
    "allowed_origins": set(),
    "max_bytes": 0,
    "log": [],
}


class _UploadHandler(http.server.BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        pass  # silence http.server's default logging

    def _origin(self):
        return self.headers.get("Origin", "")

    def _cors_headers(self):
        origin = self._origin()
        if origin and origin.rstrip("/") in _upload_state["allowed_origins"]:
            self.send_header("Access-Control-Allow-Origin", origin)
        self.send_header("Access-Control-Allow-Headers", "X-Upload-Token, Content-Type")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Vary", "Origin")

    def _reply(self, code, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(code)
        self._cors_headers()
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors_headers()
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_POST(self):
        if self._origin().rstrip("/") not in _upload_state["allowed_origins"]:
            self._reply(403, {"error": "origin not allowed (CORS)"})
            return
        if self.headers.get("X-Upload-Token", "") != _upload_state["token"]:
            self._reply(401, {"error": "invalid token"})
            return
        parsed = urlparse(self.path)
        if parsed.path != "/upload":
            self._reply(404, {"error": "unknown route, use /upload"})
            return
        qs = parse_qs(parsed.query)
        name = os.path.basename(qs.get("name", ["upload.bin"])[0]) or "upload.bin"
        overwrite = qs.get("overwrite", ["0"])[0] == "1"
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            length = 0
        max_bytes = _upload_state["max_bytes"]
        if length <= 0 or length > max_bytes:
            self._reply(413, {"error": f"invalid size or above the limit ({max_bytes} bytes)"})
            return
        dest_path = os.path.join(_upload_state["dest_dir"], name)
        if os.path.exists(dest_path) and not overwrite:
            self._reply(409, {"error": "file already exists (use ?overwrite=1)"})
            return
        data = self.rfile.read(length)
        try:
            with open(dest_path, "wb") as f:
                f.write(data)
        except OSError as e:
            self._reply(500, {"error": f"failed to write: {e}"})
            return
        _upload_state["log"].append((time.time(), dest_path, length))
        self._reply(200, {"ok": True, "path": dest_path, "bytes": length})


@mcp.tool()
def start_upload_server(dest_dir: str, allowed_origins: str = "https://labs.google,https://flow.google.com",
                         port: int = 0, max_mb: int = 500) -> str:
    """Starts an HTTP server bound only to 127.0.0.1 to receive an upload
    directly from a web page via fetch(blob) POST, without the content going
    through your chat context. Use this when save_url_to_file fails due to
    missing session cookies, or the file is too large for base64 (full-size
    image, video, Reel).

    Returns the port, the token (X-Upload-Token header) and a ready-to-run JS
    snippet for the page. allowed_origins: comma-separated list of
    http(s)://domain (no trailing slash) allowed to POST (CORS); only those
    origins can upload a file, and even then they need the right token.
    """
    if not os.path.isdir(dest_dir):
        return f"Folder does not exist: {dest_dir}"
    with _upload_lock:
        if _upload_state["httpd"] is not None:
            return (f"Already running on port {_upload_state['port']}, destination "
                     f"{_upload_state['dest_dir']}. Call stop_upload_server() before starting another.")
        token = secrets.token_urlsafe(24)
        origins = {o.strip().rstrip("/") for o in allowed_origins.split(",") if o.strip()}
        httpd = http.server.ThreadingHTTPServer(("127.0.0.1", port), _UploadHandler)
        actual_port = httpd.server_address[1]
        _upload_state.update({
            "httpd": httpd, "port": actual_port, "token": token,
            "dest_dir": dest_dir, "allowed_origins": origins,
            "max_bytes": max_mb * 1024 * 1024, "log": [],
        })
        thread = threading.Thread(target=httpd.serve_forever, daemon=True)
        _upload_state["thread"] = thread
        thread.start()
    snippet = (
        "fetch('http://127.0.0.1:%d/upload?name=YOURFILE.jpg', {\n"
        "  method: 'POST',\n"
        "  headers: {'X-Upload-Token': '%s'},\n"
        "  body: blob  // Blob/File/ArrayBuffer of what you want to save\n"
        "}).then(r => r.json()).then(console.log)"
    ) % (actual_port, token)
    return (
        f"Upload server up: http://127.0.0.1:{actual_port}/upload\n"
        f"Allowed origins: {', '.join(sorted(origins)) or '(none - nothing will go through)'}\n"
        f"Token (X-Upload-Token header): {token}\n"
        f"Per-file limit: {max_mb} MB\n"
        f"Destination folder: {dest_dir}\n\n"
        f"Snippet to run as JS on the page:\n{snippet}"
    )


@mcp.tool()
def stop_upload_server() -> str:
    """Tears down the upload server started by start_upload_server."""
    with _upload_lock:
        httpd = _upload_state["httpd"]
        if httpd is None:
            return "No upload server running."
        port = _upload_state["port"]
        httpd.shutdown()
        httpd.server_close()
        _upload_state.update({"httpd": None, "thread": None, "port": None, "token": None,
                               "dest_dir": None, "allowed_origins": set(), "max_bytes": 0, "log": []})
    return f"Upload server (port {port}) shut down."


@mcp.tool()
def upload_server_status() -> str:
    """Shows whether the upload server is running and the uploads received this session."""
    with _upload_lock:
        if _upload_state["httpd"] is None:
            return "Stopped."
        lines = [
            f"Running on 127.0.0.1:{_upload_state['port']}, destination: {_upload_state['dest_dir']}",
            f"Origins: {', '.join(sorted(_upload_state['allowed_origins'])) or '(none)'}",
            f"Uploads received: {len(_upload_state['log'])}",
        ]
        for ts, path, size in _upload_state["log"][-20:]:
            lines.append(f"  {time.strftime('%H:%M:%S', time.localtime(ts))}  {path}  ({size / 1024:.1f} KB)")
    return "\n".join(lines)


if __name__ == "__main__":
    mcp.run()
