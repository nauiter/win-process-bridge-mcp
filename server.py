"""
Windows Process Bridge - servidor MCP local (stdio).

Usado pelo mcpServers do Claude Desktop (Settings > Desenvolvedor >
Servidores MCP locais). Roda como processo filho do proprio Claude,
em stdio - sem rede, sem VM sandbox.

Ferramentas:
  - list_processes()               -> lista todos os processos do Windows
  - check_process(name)            -> 'running'/'closed' para um processo
  - open_task_manager()            -> abre o Gerenciador de Tarefas na tela
  - run_powershell(cmd)            -> cmdlet PowerShell so-leitura (allowlist)
  - run_powershell_readonly(cmd)   -> alias do run_powershell
  - get_process_stats(sort_by,top) -> CPU, RAM e threads por processo
  - list_processes_grouped()       -> soma de RAM por nome de processo
  - system_snapshot()              -> RAM, CPU e disco numa chamada
  - gpu_status()                   -> uso de GPU/VRAM via nvidia-smi
  - start_process(target)          -> abre um app por nome ou caminho
  - kill_process(pid, name, confirm) -> encerra por PID ou nome
  - list_services(name_filter)     -> servicos do Windows e estado
  - startup_apps()                 -> o que abre junto com o Windows
  - disk_usage(scan_path, top)     -> espaco livre por unidade + maiores pastas
  - save_url_to_file(url, dest)    -> baixa http(s) direto pro disco, sem popup
  - save_base64_to_file(b64, dest) -> grava base64 (de blob: extraido via JS) no disco
  - move_file(src, dest)           -> move/renomeia (ex: Downloads -> pasta do projeto)
  - copy_file(src, dest)           -> copia mantendo o original
  - list_recent_downloads(min)     -> acha o que acabou de cair em Downloads
  - ensure_dir(path)               -> cria pasta (e pais) se nao existir
  - list_dir(path, pattern, top)   -> lista qualquer pasta, nao so Downloads
  - delete_file(path, confirm)     -> manda pra Lixeira (reversivel), exige confirm
  - file_info(path)                -> tamanho, datas e sha256 de um arquivo
  - open_in_explorer(path)         -> abre o Explorer com o arquivo selecionado
  - save_clipboard_image(dest)     -> salva imagem copiada (Ctrl+C) do clipboard
  - get_clipboard_text()           -> le texto atual do clipboard
  - set_clipboard_text(text)       -> escreve texto no clipboard
  - start_upload_server(dest_dir)  -> servidor local p/ pagina enviar bytes via
                                       fetch POST, sem passar pelo contexto do chat
  - stop_upload_server()           -> derruba o servidor de upload
  - upload_server_status()         -> estado e uploads recebidos nesta sessao
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

# Allowlist de cmdlets read-only para run_powershell. O primeiro token de
# cada estagio do pipeline precisa comecar com um destes verbos seguros.
# Qualquer coisa fora disso (Stop-, Remove-, Set-, Start-, Invoke-, New-,
# redirecionamento de arquivo, chamada de .exe arbitrario) e recusada.
_SAFE_VERB_PREFIXES = ("get-", "measure-", "select-", "sort-", "where-",
                       "format-", "group-")
_BLOCKED_PATTERNS = re.compile(
    r"[;&|`]{2,}|(^|\s)(stop-|remove-|set-|start-|invoke-|new-|rm |del |"
    r"iwr|invoke-webrequest|invoke-expression|iex)\b",
    re.IGNORECASE,
)

# Extensao (18/11/2026): libera, e so isso, o pipeline do ccusage.
#  - 1o estagio pode ser exatamente "npx ccusage@latest [--flags]"
#  - estagio "ConvertFrom-Json" sem argumentos
#  - estagio "ForEach-Object" apenas com blocos { } contendo variaveis,
#    numeros e operadores aritmeticos (sem comandos, chamadas ou strings)
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
# Processos basicos (tasklist)
# --------------------------------------------------------------------------

@mcp.tool()
def list_processes() -> str:
    """Lista todos os processos rodando no Windows agora (nome, PID, memoria)."""
    result = subprocess.run(
        ["tasklist", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp850", errors="replace",
    )
    return result.stdout or result.stderr or "sem saida"


@mcp.tool()
def check_process(name: str) -> str:
    """Verifica se um processo especifico esta rodando (ex: 'CapCut.exe').

    Retorna 'STATE=running' ou 'STATE=closed' seguido dos detalhes do tasklist.
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
    """Abre o Gerenciador de Tarefas do Windows na tela do usuario."""
    subprocess.Popen(["taskmgr.exe"])
    return "Gerenciador de Tarefas aberto."


# --------------------------------------------------------------------------
# PowerShell read-only
# --------------------------------------------------------------------------

def _validate_readonly_powershell(cmd: str) -> str | None:
    """Retorna None se o comando e seguro, ou uma mensagem de erro."""
    if _BLOCKED_PATTERNS.search(cmd):
        return "Bloqueado: comando contem verbo/padrao perigoso (Stop-, Remove-, Set-, Invoke-, etc.)."
    stages = [s.strip() for s in cmd.split("|") if s.strip()]
    if not stages:
        return "Bloqueado: comando vazio."
    for i, stage in enumerate(stages):
        first_word = stage.split()[0].lower()
        if i == 0 and (_NPX_RE.match(stage) or _CCUSAGE_RE.match(stage)):
            continue
        if first_word == "convertfrom-json" and len(stage.split()) == 1:
            continue
        if first_word == "foreach-object":
            if _foreach_ok(stage):
                continue
            return ("Bloqueado: ForEach-Object so aceita blocos com variaveis, "
                    "numeros e operadores aritmeticos.")
        if not any(first_word.startswith(p) for p in _SAFE_VERB_PREFIXES):
            return (
                f"Bloqueado: '{first_word}' nao esta na allowlist "
                f"({', '.join(_SAFE_VERB_PREFIXES)})."
            )
    return None


@mcp.tool()
def run_powershell(cmd: str) -> str:
    """Roda um comando PowerShell SOMENTE LEITURA (allowlist: Get-, Measure-,
    Select-, Sort-, Where-, Format-, Group-). Qualquer verbo que altere
    estado (Stop-, Remove-, Set-, Start-, Invoke-, New-) e recusado antes
    de executar. Ex: 'Get-Process | Sort-Object CPU -Descending | Select-Object -First 5'
    """
    error = _validate_readonly_powershell(cmd)
    if error:
        return error
    # npx.ps1 e barrado pela ExecutionPolicy; o shim npx.cmd nao e script PowerShell.
    exec_cmd = re.sub(r"^\s*npx(\s)", r"npx.cmd\1", cmd, count=1, flags=re.IGNORECASE)
    exec_cmd = re.sub(r"^\s*ccusage(\s|$)", lambda m: _CCUSAGE_CMD + m.group(1), exec_cmd,
                      count=1, flags=re.IGNORECASE)
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", exec_cmd],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=120 if (_NPX_RE.match(cmd.split("|")[0].strip()) or _CCUSAGE_RE.match(cmd.split("|")[0].strip())) else 30,
    )
    return result.stdout or result.stderr or "sem saida"


@mcp.tool()
def run_powershell_readonly(cmd: str) -> str:
    """Alias de run_powershell: cmdlet PowerShell so-leitura, mesma allowlist."""
    return run_powershell(cmd)


# --------------------------------------------------------------------------
# Estatisticas via psutil
# --------------------------------------------------------------------------

@mcp.tool()
def get_process_stats(sort_by: str = "cpu", top: int = 10) -> str:
    """CPU (%), RAM (MB) e threads por processo, ordenado.

    sort_by: 'cpu', 'mem' ou 'threads'. top: quantos processos retornar.
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
    lines = ["nome,pid,cpu_%,ram_mb,threads"]
    for name, pid, cpu, mem_mb, threads in rows[:top]:
        lines.append(f"{name},{pid},{cpu:.1f},{mem_mb:.1f},{threads}")
    return "\n".join(lines)


@mcp.tool()
def list_processes_grouped() -> str:
    """Soma de RAM e contagem de instancias por nome de processo (ex: Chrome)."""
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
    lines = ["nome,ram_total_mb,instancias"]
    for name, (mem_total, count) in rows:
        lines.append(f"{name},{mem_total:.1f},{count}")
    return "\n".join(lines)


@mcp.tool()
def system_snapshot() -> str:
    """RAM, CPU e disco do sistema numa unica chamada."""
    vm = psutil.virtual_memory()
    cpu_pct = psutil.cpu_percent(interval=0.5)
    lines = [
        f"CPU: {cpu_pct:.1f}%",
        f"RAM: {vm.used / (1024**3):.1f} GB usados de {vm.total / (1024**3):.1f} GB ({vm.percent:.1f}%)",
        "Discos:",
    ]
    for part in psutil.disk_partitions(all=False):
        try:
            u = psutil.disk_usage(part.mountpoint)
            lines.append(
                f"  {part.device} {u.used / (1024**3):.1f} GB usados de "
                f"{u.total / (1024**3):.1f} GB ({u.percent:.1f}%)"
            )
        except (PermissionError, OSError):
            continue
    return "\n".join(lines)


@mcp.tool()
def gpu_status() -> str:
    """Uso de GPU e VRAM via nvidia-smi (NVIDIA apenas)."""
    result = subprocess.run(
        ["nvidia-smi",
         "--query-gpu=name,utilization.gpu,memory.used,memory.total,temperature.gpu",
         "--format=csv,noheader"],
        capture_output=True, text=True, errors="replace", timeout=10,
    )
    if result.returncode != 0:
        return f"nvidia-smi indisponivel: {result.stderr.strip() or 'sem GPU NVIDIA'}"
    return result.stdout.strip() or "sem saida"


# --------------------------------------------------------------------------
# Controle de processos (requer confirm explicito para a parte destrutiva)
# --------------------------------------------------------------------------

@mcp.tool()
def start_process(target: str) -> str:
    """Abre um aplicativo pelo nome (ex: 'notepad') ou caminho completo do .exe."""
    try:
        subprocess.Popen(target, shell=True)
        return f"Comando de abertura enviado para: {target}"
    except Exception as e:
        return f"Erro ao abrir '{target}': {e}"


@mcp.tool()
def kill_process(pid: int = 0, name: str = "", confirm: bool = False) -> str:
    """Encerra um processo por PID ou nome. Exige confirm=True (defesa extra
    alem da tela de permissao do Claude) - sem isso, so simula e nao mata nada.
    """
    if not confirm:
        return "Nao executado: passe confirm=True para realmente encerrar o processo."
    if pid:
        result = subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                                 capture_output=True, text=True, encoding="cp850", errors="replace")
    elif name:
        result = subprocess.run(["taskkill", "/F", "/IM", name],
                                 capture_output=True, text=True, encoding="cp850", errors="replace")
    else:
        return "Informe pid ou name."
    return result.stdout or result.stderr or "sem saida"


# --------------------------------------------------------------------------
# Servicos, inicializacao e disco
# --------------------------------------------------------------------------

@mcp.tool()
def list_services(name_filter: str = "") -> str:
    """Lista servicos do Windows e seu estado (running/stopped).

    name_filter opcional: substring do nome do servico (case-insensitive).
    """
    lines = ["nome,display_name,status"]
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
    """Lista programas configurados para abrir junto com o Windows."""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command",
         "Get-CimInstance Win32_StartupCommand | Select-Object Name,Command,Location "
         "| ConvertTo-Csv -NoTypeInformation"],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30,
    )
    return result.stdout or result.stderr or "sem saida"


@mcp.tool()
def disk_usage(scan_path: str = "", top: int = 10) -> str:
    """Espaco livre/usado por unidade. Se scan_path for passado, tambem lista
    as 'top' maiores subpastas de 1o nivel dentro desse caminho (pode demorar
    em pastas muito grandes).
    """
    lines = ["Unidades:"]
    for part in psutil.disk_partitions(all=False):
        try:
            u = psutil.disk_usage(part.mountpoint)
            lines.append(
                f"  {part.device} livre: {u.free / (1024**3):.1f} GB de "
                f"{u.total / (1024**3):.1f} GB"
            )
        except (PermissionError, OSError):
            continue

    if scan_path:
        if not os.path.isdir(scan_path):
            lines.append(f"\nscan_path invalido: {scan_path}")
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
            lines.append(f"\nSem permissao para escanear: {scan_path}")
            return "\n".join(lines)
        sizes.sort(key=lambda x: x[1], reverse=True)
        lines.append(f"\nMaiores pastas em {scan_path}:")
        for name, total in sizes[:top]:
            lines.append(f"  {name}: {total / (1024**3):.2f} GB")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# Arquivos: salvar direto no disco sem o popup "Salvar como" do navegador
# --------------------------------------------------------------------------

def _safe_dest(dest_path: str, overwrite: bool) -> str | None:
    """Retorna mensagem de erro, ou None se dest_path pode ser escrito."""
    parent = os.path.dirname(dest_path) or "."
    if not os.path.isdir(parent):
        return f"Pasta nao existe: {parent}"
    if os.path.exists(dest_path) and not overwrite:
        return f"Ja existe: {dest_path} (passe overwrite=True para substituir)."
    return None


@mcp.tool()
def save_url_to_file(url: str, dest_path: str, overwrite: bool = False) -> str:
    """Baixa uma URL http(s) direto para dest_path, sem abrir dialogo do navegador.

    Nao funciona com URLs blob: (essas so existem dentro da aba do navegador;
    use save_base64_to_file para esse caso, extraindo o base64 via JS na pagina).
    """
    if not re.match(r"^https?://", url, re.IGNORECASE):
        return "Bloqueado: so aceito URL http:// ou https:// (blob: nao e baixavel aqui)."
    error = _safe_dest(dest_path, overwrite)
    if error:
        return error
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest_path, "wb") as f:
            f.write(resp.read())
        size = os.path.getsize(dest_path)
        return f"Salvo: {dest_path} ({size / 1024:.1f} KB)"
    except Exception as e:
        return f"Erro ao baixar '{url}': {e}"


@mcp.tool()
def save_base64_to_file(data_b64: str, dest_path: str, overwrite: bool = False) -> str:
    """Decodifica base64 (ex: extraido de um blob: via JS na pagina, sem o data:URI
    prefix) e grava em dest_path. Use quando a origem for uma blob: URL do navegador.
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
        return f"Salvo: {dest_path} ({len(raw) / 1024:.1f} KB)"
    except Exception as e:
        return f"Erro ao decodificar/gravar '{dest_path}': {e}"


@mcp.tool()
def move_file(src: str, dest: str, overwrite: bool = False) -> str:
    """Move/renomeia um arquivo (ex: do que acabou de cair em Downloads para a
    pasta certa do projeto). Falha se dest ja existir e overwrite=False.
    """
    if not os.path.isfile(src):
        return f"Nao encontrado: {src}"
    error = _safe_dest(dest, overwrite)
    if error:
        return error
    try:
        if overwrite and os.path.exists(dest):
            os.remove(dest)
        os.replace(src, dest)
        return f"Movido: {src} -> {dest}"
    except Exception as e:
        return f"Erro ao mover '{src}' -> '{dest}': {e}"


@mcp.tool()
def copy_file(src: str, dest: str, overwrite: bool = False) -> str:
    """Copia um arquivo mantendo o original (ex: duplicar antes de editar)."""
    if not os.path.isfile(src):
        return f"Nao encontrado: {src}"
    error = _safe_dest(dest, overwrite)
    if error:
        return error
    try:
        shutil.copy2(src, dest)
        return f"Copiado: {src} -> {dest}"
    except Exception as e:
        return f"Erro ao copiar '{src}' -> '{dest}': {e}"


@mcp.tool()
def ensure_dir(path: str) -> str:
    """Cria a pasta (e as pastas pai necessarias) se ainda nao existir."""
    try:
        os.makedirs(path, exist_ok=True)
        return f"Pasta pronta: {path}"
    except Exception as e:
        return f"Erro ao criar '{path}': {e}"


@mcp.tool()
def list_recent_downloads(minutes: int = 10, pattern: str = "") -> str:
    """Lista arquivos criados/modificados em Downloads nos ultimos N minutos
    (para achar o que acabou de ser salvo pelo popup do navegador e mover
    com move_file). pattern opcional: substring do nome (case-insensitive).
    """
    downloads = os.path.join(os.path.expanduser("~"), "Downloads")
    if not os.path.isdir(downloads):
        return f"Pasta nao encontrada: {downloads}"
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
        return f"Nada em {downloads} nos ultimos {minutes} min (filtro: '{pattern}')."
    rows.sort(reverse=True)
    lines = [f"Downloads recentes (ultimos {minutes} min):"]
    for mtime, name, path, size in rows:
        lines.append(f"  {name}  ({size / 1024:.1f} KB)  {path}")
    return "\n".join(lines)


@mcp.tool()
def list_dir(path: str, pattern: str = "", top: int = 50) -> str:
    """Lista arquivos e pastas dentro de 'path' (nao recursivo).

    pattern opcional: substring do nome (case-insensitive). top: limite de linhas.
    """
    if not os.path.isdir(path):
        return f"Pasta nao encontrada: {path}"
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
        return f"Sem permissao para listar: {path}"
    rows.sort(reverse=True)
    lines = [f"{path} ({len(rows)} itens, mais recentes primeiro):"]
    for mtime, kind, name, size in rows[:top]:
        size_s = f"{size / 1024:.1f} KB" if kind == "FILE" else "-"
        lines.append(f"  [{kind}] {name}  {size_s}")
    return "\n".join(lines)


@mcp.tool()
def delete_file(path: str, confirm: bool = False) -> str:
    """Manda um arquivo para a Lixeira do Windows (reversivel, nao apaga de vez).

    Exige confirm=True (defesa extra alem da tela de permissao do Claude).
    """
    if not confirm:
        return "Nao executado: passe confirm=True para realmente enviar a Lixeira."
    if not os.path.isfile(path):
        return f"Nao encontrado: {path}"
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
            return f"Falha ao enviar a Lixeira (codigo {ret})."
        return f"Enviado para a Lixeira: {path}"
    except Exception as e:
        return f"Erro ao apagar '{path}': {e}"


@mcp.tool()
def file_info(path: str) -> str:
    """Tamanho, data de criacao/modificacao e sha256 de um arquivo."""
    if not os.path.isfile(path):
        return f"Nao encontrado: {path}"
    try:
        st = os.stat(path)
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
        return (
            f"Caminho: {path}\n"
            f"Tamanho: {st.st_size / 1024:.1f} KB\n"
            f"Modificado: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.st_mtime))}\n"
            f"Criado: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(st.st_ctime))}\n"
            f"SHA256: {h.hexdigest()}"
        )
    except Exception as e:
        return f"Erro ao ler '{path}': {e}"


@mcp.tool()
def open_in_explorer(path: str) -> str:
    """Abre o Windows Explorer com o arquivo ou pasta ja selecionado na tela."""
    if not os.path.exists(path):
        return f"Nao encontrado: {path}"
    try:
        subprocess.Popen(["explorer", "/select,", os.path.abspath(path)])
        return f"Explorer aberto em: {path}"
    except Exception as e:
        return f"Erro ao abrir Explorer: {e}"


# --------------------------------------------------------------------------
# Clipboard: salvar imagem/texto copiados sem passar por download algum
# --------------------------------------------------------------------------

def _bitmapfileheader(dib_size: int) -> bytes:
    # BITMAPFILEHEADER (14 bytes): 'BM', tamanho total, reservado, offset dos pixels.
    # DIB vindo do clipboard sempre comeca com BITMAPINFOHEADER (40 bytes).
    offset = 14 + 40
    return struct.pack("<2sIHHI", b"BM", 14 + dib_size, 0, 0, offset)


@mcp.tool()
def save_clipboard_image(dest_path: str, overwrite: bool = False) -> str:
    """Salva como .bmp a imagem que estiver no clipboard (ex: apos Ctrl+C numa
    imagem no navegador, sem precisar do dialogo 'Salvar como'). dest_path deve
    terminar em .bmp.
    """
    error = _safe_dest(dest_path, overwrite)
    if error:
        return error
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                return "Clipboard nao tem imagem (formato CF_DIB ausente). Copie a imagem (Ctrl+C) primeiro."
            dib = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
        finally:
            win32clipboard.CloseClipboard()
        header = _bitmapfileheader(len(dib))
        with open(dest_path, "wb") as f:
            f.write(header)
            f.write(dib)
        return f"Salvo: {dest_path} ({len(dib) / 1024:.1f} KB, .bmp)"
    except Exception as e:
        return f"Erro ao salvar imagem do clipboard: {e}"


@mcp.tool()
def get_clipboard_text() -> str:
    """Le o texto atual do clipboard do Windows."""
    try:
        win32clipboard.OpenClipboard()
        try:
            if not win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                return "Clipboard nao tem texto no momento."
            return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        return f"Erro ao ler clipboard: {e}"


@mcp.tool()
def set_clipboard_text(text: str) -> str:
    """Escreve um texto no clipboard do Windows (para colar manualmente depois)."""
    try:
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
        finally:
            win32clipboard.CloseClipboard()
        return f"Clipboard atualizado ({len(text)} caracteres)."
    except Exception as e:
        return f"Erro ao escrever no clipboard: {e}"


# --------------------------------------------------------------------------
# Upload server: pagina manda bytes via fetch POST direto pro disco, sem
# o base64 passar pelo contexto do chat (resolve imagem grande/video do Flow)
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
        pass  # silencia o log padrao do http.server

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
            self._reply(403, {"error": "origem nao autorizada (CORS)"})
            return
        if self.headers.get("X-Upload-Token", "") != _upload_state["token"]:
            self._reply(401, {"error": "token invalido"})
            return
        parsed = urlparse(self.path)
        if parsed.path != "/upload":
            self._reply(404, {"error": "rota desconhecida, use /upload"})
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
            self._reply(413, {"error": f"tamanho invalido ou acima do limite ({max_bytes} bytes)"})
            return
        dest_path = os.path.join(_upload_state["dest_dir"], name)
        if os.path.exists(dest_path) and not overwrite:
            self._reply(409, {"error": "arquivo ja existe (use ?overwrite=1)"})
            return
        data = self.rfile.read(length)
        try:
            with open(dest_path, "wb") as f:
                f.write(data)
        except OSError as e:
            self._reply(500, {"error": f"falha ao gravar: {e}"})
            return
        _upload_state["log"].append((time.time(), dest_path, length))
        self._reply(200, {"ok": True, "path": dest_path, "bytes": length})


@mcp.tool()
def start_upload_server(dest_dir: str, allowed_origins: str = "https://labs.google,https://flow.google.com",
                         port: int = 0, max_mb: int = 500) -> str:
    """Sobe um servidor HTTP so em 127.0.0.1 pra receber upload direto de uma
    pagina via fetch(blob) POST, sem o conteudo passar pelo seu contexto de
    chat. Use quando save_url_to_file falhar por falta de cookies de sessao,
    ou o arquivo for grande demais pra ir em base64 (imagem em tamanho real,
    video, Reel).

    Retorna a porta, o token (header X-Upload-Token) e um snippet de JS pronto
    pra rodar na pagina. allowed_origins: lista separada por virgula dos
    http(s)://dominio (sem barra final) que podem fazer POST (CORS); so eles
    conseguem subir arquivo, e mesmo assim precisam do token certo.
    """
    if not os.path.isdir(dest_dir):
        return f"Pasta nao existe: {dest_dir}"
    with _upload_lock:
        if _upload_state["httpd"] is not None:
            return (f"Ja rodando na porta {_upload_state['port']}, destino "
                     f"{_upload_state['dest_dir']}. Chame stop_upload_server() antes de subir outro.")
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
        "fetch('http://127.0.0.1:%d/upload?name=SEUARQUIVO.jpg', {\n"
        "  method: 'POST',\n"
        "  headers: {'X-Upload-Token': '%s'},\n"
        "  body: blob  // Blob/File/ArrayBuffer do que voce quer salvar\n"
        "}).then(r => r.json()).then(console.log)"
    ) % (actual_port, token)
    return (
        f"Servidor de upload no ar: http://127.0.0.1:{actual_port}/upload\n"
        f"Origens liberadas: {', '.join(sorted(origins)) or '(nenhuma - nada vai passar)'}\n"
        f"Token (header X-Upload-Token): {token}\n"
        f"Limite por arquivo: {max_mb} MB\n"
        f"Pasta destino: {dest_dir}\n\n"
        f"Snippet pra rodar como JS na pagina:\n{snippet}"
    )


@mcp.tool()
def stop_upload_server() -> str:
    """Derruba o servidor de upload iniciado por start_upload_server."""
    with _upload_lock:
        httpd = _upload_state["httpd"]
        if httpd is None:
            return "Nenhum servidor de upload rodando."
        port = _upload_state["port"]
        httpd.shutdown()
        httpd.server_close()
        _upload_state.update({"httpd": None, "thread": None, "port": None, "token": None,
                               "dest_dir": None, "allowed_origins": set(), "max_bytes": 0, "log": []})
    return f"Servidor de upload (porta {port}) encerrado."


@mcp.tool()
def upload_server_status() -> str:
    """Mostra se o servidor de upload esta rodando e os uploads recebidos nesta sessao."""
    with _upload_lock:
        if _upload_state["httpd"] is None:
            return "Parado."
        lines = [
            f"Rodando em 127.0.0.1:{_upload_state['port']}, destino: {_upload_state['dest_dir']}",
            f"Origens: {', '.join(sorted(_upload_state['allowed_origins'])) or '(nenhuma)'}",
            f"Uploads recebidos: {len(_upload_state['log'])}",
        ]
        for ts, path, size in _upload_state["log"][-20:]:
            lines.append(f"  {time.strftime('%H:%M:%S', time.localtime(ts))}  {path}  ({size / 1024:.1f} KB)")
    return "\n".join(lines)


if __name__ == "__main__":
    mcp.run()
