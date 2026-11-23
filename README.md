# putinha-faz-tudo (ex Windows Process Bridge MCP)

> Servidor MCP (Model Context Protocol) local que dá a um cliente de IA como o Claude Desktop acesso real a processos, serviços, arquivos, clipboard e um servidor de upload do Windows — sem controle de mouse/tela e sem sandbox de VM.

A maioria dos assistentes de IA desktop roda numa VM sandboxed (sem acesso real ao SO) ou cai pra controle de mouse/tela pra checar algo tão simples quanto "esse app está rodando". Esse servidor expõe isso direto, como tool calls de verdade, com design safety-first: operações de leitura são livres, operações que mudam estado exigem `confirm=True` explícito mais a tela de permissão do próprio cliente, e PowerShell arbitrário é bloqueado por uma allowlist de verbos só-leitura (`Get-`, `Measure-`, `Select-`, `Sort-`, `Where-`, `Format-`, `Group-`). Qualquer tentativa de rodar `Stop-`, `Remove-`, `Set-`, `Start-`, `Invoke-`, `New-`, ou de redirecionar pra um arquivo/exe, é recusada antes de executar.

## 🧰 O que ele dá pro seu cliente de IA

**Processos e sistema** (lote original)
- `list_processes` / `check_process(name)` / `open_task_manager` — visibilidade de processo sem mouse
- `run_powershell(cmd)` / `run_powershell_readonly(cmd)` — qualquer one-liner/pipeline PowerShell só-leitura, allowlist-enforced
- `get_process_stats`, `list_processes_grouped`, `system_snapshot`, `gpu_status` — CPU/RAM/VRAM/disco, via `psutil` e `nvidia-smi`
- `start_process(target)` / `kill_process(pid, name, confirm=True)` — abrir ou encerrar, com ação destrutiva atrás de confirm explícito
- `list_services`, `startup_apps`, `disk_usage` — serviços do Windows, programas de inicialização e maiores pastas

**Arquivos** (salvar direto no disco, sem o popup "Salvar como" do navegador)
- `save_url_to_file(url, dest_path, overwrite=False)` — baixa http(s) direto pro disco
- `save_base64_to_file(data_b64, dest_path, overwrite=False)` — grava base64 (ex: extraído de uma `blob:` URL via JS na página)
- `move_file(src, dest, overwrite=False)` / `copy_file(src, dest, overwrite=False)`
- `ensure_dir(path)` / `list_dir(path, pattern="", top=50)` / `list_recent_downloads(minutes=10, pattern="")`
- `delete_file(path, confirm=True)` — manda pra Lixeira via `SHFileOperationW` (reversível, não apaga de vez)
- `file_info(path)` — tamanho, datas, sha256
- `open_in_explorer(path)` — abre o Explorer com o item selecionado

**Clipboard**
- `save_clipboard_image(dest_path, overwrite=False)` — salva como `.bmp` a imagem no clipboard (ex: depois de copiar uma imagem no navegador)
- `get_clipboard_text()` / `set_clipboard_text(text)`

**Servidor de upload** (pra arquivo grande/vídeo, sem passar pelo contexto do chat)
- `start_upload_server(dest_dir, allowed_origins="https://labs.google,https://flow.google.com", port=0, max_mb=500)` — sobe um `http.server.ThreadingHTTPServer` em `127.0.0.1` (porta efêmera), CORS restrito a uma allowlist de origem + token aleatório exigido no header `X-Upload-Token`. A página faz `fetch(blob)` POST direto pro `/upload?name=...`; os bytes vão pro disco sem passar pelo contexto do cliente de IA.
- `stop_upload_server()` / `upload_server_status()`

## 🚫 Limitações conhecidas (testadas ao vivo, não suposições)

### CDP / `--remote-debugging-port` no Claude Desktop: bloqueado de propósito

O Claude Desktop é um app MSIX *full-trust* (Desktop Bridge, não é AppContainer/UWP sandboxed — os dados do usuário vão pra uma pasta normal, não pra `LocalState` virtualizado). Mas o processo principal **recusa iniciar** se a linha de comando contiver qualquer uma de 16 switches de debug/rede (`--remote-debugging-port`, `--inspect`, etc.):

```
Claude: refusing to start — a debugging or network-override switch is present on the command line.
```

Só destrava com um token `CLAUDE_CDP_AUTH` assinado (ed25519, chave compilada no app) — só a Anthropic tem a chave privada. Confirmado de forma independente por terceiros:
- [`bigfnj/desktop-ai-companion`](https://github.com/bigfnj/desktop-ai-companion/blob/master/docs/IDEAS.md) — mesma descoberta via engenharia reversa do `app.asar`.
- [`RemyMachado/claude-desktop-theme`](https://github.com/RemyMachado/claude-desktop-theme) — lista as 16 switches.
- [`anthropics/claude-code#89160`](https://github.com/anthropics/claude-code/issues/89160) e [`#92044`](https://github.com/anthropics/claude-code/issues/92044) — devs tentando debugar outros bugs esbarraram no mesmo bloqueio.

Não é bug, é proteção deliberada contra extração de sessão/token via DevTools. Nenhuma atualização futura do app vai "destravar" isso — não vale a pena tentar de novo.

### Painel de navegador embutido do Claude Desktop tem comportamento próprio

- `save_url_to_file` só funciona com URL http(s) pública, sem exigir cookie de sessão — falha com HTTP 400 em URLs que exigem auth (ex: Google Flow).
- `save_base64_to_file` funciona, mas o base64 passa pelo contexto do chat — viável só pra arquivo pequeno (testado ok ~200 KB).
- `save_clipboard_image` **não funciona** dentro do painel embutido do Claude Desktop: a página recebe `Write permission denied` ao tentar escrever no clipboard, e o clique direito não abre menu nativo ali (diferente de um Chrome de verdade).
- `start_upload_server` funciona bem pra **imagem** em tamanho real (testado: 204 KB, HTTP 200, sem popup). Pra **vídeo** falha: `fetch` do `<video>` dá "Failed to fetch", `captureStream` dá `SecurityError: cross-origin data` — o host de vídeo (ex: Flow) serve por um caminho que o JS da página não consegue ler. Isso é limite do lado do site, não do MCP; o caminho que resta é baixar nativo (aceitar o popup) e usar `list_recent_downloads` + `move_file` pra organizar depois.

## 🚀 Tech Stack

- Python 3.11+
- [`mcp`](https://pypi.org/project/mcp/) (SDK oficial do Model Context Protocol, `FastMCP`)
- [`psutil`](https://pypi.org/project/psutil/) pra stats entre processos
- [`pywin32`](https://pypi.org/project/pywin32/) pra clipboard e Lixeira (`win32clipboard`, `SHFileOperationW`)
- Stdlib `subprocess`, `http.server`, `urllib.request`

## 🛠️ Instalação e desenvolvimento local

### Pré-requisitos
- Windows 10/11
- Python 3.11+

### Passos

```
git clone <repository-url>
cd win-process-bridge-mcp

python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

python server.py   # roda como servidor MCP stdio, esperando um cliente
```

### Registrar no Claude Desktop

Edite `claude_desktop_config.json` (Settings → Developer → Edit Config) e adicione:

```json
{
  "mcpServers": {
    "putinha-faz-tudo": {
      "command": "C:\\caminho\\pro\\.venv\\Scripts\\python.exe",
      "args": ["C:\\caminho\\pro\\win-process-bridge-mcp\\server.py"]
    }
  }
}
```

Reinicie o Claude Desktop por completo (sair pela bandeja do sistema, não só fechar a janela) pro novo servidor MCP carregar. No primeiro uso de qualquer tool, o cliente mostra sua própria tela de permissão antes de rodar qualquer coisa.

## 📄 Estrutura do projeto

```
server.py            Todas as 30 tools num único servidor FastMCP (transporte stdio)
requirements.txt      mcp + psutil + pywin32
```

## 🔒 Modelo de segurança

| Aspecto | Design |
|---|---|
| Operações de leitura | Sem restrição (`list_processes`, `system_snapshot`, `gpu_status`, etc.) |
| PowerShell | Allowlist de verbos só-leitura; regex de bloqueio como segunda camada |
| Operações destrutivas | `kill_process` e `delete_file` exigem `confirm=True` além da UI de permissão do cliente |
| Servidor de upload | Só escuta em `127.0.0.1` (nunca `0.0.0.0`), CORS restrito a uma allowlist explícita de origem, e exige um token aleatório por sessão no header — mesmo uma origem liberada não sobe arquivo sem o token certo |
| Escopo | Só processos/serviços/arquivos explicitamente nomeados na chamada da tool; sem daemon de fundo persistente além do upload server opcional, que o próprio cliente sobe e derruba |
| Transporte | stdio como processo filho do cliente de IA; o único listener de rede é o upload server opcional, e mesmo esse é efêmero e local |

Este projeto não tenta fazer sandbox do próprio PowerShell (sem AppContainer, sem token restrito) — a allowlist é uma defesa cooperativa contra uso acidental ou descuidado por um cliente de IA que já deveria se comportar, não uma fronteira de segurança contra um malicioso. Não exponha esse servidor a um chamador não confiável.

## License

MIT. Veja [LICENSE](LICENSE).

---

© 2026 Nauiter Master | AI Strategist, Digital Artist & Automation
