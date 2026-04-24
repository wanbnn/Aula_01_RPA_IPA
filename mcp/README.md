# Browser MCP Server

Servidor MCP em Python usando FastMCP e Playwright para permitir que um modelo navegue pela web de forma autonoma.

## O que ele expoe

- `browser_navigate`: abre uma URL e devolve URL, titulo e texto visivel.
- `browser_get_page_state`: captura o estado atual da pagina.
- `browser_click`: clica em um elemento via seletor CSS.
- `browser_type`: digita em um campo via seletor CSS.
- `browser_press_key`: envia teclas para a pagina ativa.
- `browser_wait_for`: espera um elemento atingir um estado.
- `browser_extract_text`: extrai texto de um elemento.
- `browser_go_back`: volta no historico do navegador.
- `browser_inspect_elements`: lista elementos interativos para ajudar o modelo a escolher seletores.

## Requisitos

- Python 3.11+
- `uv` instalado

## Instalar

Com `uv`:

```powershell
cd c:\Users\jose.wcavalcante\Documents\aula\mcp
uv sync
uv run playwright install chromium
```

Sem `uv`, usando Python diretamente:

```powershell
cd c:\Users\jose.wcavalcante\Documents\aula\mcp
py -m pip install "mcp[cli]>=1.2.0" "playwright>=1.52.0"
py -m playwright install chromium
```

## Executar localmente

Com `uv`:

```powershell
cd c:\Users\jose.wcavalcante\Documents\aula\mcp
uv run browser_server.py
```

Sem `uv`:

```powershell
cd c:\Users\jose.wcavalcante\Documents\aula\mcp
py browser_server.py
```

Para desenvolvimento com Inspector:

```powershell
uv run mcp dev browser_server.py
```

## Variaveis de ambiente

- `BROWSER_MCP_HEADLESS=true|false`
- `BROWSER_MCP_TIMEOUT_MS=15000`
- `BROWSER_MCP_VIEWPORT_WIDTH=1440`
- `BROWSER_MCP_VIEWPORT_HEIGHT=900`
- `BROWSER_MCP_MAX_TEXT=4000`

## Exemplo de configuracao MCP

Exemplo para um cliente que inicia o servidor por `uv`:

```json
{
  "mcpServers": {
    "browser": {
      "command": "uv",
      "args": [
        "--directory",
        "c:/Users/jose.wcavalcante/Documents/aula/mcp",
        "run",
        "browser_server.py"
      ]
    }
  }
}
```

## Observacoes

- O servidor aceita apenas URLs `http` e `https`.
- Para nao quebrar o protocolo MCP em `stdio`, o codigo nao escreve em `stdout` fora do SDK.
- O estado do navegador e mantido durante a sessao, o que permite navegacao multi-etapas.