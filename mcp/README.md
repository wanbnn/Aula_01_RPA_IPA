# Automation MCP Servers

Servidores MCP em Python usando FastMCP para automacao autonoma por modelos de IA.

- `browser_server.py`: controla navegador via Playwright.
- `excel_server.py`: controla Microsoft Excel via pywin32/COM no Windows.

## Requisitos

- Python 3.11+
- `uv` instalado ou `pip`
- Para navegador: Playwright Chromium
- Para Excel: Windows + Microsoft Excel instalado + `pywin32`

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
py -m pip install "mcp[cli]>=1.2.0" "playwright>=1.52.0" "pywin32>=306"
py -m playwright install chromium
```

## Browser MCP Server

### Ferramentas expostas

- `browser_navigate`: abre uma URL e devolve URL, titulo e texto visivel.
- `browser_get_page_state`: captura o estado atual da pagina.
- `browser_click`: clica em um elemento via seletor CSS.
- `browser_type`: digita em um campo via seletor CSS.
- `browser_press_key`: envia teclas para a pagina ativa.
- `browser_wait_for`: espera um elemento atingir um estado.
- `browser_extract_text`: extrai texto de um elemento.
- `browser_go_back`: volta no historico do navegador.
- `browser_inspect_elements`: lista elementos interativos para ajudar o modelo a escolher seletores.

### Executar

```powershell
uv run browser_server.py
# ou
py browser_server.py
```

### Variaveis de ambiente

- `BROWSER_MCP_HEADLESS=true|false`
- `BROWSER_MCP_TIMEOUT_MS=15000`
- `BROWSER_MCP_VIEWPORT_WIDTH=1440`
- `BROWSER_MCP_VIEWPORT_HEIGHT=900`
- `BROWSER_MCP_MAX_TEXT=4000`

## Excel MCP Server

Servidor MCP para criar relatorios, planilhas financeiras, estudos, analises estatisticas, tabelas, graficos e exportacoes em PDF usando Microsoft Excel real via pywin32.

### Ferramentas expostas

- `excel_create_workbook`: cria um novo arquivo Excel.
- `excel_open_workbook`: abre um arquivo existente.
- `excel_save_workbook`: salva ou salva como `.xlsx`, `.xlsm`, `.xlsb`, `.xls` ou `.csv`.
- `excel_close_workbook`: fecha o workbook ativo.
- `excel_quit`: encerra a instancia do Excel controlada pelo servidor.
- `excel_workbook_info`: retorna metadados do workbook ativo.
- `excel_list_sheets`: lista planilhas.
- `excel_add_sheet`: cria planilha.
- `excel_delete_sheet`: exclui planilha.
- `excel_activate_sheet`: ativa planilha.
- `excel_write_range`: escreve matriz de dados em um range.
- `excel_read_range`: le valores de um range com limite de seguranca.
- `excel_clear_range`: limpa conteudo ou conteudo e formatos.
- `excel_set_formula`: define formulas Excel.
- `excel_format_range`: aplica negrito, fonte, cores, formatos numericos, alinhamento, bordas e autofit.
- `excel_create_table`: cria tabela estruturada do Excel.
- `excel_create_chart`: cria graficos de coluna, barra, linha, pizza, area ou dispersao.
- `excel_autofit`: ajusta linhas e colunas.
- `excel_export_pdf`: exporta workbook ou planilha para PDF.

### Executar

```powershell
uv run excel_server.py
# ou
py excel_server.py
```

### Variaveis de ambiente

- `EXCEL_MCP_VISIBLE=true|false` — mostra ou oculta a janela do Excel. Padrao: `true`.
- `EXCEL_MCP_DISPLAY_ALERTS=true|false` — controla alertas interativos do Excel. Padrao: `false`.
- `EXCEL_MCP_MAX_READ_CELLS=5000` — limite de leitura por chamada.
- `EXCEL_MCP_OUTPUT_DIR=c:\relatorios` — diretorio padrao para salvar arquivos quando usado caminho relativo ou nome simples.

### Fluxo recomendado para modelos de IA

1. Use `excel_create_workbook` ou `excel_open_workbook`.
2. Use `excel_add_sheet` para organizar secoes do relatorio.
3. Use `excel_write_range` para inserir dados e textos.
4. Use `excel_set_formula` para calculos.
5. Use `excel_format_range`, `excel_create_table` e `excel_create_chart` para acabamento.
6. Use `excel_save_workbook` e, se necessario, `excel_export_pdf`.
7. Use `excel_quit` ao finalizar a sessao.

## Exemplo de configuracao MCP

Exemplo para um cliente que inicia os servidores por `uv`:

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
    },
    "excel": {
      "command": "uv",
      "args": [
        "--directory",
        "c:/Users/jose.wcavalcante/Documents/aula/mcp",
        "run",
        "excel_server.py"
      ],
      "env": {
        "EXCEL_MCP_VISIBLE": "true",
        "EXCEL_MCP_DISPLAY_ALERTS": "false",
        "EXCEL_MCP_OUTPUT_DIR": "c:/Users/jose.wcavalcante/Documents/aula/relatorios"
      }
    }
  }
}
```

## Observacoes

- O servidor Browser aceita apenas URLs `http` e `https`.
- O servidor Excel funciona apenas no Windows com Microsoft Excel instalado.
- O servidor Excel usa COM single-instance com lock interno para evitar chamadas concorrentes no Excel.
- Para nao quebrar o protocolo MCP em `stdio`, os servidores nao escrevem em `stdout` fora do SDK.
