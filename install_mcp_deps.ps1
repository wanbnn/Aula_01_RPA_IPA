[CmdletBinding()]
param(
    [switch]$User,
    [switch]$SkipBrowserInstall,
    [switch]$SkipExcelCheck,
    [switch]$Help
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Show-Help {
    Write-Host "Instala dependencias dos MCP servers browser_server.py e excel_server.py."
    Write-Host ""
    Write-Host "Uso:"
    Write-Host "  powershell -ExecutionPolicy Bypass -File .\install_mcp_deps.ps1"
    Write-Host ""
    Write-Host "Opcoes:"
    Write-Host "  -User                 Instala pacotes Python com --user."
    Write-Host "  -SkipBrowserInstall   Nao executa playwright install chromium."
    Write-Host "  -SkipExcelCheck       Nao valida COM do Microsoft Excel."
    Write-Host "  -Help                 Exibe esta ajuda."
}

if ($Help) {
    Show-Help
    exit 0
}

function Resolve-PythonCommand {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($null -ne $python) {
        return @{ Command = $python.Source; Args = @() }
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($null -ne $py) {
        return @{ Command = $py.Source; Args = @("-3") }
    }

    throw "Python 3.11+ nao encontrado. Instale em https://www.python.org/downloads/windows/ e marque 'Add python.exe to PATH'."
}

$pythonSpec = Resolve-PythonCommand
$pythonCommand = [string]$pythonSpec.Command
$pythonPrefixArgs = @($pythonSpec.Args)

function Invoke-Python {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    & $pythonCommand @pythonPrefixArgs @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Comando Python falhou: $pythonCommand $($pythonPrefixArgs -join ' ') $($Arguments -join ' ')"
    }
}

function Get-PythonOutput {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $output = & $pythonCommand @pythonPrefixArgs @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Comando Python falhou: $pythonCommand $($pythonPrefixArgs -join ' ') $($Arguments -join ' ')"
    }
    return ($output | Out-String).Trim()
}

$repoRoot = $PSScriptRoot
$mcpDir = Join-Path $repoRoot "mcp"
if (-not (Test-Path (Join-Path $mcpDir "browser_server.py"))) {
    throw "Arquivo nao encontrado: $mcpDir\browser_server.py"
}
if (-not (Test-Path (Join-Path $mcpDir "excel_server.py"))) {
    throw "Arquivo nao encontrado: $mcpDir\excel_server.py"
}

$version = Get-PythonOutput -Arguments @("-c", "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')")
$versionOk = Get-PythonOutput -Arguments @("-c", "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)")
Write-Host "Python detectado: $version"

Write-Host "Atualizando pip..."
Invoke-Python -Arguments @("-m", "pip", "install", "--upgrade", "pip")

$pipInstallArgs = @("-m", "pip", "install", "--upgrade")
if ($User) {
    $pipInstallArgs += "--user"
}
$pipInstallArgs += @("mcp[cli]>=1.2.0", "playwright>=1.52.0", "pywin32>=306")

Write-Host "Instalando dependencias Python: mcp[cli], playwright, pywin32..."
Invoke-Python -Arguments $pipInstallArgs

if (-not $SkipBrowserInstall) {
    Write-Host "Instalando navegador Chromium do Playwright..."
    Invoke-Python -Arguments @("-m", "playwright", "install", "chromium")
}

Write-Host "Validando sintaxe dos servidores MCP..."
Invoke-Python -Arguments @("-m", "py_compile", (Join-Path $mcpDir "browser_server.py"), (Join-Path $mcpDir "excel_server.py"))

Write-Host "Validando imports principais..."
Invoke-Python -Arguments @("-c", "import mcp, playwright, win32com.client; print('imports ok')")

if (-not $SkipExcelCheck) {
    Write-Host "Validando Microsoft Excel via COM..."
    Invoke-Python -Arguments @("-c", "import win32com.client; excel = win32com.client.DispatchEx('Excel.Application'); print('Excel COM ok - version', excel.Version); excel.Quit()")
}

Write-Host ""
Write-Host "Instalacao concluida com sucesso."
Write-Host "Servidores disponiveis:"
Write-Host "  Browser: python .\mcp\browser_server.py"
Write-Host "  Excel:   python .\mcp\excel_server.py"
