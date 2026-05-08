@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "PS1=%SCRIPT_DIR%install_mcp_deps.ps1"

if not exist "%PS1%" (
    echo ERRO: Arquivo PowerShell nao encontrado: %PS1%
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %*
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo Instalacao falhou com codigo %EXIT_CODE%.
    exit /b %EXIT_CODE%
)

echo.
echo Instalacao concluida com sucesso.
exit /b 0
