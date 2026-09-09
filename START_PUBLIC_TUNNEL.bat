@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0cloudflare-tunnel-run.ps1"
pause
endlocal
