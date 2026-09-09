@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if "%~1"=="" (
  echo Usage:
  echo   SETUP_CLOUDFLARE_TUNNEL.bat your-subdomain.example.com
  echo.
  echo Example:
  echo   SETUP_CLOUDFLARE_TUNNEL.bat cases.example.com
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0cloudflare-tunnel-setup.ps1" -Hostname "%~1"
pause
endlocal
