[CmdletBinding()]
param(
    [int]$Port = 8501,
    [string]$CloudflaredPath = "F:\tools\cloudflared.exe",
    [string]$ConfigPath = "$env:USERPROFILE\.cloudflared\config.yml"
)

$ErrorActionPreference = "Stop"
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$appPath = Join-Path $projectDir "app.py"
$localUrl = "http://127.0.0.1:$Port"

function Test-LocalApp {
    try {
        $response = Invoke-WebRequest -UseBasicParsing -Uri $localUrl -TimeoutSec 3
        return $response.StatusCode -ge 200 -and $response.StatusCode -lt 500
    }
    catch {
        return $false
    }
}

if (-not (Test-Path -LiteralPath $CloudflaredPath)) {
    throw "cloudflared not found: $CloudflaredPath"
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Cloudflare tunnel config not found: $ConfigPath. Run .\cloudflare-tunnel-setup.ps1 first."
}

if (-not (Test-Path -LiteralPath $appPath)) {
    throw "app.py not found: $appPath"
}

if (-not (Test-LocalApp)) {
    Write-Host "Starting Streamlit on $localUrl ..."
    Start-Process `
        -FilePath "python" `
        -ArgumentList @("-m", "streamlit", "run", $appPath, "--server.headless", "true", "--server.port", "$Port") `
        -WorkingDirectory $projectDir `
        -WindowStyle Hidden

    $ready = $false
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 1
        if (Test-LocalApp) {
            $ready = $true
            break
        }
    }

    if (-not $ready) {
        throw "Streamlit did not become ready at $localUrl."
    }
}

Write-Host "Local app is ready: $localUrl"
Write-Host "Starting Cloudflare Tunnel. Keep this window open unless cloudflared is installed as a service."
& $CloudflaredPath tunnel --config $ConfigPath run
