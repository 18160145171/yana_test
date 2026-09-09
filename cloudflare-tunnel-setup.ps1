[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Hostname,

    [string]$TunnelName = "testcase-generator",
    [string]$CloudflaredPath = "F:\tools\cloudflared.exe",
    [string]$LocalUrl = "http://127.0.0.1:8501",
    [switch]$OverwriteDns,
    [switch]$InstallService
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Get-TunnelByName {
    param([string]$Name)

    $json = & $CloudflaredPath tunnel list --output json
    if ($LASTEXITCODE -ne 0) {
        throw "cloudflared tunnel list failed."
    }

    $items = @($json | ConvertFrom-Json)
    return $items | Where-Object {
        $_.name -eq $Name -and (-not $_.deleted_at)
    } | Select-Object -First 1
}

if (-not (Test-Path -LiteralPath $CloudflaredPath)) {
    throw "cloudflared not found: $CloudflaredPath"
}

Write-Step "Checking cloudflared"
& $CloudflaredPath --version

$cloudflaredDir = Join-Path $env:USERPROFILE ".cloudflared"
$originCert = Join-Path $cloudflaredDir "cert.pem"

if (-not (Test-Path -LiteralPath $originCert)) {
    Write-Step "Cloudflare login required"
    Write-Host "A browser window will open. Log in to Cloudflare and choose the domain you want to publish under."
    & $CloudflaredPath tunnel login

    if (-not (Test-Path -LiteralPath $originCert)) {
        throw "Cloudflare login did not create $originCert. Please run the script again after login succeeds."
    }
}

Write-Step "Finding or creating tunnel '$TunnelName'"
$tunnel = Get-TunnelByName -Name $TunnelName
if ($null -eq $tunnel) {
    & $CloudflaredPath tunnel create $TunnelName
    if ($LASTEXITCODE -ne 0) {
        throw "cloudflared tunnel create failed."
    }
    $tunnel = Get-TunnelByName -Name $TunnelName
}

if ($null -eq $tunnel) {
    throw "Could not find tunnel '$TunnelName' after create/list."
}

$tunnelId = [string]$tunnel.id
$credentialsFile = Join-Path $cloudflaredDir "$tunnelId.json"
if (-not (Test-Path -LiteralPath $credentialsFile)) {
    throw "Tunnel credentials file not found: $credentialsFile"
}

Write-Step "Writing config.yml"
$configPath = Join-Path $cloudflaredDir "config.yml"
$credentialsForYaml = $credentialsFile.Replace("\", "/")
$config = @"
tunnel: $tunnelId
credentials-file: $credentialsForYaml

ingress:
  - hostname: $Hostname
    service: $LocalUrl
  - service: http_status:404
"@

New-Item -ItemType Directory -Force -Path $cloudflaredDir | Out-Null
Set-Content -LiteralPath $configPath -Value $config -Encoding UTF8
Write-Host "Config written to: $configPath"

Write-Step "Creating DNS route"
$routeArgs = @("tunnel", "route", "dns", $TunnelName, $Hostname)
if ($OverwriteDns) {
    $routeArgs += "--overwrite-dns"
}
& $CloudflaredPath @routeArgs
if ($LASTEXITCODE -ne 0) {
    throw "cloudflared tunnel route dns failed."
}

if ($InstallService) {
    Write-Step "Installing cloudflared Windows service"
    Write-Host "This may require running PowerShell as Administrator."
    & $CloudflaredPath service install
    if ($LASTEXITCODE -ne 0) {
        throw "cloudflared service install failed. Re-run PowerShell as Administrator, or run without -InstallService."
    }
}

Write-Step "Done"
Write-Host "Fixed public URL: https://$Hostname" -ForegroundColor Green
Write-Host ""
Write-Host "Run the app and tunnel with:"
Write-Host "  powershell -ExecutionPolicy Bypass -File .\cloudflare-tunnel-run.ps1"
Write-Host ""
Write-Host "For long-running use, keep the app process running and install cloudflared as a Windows service."
