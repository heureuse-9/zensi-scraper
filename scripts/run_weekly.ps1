param(
    [string]$ConfigPath = ""
)

$ErrorActionPreference = "Stop"
$Root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")
Set-Location -LiteralPath $Root

if (-not $ConfigPath) {
    $ConfigPath = Join-Path $Root "config.json"
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Missing config file: $ConfigPath. Copy config.example.json to config.json and edit paths first."
}

$VenvPython = Join-Path $Root ".venv\Scripts\python.exe"
if (Test-Path -LiteralPath $VenvPython) {
    & $VenvPython -m zensi_scraper run-weekly --config $ConfigPath
} else {
    & py -m zensi_scraper run-weekly --config $ConfigPath
}
