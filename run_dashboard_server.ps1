$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Config = Join-Path $Root "dashboard.server.json"
$Example = Join-Path $Root "dashboard.server.example.json"

if (-not (Test-Path -LiteralPath $Config)) {
  Copy-Item -LiteralPath $Example -Destination $Config
  Write-Host "Created $Config from example. Edit dashboard.auth_username/auth_password or set TENDER_DASHBOARD_USER/TENDER_DASHBOARD_PASSWORD before publishing."
}

$PythonCandidates = @(
  (Join-Path $Root ".venv\Scripts\python.exe"),
  "python"
)

$Python = $PythonCandidates | Where-Object {
  if ($_ -eq "python") { return $true }
  Test-Path -LiteralPath $_
} | Select-Object -First 1

Push-Location $Root
try {
  & $Python ".\tender_dashboard.py" serve --config ".\dashboard.server.json"
}
finally {
  Pop-Location
}
