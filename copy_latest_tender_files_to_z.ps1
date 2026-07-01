param(
    [string]$Destination = "Z:\Тендеры",
    [string]$Excel,
    [ValidateSet("docs", "full")]
    [string]$Mode = "docs",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

$python = "C:\Program Files\Blender Foundation\Blender 5.1\5.1\python\bin\python.exe"
$script = Join-Path $PSScriptRoot "copy_tender_files_to_drive.py"

if (-not (Test-Path -LiteralPath $python)) {
    throw "Python not found: $python"
}

$arguments = @($script, "--dest", $Destination, "--mode", $Mode)
if ($Excel) {
    $arguments += @("--excel", $Excel)
}
if ($DryRun) {
    $arguments += "--dry-run"
}

& $python @arguments
