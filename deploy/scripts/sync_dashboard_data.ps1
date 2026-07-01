[CmdletBinding()]
param(
    [string]$ProjectRoot = "",
    [string]$RemoteHost = "194.113.209.237",
    [string]$RemoteUser = "root",
    [string]$RemoteAppDir = "/opt/tender-dashboard",
    [string]$SshKeyPath = "",
    [ValidateSet("Minimal", "Full")]
    [string]$Mode = "Minimal",
    [int]$KeepBackups = 7,
    [switch]$NoRestart
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not $ProjectRoot) {
    $ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
}

function Require-Command {
    param([string]$Name)
    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "Required command '$Name' was not found in PATH."
    }
}

Require-Command ssh
Require-Command scp
Require-Command tar

$ProjectRoot = (Resolve-Path $ProjectRoot).Path
$sshArgs = @("-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new")
if ($SshKeyPath) {
    $sshArgs = @("-i", $SshKeyPath) + $sshArgs
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$archivePath = Join-Path $env:TEMP "tender-dashboard-data-$stamp.tar.gz"
$remoteArchivePath = "/tmp/tender-dashboard-data-$stamp.tar.gz"
$remoteApplyPath = "/tmp/tender-dashboard-apply-data-$stamp.sh"
$localApplyPath = Join-Path $env:TEMP "tender-dashboard-apply-data-$stamp.sh"
$stagingRoot = Join-Path $env:TEMP "tender-dashboard-data-staging-$stamp"
$remote = "${RemoteUser}@${RemoteHost}"
$restartFlag = if ($NoRestart) { "0" } else { "1" }

function Get-RelativePath {
    param([string]$Path)
    $root = (Get-Item -LiteralPath $ProjectRoot).FullName.TrimEnd("\", "/")
    $full = (Get-Item -LiteralPath $Path).FullName
    return $full.Substring($root.Length).TrimStart("\", "/")
}

function Add-StagedFile {
    param(
        [System.IO.FileInfo]$File,
        [hashtable]$Seen
    )

    $relative = Get-RelativePath -Path $File.FullName
    if ($Seen.ContainsKey($relative)) {
        return
    }
    $target = Join-Path $stagingRoot $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $File.FullName -Destination $target -Force
    $Seen[$relative] = $true
}

function New-MinimalPackage {
    New-Item -ItemType Directory -Path (Join-Path $stagingRoot "reports") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $stagingRoot "релевантные") -Force | Out-Null

    $seen = @{}
    $selectedNamePattern = "(?i)(^заявки_|^report-|тз|техническ|задани|смет|калькуляц|расче[тт]|аналит|записк)"
    $selectedExtensions = @(".pdf", ".doc", ".docx", ".xls", ".xlsx", ".txt")
    $analysisMetaNames = @(
        "03_document_index.json",
        "04_project_card.json",
        "05_tep.json",
        "06_scope.json",
        "07_requirements.json",
        "08_red_flags.json",
        "09_missing_data.json",
        "12_analysis_report.json",
        "14_pir_normative_estimate.json",
        "request_meta.json",
        "links.txt",
        "source_message_path.txt"
    )

    foreach ($rootName in @("reports", "релевантные")) {
        $rootPath = Join-Path $ProjectRoot $rootName
        if (-not (Test-Path -LiteralPath $rootPath)) {
            continue
        }

        Get-ChildItem -LiteralPath $rootPath -Recurse -File | ForEach-Object {
            $relative = Get-RelativePath -Path $_.FullName
            $extension = $_.Extension.ToLowerInvariant()
            $isDashboardWorkbook = $relative -match "^[\\/]?reports[\\/](Заявки_|report-).+\.xlsx$"
            $isAnalysisMeta = ($relative -match "^[\\/]?reports[\\/]analysis_") -and ($analysisMetaNames -contains $_.Name)
            $isSelectedDocument = ($selectedExtensions -contains $extension) -and ($_.Name -match $selectedNamePattern)

            if ($isDashboardWorkbook -or $isAnalysisMeta -or $isSelectedDocument) {
                Add-StagedFile -File $_ -Seen $seen
            }
        }
    }

    if ($seen.Count -eq 0) {
        throw "No minimal dashboard files found. Expected reports/Заявки_*.xlsx, analysis JSON, TZ, estimate, or analytic files."
    }

    return @("reports", "релевантные")
}

$remoteScript = @'
#!/usr/bin/env bash
set -Eeuo pipefail

ARCHIVE="$1"
APP_DIR="$2"
KEEP_BACKUPS="$3"
RESTART="$4"
STAGING="/tmp/tender-dashboard-data-$$"
BACKUP_ROOT="/var/backups/tender-dashboard-data"

cleanup() {
    rm -rf "$STAGING"
}
trap cleanup EXIT

mkdir -p "$STAGING" "$BACKUP_ROOT" "$APP_DIR"
tar -xzf "$ARCHIVE" -C "$STAGING"

stamp="$(date +%Y%m%d-%H%M%S)"
backup="$BACKUP_ROOT/$stamp"
mkdir -p "$backup"

changed=0
for name in reports релевантные; do
    if [ -e "$STAGING/$name" ]; then
        if [ -e "$APP_DIR/$name" ]; then
            mv "$APP_DIR/$name" "$backup/$name"
        fi
        mv "$STAGING/$name" "$APP_DIR/$name"
        changed=1
    fi
done

if [ "$changed" = "0" ]; then
    echo "No recognized data folders found in uploaded archive." >&2
    exit 1
fi

chown -R tender-dashboard:tender-dashboard "$APP_DIR/reports" "$APP_DIR/релевантные" 2>/dev/null || true
chmod -R u+rwX,go+rX,go-w "$APP_DIR/reports" "$APP_DIR/релевантные" 2>/dev/null || true
find "$BACKUP_ROOT" -mindepth 1 -maxdepth 1 -type d | sort -r | tail -n +"$((KEEP_BACKUPS + 1))" | xargs -r rm -rf

if [ "$RESTART" = "1" ]; then
    systemctl restart tender-dashboard.service
fi

http_code="000"
for attempt in $(seq 1 30); do
    http_code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 10 http://127.0.0.1:8765/ || true)"
    case "$http_code" in
        200|401)
            break
            ;;
    esac
    sleep 2
done

case "$http_code" in
    200|401)
        echo "Dashboard health OK: HTTP $http_code"
        ;;
    *)
        echo "Dashboard health FAILED: HTTP ${http_code:-000}" >&2
        systemctl status tender-dashboard.service --no-pager || true
        exit 1
        ;;
esac

rm -f "$ARCHIVE" "$0"
echo "Dashboard data sync complete."
'@

try {
    if ($Mode -eq "Full") {
        $packageRoot = $ProjectRoot
        $tarRoots = @("reports", "релевантные") | Where-Object {
            Test-Path -LiteralPath (Join-Path $ProjectRoot $_)
        }
        if ($tarRoots.Count -eq 0) {
            throw "No dashboard data folders found. Expected 'reports' or 'релевантные' under $ProjectRoot."
        }
    }
    else {
        $packageRoot = $stagingRoot
        $tarRoots = New-MinimalPackage
    }

    Write-Host "Packing dashboard data ($Mode): $($tarRoots -join ', ')"
    & tar -czf $archivePath -C $packageRoot @tarRoots
    if ($LASTEXITCODE -ne 0) {
        throw "tar failed with exit code $LASTEXITCODE"
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($localApplyPath, $remoteScript, $utf8NoBom)

    Write-Host "Uploading archive to ${remote}:$remoteArchivePath"
    & scp @sshArgs $archivePath "${remote}:$remoteArchivePath"
    if ($LASTEXITCODE -ne 0) {
        throw "scp archive upload failed with exit code $LASTEXITCODE"
    }

    Write-Host "Uploading remote apply script"
    & scp @sshArgs $localApplyPath "${remote}:$remoteApplyPath"
    if ($LASTEXITCODE -ne 0) {
        throw "scp apply script upload failed with exit code $LASTEXITCODE"
    }

    Write-Host "Applying data on server"
    & ssh @sshArgs $remote "bash '$remoteApplyPath' '$remoteArchivePath' '$RemoteAppDir' '$KeepBackups' '$restartFlag'"
    if ($LASTEXITCODE -ne 0) {
        throw "remote apply failed with exit code $LASTEXITCODE"
    }
}
finally {
    Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $localApplyPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
