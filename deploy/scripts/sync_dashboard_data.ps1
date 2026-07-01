[CmdletBinding()]
param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$RemoteHost = "194.113.209.237",
    [string]$RemoteUser = "root",
    [string]$RemoteAppDir = "/opt/tender-dashboard",
    [string]$SshKeyPath = "",
    [int]$KeepBackups = 7,
    [switch]$NoRestart
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

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
$dataRoots = @("reports", "релевантные") | Where-Object {
    Test-Path -LiteralPath (Join-Path $ProjectRoot $_)
}

if ($dataRoots.Count -eq 0) {
    throw "No dashboard data folders found. Expected 'reports' or 'релевантные' under $ProjectRoot."
}

$sshArgs = @("-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new")
if ($SshKeyPath) {
    $sshArgs = @("-i", $SshKeyPath) + $sshArgs
}

$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$archivePath = Join-Path $env:TEMP "tender-dashboard-data-$stamp.tar.gz"
$remoteArchivePath = "/tmp/tender-dashboard-data-$stamp.tar.gz"
$remoteApplyPath = "/tmp/tender-dashboard-apply-data-$stamp.sh"
$localApplyPath = Join-Path $env:TEMP "tender-dashboard-apply-data-$stamp.sh"
$remote = "${RemoteUser}@${RemoteHost}"
$restartFlag = if ($NoRestart) { "0" } else { "1" }

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

http_code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 10 http://127.0.0.1:8765/ || true)"
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
    Write-Host "Packing dashboard data: $($dataRoots -join ', ')"
    & tar -czf $archivePath -C $ProjectRoot @dataRoots
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
}
