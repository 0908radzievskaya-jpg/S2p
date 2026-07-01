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
    [bool]$ApplyRemoteDeletions = $true,
    [string[]]$LocalDeleteRoots = @("reports", "релевантные", "sorted_mail", "_review_ai_tender", "archive"),
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
$localStatePath = Join-Path $env:TEMP "tender-dashboard-state-$stamp.json"
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

$metadataNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
@(
    "03_document_index.json",
    "04_project_card.json",
    "05_tep.json",
    "06_scope.json",
    "07_requirements.json",
    "08_red_flags.json",
    "09_missing_data.json",
    "12_analysis_report.json",
    "14_pir_normative_estimate.json",
    "batch_index.json",
    "links.txt",
    "request_meta.json",
    "run_summary.json",
    "source_message_path.txt"
) | ForEach-Object { [void]$metadataNames.Add($_) }

function Get-LocalDeleteRootPaths {
    $roots = @()
    foreach ($rootName in $LocalDeleteRoots) {
        $path = Join-Path $ProjectRoot $rootName
        if (Test-Path -LiteralPath $path) {
            $roots += (Get-Item -LiteralPath $path -Force).FullName.TrimEnd("\", "/")
        }
    }
    return $roots
}

function Test-PathWithinLocalDeleteRoots {
    param(
        [string]$Path,
        [string[]]$Roots
    )

    $full = [System.IO.Path]::GetFullPath($Path).TrimEnd("\", "/")
    foreach ($root in $Roots) {
        if ($full.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
        $prefix = $root + [System.IO.Path]::DirectorySeparatorChar
        if ($full.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

function Resolve-LocalDeleteCandidate {
    param([string]$Candidate)

    $text = ($Candidate | Out-String).Trim().Trim('"')
    if (-not $text -or $text -match "^[a-z][a-z0-9+.-]*://") {
        return $null
    }
    $text = [Environment]::ExpandEnvironmentVariables($text)
    if (-not [System.IO.Path]::IsPathRooted($text)) {
        $text = Join-Path $ProjectRoot $text
    }
    if (-not (Test-Path -LiteralPath $text)) {
        return $null
    }
    return (Get-Item -LiteralPath $text -Force).FullName
}

function Test-KeepDashboardMetadata {
    param([System.IO.FileInfo]$File)

    if ($metadataNames.Contains($File.Name)) {
        return $true
    }
    if ($File.Extension.Equals(".json", [System.StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }
    return $false
}

function Test-DirectoryEmpty {
    param([string]$Path)

    $child = Get-ChildItem -LiteralPath $Path -Force | Select-Object -First 1
    return $null -eq $child
}

function Remove-LocalDashboardFiles {
    param(
        [string]$TargetPath,
        [string[]]$Roots
    )

    $deleted = 0
    $item = Get-Item -LiteralPath $TargetPath -Force
    if (-not (Test-PathWithinLocalDeleteRoots -Path $item.FullName -Roots $Roots)) {
        return 0
    }

    if ($item.PSIsContainer) {
        $files = Get-ChildItem -LiteralPath $item.FullName -Recurse -File -Force | Sort-Object FullName -Descending
        foreach ($file in $files) {
            if (-not (Test-PathWithinLocalDeleteRoots -Path $file.FullName -Roots $Roots)) {
                continue
            }
            if (Test-KeepDashboardMetadata -File $file) {
                continue
            }
            Remove-Item -LiteralPath $file.FullName -Force
            $deleted++
        }
        $dirs = Get-ChildItem -LiteralPath $item.FullName -Recurse -Directory -Force | Sort-Object FullName -Descending
        foreach ($dir in $dirs) {
            if (-not (Test-PathWithinLocalDeleteRoots -Path $dir.FullName -Roots $Roots)) {
                continue
            }
            if (Test-DirectoryEmpty -Path $dir.FullName) {
                Remove-Item -LiteralPath $dir.FullName -Force
            }
        }
        if (Test-DirectoryEmpty -Path $item.FullName) {
            Remove-Item -LiteralPath $item.FullName -Force
        }
        return $deleted
    }

    if ($item -is [System.IO.FileInfo] -and -not (Test-KeepDashboardMetadata -File $item)) {
        Remove-Item -LiteralPath $item.FullName -Force
        return 1
    }
    return 0
}

function Add-DeletionCandidate {
    param(
        [object]$Value,
        [System.Collections.Generic.HashSet[string]]$Candidates
    )

    if ($null -eq $Value) {
        return
    }
    if ($Value -is [System.Array]) {
        foreach ($item in $Value) {
            Add-DeletionCandidate -Value $item -Candidates $Candidates
        }
        return
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        return
    }
    $text = ($Value | Out-String).Trim()
    if ($text) {
        [void]$Candidates.Add($text)
    }
}

function Invoke-RemoteDeletionSync {
    if (-not $ApplyRemoteDeletions) {
        return
    }

    $roots = Get-LocalDeleteRootPaths
    if ($roots.Count -eq 0) {
        Write-Warning "No local delete roots exist. Skipping remote deletion sync."
        return
    }

    Write-Host "Pulling dashboard deletion decisions from server"
    & scp @sshArgs "${remote}:$RemoteAppDir/.state/dashboard_statuses.json" $localStatePath
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Could not pull remote dashboard state. Local deletion sync skipped."
        return
    }

    $state = Get-Content -LiteralPath $localStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $state.items) {
        return
    }

    $candidates = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($property in $state.items.PSObject.Properties) {
        $record = $property.Value
        if (-not $record -or $record.status -ne "not_relevant" -or -not $record.files_deleted_at) {
            continue
        }
        Add-DeletionCandidate -Value $record.deleted_relative_paths -Candidates $candidates
        Add-DeletionCandidate -Value $record.local_delete_candidates -Candidates $candidates
    }

    $deleted = 0
    foreach ($candidate in $candidates) {
        $target = Resolve-LocalDeleteCandidate -Candidate $candidate
        if (-not $target) {
            continue
        }
        $deleted += Remove-LocalDashboardFiles -TargetPath $target -Roots $roots
    }

    if ($deleted -gt 0) {
        Write-Host "Local dashboard cleanup deleted $deleted file(s)."
    }
}

function New-MinimalPackage {
    New-Item -ItemType Directory -Path (Join-Path $stagingRoot "reports") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $stagingRoot "релевантные") -Force | Out-Null

    $seen = @{}
    $selectedNamePattern = "(?i)(^заявки_|^report-|тз|техническ|задани|описан|объект.*закуп|смет|калькуляц|расче[тт]|аналит|записк)"
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
    Invoke-RemoteDeletionSync

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
    Remove-Item -LiteralPath $localStatePath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
