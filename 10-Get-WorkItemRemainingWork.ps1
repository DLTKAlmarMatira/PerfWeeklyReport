<#
.SYNOPSIS
    Fetches Remaining Work change history for every Task in the pipeline.

.DESCRIPTION
    Walks the work item updates API for each Task and records any revision
    where Microsoft.VSTS.Scheduling.RemainingWork changed value (old -> new).
    The HTML report uses the most-recent change per task to show scripting
    task progress as "Xh -> Yh" or "no changes".

    Unreadable items are silently skipped rather than hard-failing the pipeline.

.PARAMETER CsvDir
    Folder holding pbi_task_links.csv and receiving the output.
    Default: the "csv" folder next to this script.

.PARAMETER OutputPath
    Default: <CsvDir>\workitem_remwork_history.csv

.PARAMETER Pat
    Personal Access Token. If omitted: the shared encrypted cache, then
    $env:AZURE_DEVOPS_PAT, then a prompt (Enter = Windows auth).

.EXAMPLE
    .\10-Get-WorkItemRemainingWork.ps1

.NOTES
    Read-only. Writes nothing to ADO.
    Requires step 2 (2-Get-AdoQueryResults.ps1) to have run first.
#>

[CmdletBinding()]
param(
    [string]$CsvDir,
    [string]$OutputPath,
    [string]$Organization = 'https://tfs.deltek.com/tfs/Deltek',
    [string]$Project = 'QEAutomation',
    [string]$Pat,
    [string]$ApiVersion
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = 'Stop'

$ScriptDir = $PSScriptRoot
if (-not $ScriptDir)  { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $CsvDir)     { $CsvDir = Join-Path $ScriptDir 'csv' }
if (-not $OutputPath) { $OutputPath = Join-Path $CsvDir 'workitem_remwork_history.csv' }

$BaseUrl = $Organization.TrimEnd('/')
$CredentialFile = Join-Path $env:LOCALAPPDATA 'AdoTestPlanExtractor\pat.dat'

function Get-SavedPat {
    if (-not (Test-Path $CredentialFile)) { return $null }
    try {
        $secure = Get-Content -Path $CredentialFile -ErrorAction Stop | ConvertTo-SecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch { Write-Warning 'Could not read the saved PAT. Ignoring it.'; return $null }
}

if (-not $Pat) { $Pat = Get-SavedPat }
if (-not $Pat) { $Pat = $env:AZURE_DEVOPS_PAT }
$UseWindowsAuth = [string]::IsNullOrWhiteSpace($Pat)
$AuthHeader = if ($UseWindowsAuth) { @{} } else {
    @{ Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat")) }
}

function Invoke-Ado {
    param([string]$Uri, [string]$Method = 'Get', [string]$Body)
    $p = @{ Uri = $Uri; Method = $Method; ErrorAction = 'Stop' }
    if ($UseWindowsAuth) { $p['UseDefaultCredentials'] = $true } else { $p['Headers'] = $AuthHeader }
    if ($Body) { $p['Body'] = $Body; $p['ContentType'] = 'application/json' }
    Invoke-RestMethod @p
}

try {
    $pbiTaskPath = Join-Path $CsvDir 'pbi_task_links.csv'
    if (-not (Test-Path -LiteralPath $pbiTaskPath)) {
        throw "Not found: $pbiTaskPath`nRun Run-AdoExtracts.bat steps 1-3 first."
    }

    $rows = @(Import-Csv -LiteralPath $pbiTaskPath)
    $taskIds = @($rows |
                 Where-Object { $_.'Target System.WorkItemType' -eq 'Task' -and $_.'Target ID' } |
                 Select-Object -ExpandProperty 'Target ID' | Sort-Object -Unique)

    if ($taskIds.Count -eq 0) { throw "No Task IDs found in $pbiTaskPath" }
    Write-Host ("Tasks to check: {0}" -f $taskIds.Count)

    # API version probe against the updates endpoint.
    if (-not $ApiVersion) {
        foreach ($v in '7.1','7.0','6.0','5.1','5.0','4.1','3.2','2.0') {
            try {
                $null = Invoke-Ado -Uri ("$BaseUrl/$Project/_apis/wit/workItems/{0}/updates?`$top=1&api-version=$v" -f $taskIds[0])
                $ApiVersion = $v; break
            } catch { }
        }
        if (-not $ApiVersion) { throw 'Could not reach the work item updates endpoint at any api-version.' }
        Write-Host "Using api-version=$ApiVersion"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $found = 0; $skipped = 0

    foreach ($id in $taskIds) {
        try {
            $resp = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workItems/$id/updates?api-version=$ApiVersion"
            $withRw = @($resp.value | Where-Object {
                $_.fields -and
                $_.fields.'Microsoft.VSTS.Scheduling.RemainingWork'
            })

            if ($withRw.Count -eq 0) { $skipped++; continue }

            # Emit one row per change, oldest first. Script 7 keeps only the
            # last row per task (most-recent change) to show in the report.
            foreach ($entry in $withRw) {
                $rw = $entry.fields.'Microsoft.VSTS.Scheduling.RemainingWork'
                $results.Add([pscustomobject][ordered]@{
                    WorkItemId   = $id
                    RevisionDate = if ($entry.revisedDate) { [string]$entry.revisedDate } else { '' }
                    OldValue     = if ($null -ne $rw.oldValue) { [string]$rw.oldValue } else { '' }
                    NewValue     = if ($null -ne $rw.newValue) { [string]$rw.newValue } else { '' }
                })
            }
            $found++
        } catch {
            Write-Warning ("Skipped work item {0}: {1}" -f $id, $_.Exception.Message)
            $skipped++
        }
    }

    Write-Host ''
    Write-Host ("Tasks with RemainingWork changes: {0}" -f $found)
    Write-Host ("Total change rows written        : {0}" -f $results.Count)
    Write-Host ("No changes / skipped             : {0}" -f $skipped)

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Wrote {0} ({1} rows)" -f $OutputPath, $results.Count)
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
