<#
.SYNOPSIS
    Fetches the latest discussion entry for each PBI and Task in the pipeline.

.DESCRIPTION
    The weekly HTML report can display the most recent comment left on each work
    item. This script walks the work item updates API to find it. It reads the
    PBI/Task IDs from pbi_task_links.csv (produced by step 2) and writes one
    row per work item that has at least one discussion entry.

    Only the most-recent update that carries a System.History value is captured.
    Items with no discussion at all are omitted (the report simply shows no icon
    for those rows).

    Unreadable items (e.g. cross-project work items, permission gaps) are
    silently skipped rather than hard-failing the pipeline.

.PARAMETER CsvDir
    Folder holding pbi_task_links.csv and receiving the output.
    Default: the "csv" folder next to this script.

.PARAMETER OutputPath
    Default: <CsvDir>\workitem_comments.csv

.PARAMETER Pat
    Personal Access Token. If omitted: the shared encrypted cache, then
    $env:AZURE_DEVOPS_PAT, then a prompt (Enter = Windows auth).

.EXAMPLE
    .\9-Get-WorkItemComments.ps1

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
if (-not $OutputPath) { $OutputPath = Join-Path $CsvDir 'workitem_comments.csv' }

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
    $pbiIds = @($rows |
                Where-Object { $_.'Source System.WorkItemType' -eq 'Product Backlog Item' -and $_.'Source ID' } |
                Select-Object -ExpandProperty 'Source ID' | Sort-Object -Unique)
    $taskIds = @($rows |
                 Where-Object { $_.'Target System.WorkItemType' -eq 'Task' -and $_.'Target ID' } |
                 Select-Object -ExpandProperty 'Target ID' | Sort-Object -Unique)
    $allIds = @(($pbiIds + $taskIds) | Sort-Object -Unique)

    if ($allIds.Count -eq 0) { throw "No PBI or Task IDs found in $pbiTaskPath" }
    Write-Host ("Work items to check: {0} PBI(s) + {1} Task(s) = {2} total" -f `
        $pbiIds.Count, $taskIds.Count, $allIds.Count)

    # API version probe against the updates endpoint.
    if (-not $ApiVersion) {
        foreach ($v in '7.1','7.0','6.0','5.1','5.0','4.1','3.2','2.0') {
            try {
                $null = Invoke-Ado -Uri ("$BaseUrl/$Project/_apis/wit/workItems/{0}/updates?`$top=1&api-version=$v" -f $allIds[0])
                $ApiVersion = $v; break
            } catch { }
        }
        if (-not $ApiVersion) { throw 'Could not reach the work item updates endpoint at any api-version.' }
        Write-Host "Using api-version=$ApiVersion"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $found = 0; $skipped = 0

    foreach ($id in $allIds) {
        try {
            # Fetch all updates for this work item. Updates come back oldest-first.
            # We walk them in reverse to find the most-recent discussion entry.
            # Most work items have well under 100 updates so fetching without
            # $top is acceptable; the server's own default cap applies.
            $resp = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workItems/$id/updates?api-version=$ApiVersion"
            $withHistory = @($resp.value | Where-Object {
                $_.fields -and
                $_.fields.'System.History' -and
                -not [string]::IsNullOrWhiteSpace($_.fields.'System.History'.newValue)
            })

            if ($withHistory.Count -eq 0) { $skipped++; continue }

            $latest = $withHistory[-1]   # last = most recent
            $commentHtml = [string]$latest.fields.'System.History'.newValue
            $author = if ($latest.revisedBy -and $latest.revisedBy.displayName) {
                          [string]$latest.revisedBy.displayName
                      } else { '' }
            $date = if ($latest.revisedDate) { [string]$latest.revisedDate } else { '' }

            $results.Add([pscustomobject][ordered]@{
                WorkItemId  = $id
                Author      = $author
                Date        = $date
                CommentHtml = $commentHtml
            })
            $found++
        } catch {
            Write-Warning ("Skipped work item {0}: {1}" -f $id, $_.Exception.Message)
            $skipped++
        }
    }

    Write-Host ''
    Write-Host ("Discussion entries found : {0}" -f $found)
    Write-Host ("No discussion / skipped  : {0}" -f $skipped)

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Wrote {0} ({1} rows)" -f $OutputPath, $results.Count)
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
