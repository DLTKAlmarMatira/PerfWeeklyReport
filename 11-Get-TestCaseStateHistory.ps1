<#
.SYNOPSIS
    Fetches System.State change history for every Test Case linked to a scripting task.

.DESCRIPTION
    Walks the work item updates API for each Test Case found in
    task_tests_link_results.csv and records any revision where System.State
    changed value (old -> new).
    The HTML report uses the most-recent change per test case to show, per
    scripting task, how many of its linked test cases transitioned to Ready or
    Design state.

    Unreadable items are silently skipped rather than hard-failing the pipeline.

.PARAMETER CsvDir
    Folder holding task_tests_link_results.csv and receiving the output.
    Default: the "csv" folder next to this script.

.PARAMETER OutputPath
    Default: <CsvDir>\workitem_tc_state_history.csv

.PARAMETER Pat
    Personal Access Token. If omitted: the shared encrypted cache, then
    $env:AZURE_DEVOPS_PAT, then a prompt (Enter = Windows auth).

.EXAMPLE
    .\11-Get-TestCaseStateHistory.ps1

.NOTES
    Read-only. Writes nothing to ADO.
    Requires step 3 (3-Get-TaskTestsLinkResults.ps1) to have run first.
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
if (-not $OutputPath) { $OutputPath = Join-Path $CsvDir 'workitem_tc_state_history.csv' }

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
    $taskTestsPath = Join-Path $CsvDir 'task_tests_link_results.csv'
    if (-not (Test-Path -LiteralPath $taskTestsPath)) {
        throw "Not found: $taskTestsPath`nRun Run-AdoExtracts.bat steps 1-3 first."
    }

    $rows = @(Import-Csv -LiteralPath $taskTestsPath)
    $tcIds = @($rows |
               Where-Object { $_.'Target System.WorkItemType' -eq 'Test Case' -and $_.'Target ID' } |
               Select-Object -ExpandProperty 'Target ID' | Sort-Object -Unique)

    if ($tcIds.Count -eq 0) { throw "No Test Case IDs found in $taskTestsPath" }
    Write-Host ("Test cases to check: {0}" -f $tcIds.Count)

    # API version probe against the updates endpoint.
    if (-not $ApiVersion) {
        foreach ($v in '7.1','7.0','6.0','5.1','5.0','4.1','3.2','2.0') {
            try {
                $null = Invoke-Ado -Uri ("$BaseUrl/$Project/_apis/wit/workItems/{0}/updates?`$top=1&api-version=$v" -f $tcIds[0])
                $ApiVersion = $v; break
            } catch { }
        }
        if (-not $ApiVersion) { throw 'Could not reach the work item updates endpoint at any api-version.' }
        Write-Host "Using api-version=$ApiVersion"
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $found = 0; $skipped = 0

    foreach ($id in $tcIds) {
        try {
            $resp = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workItems/$id/updates?api-version=$ApiVersion"
            $withState = @($resp.value | Where-Object {
                $_.fields -and
                $_.fields.'System.State'
            })

            if ($withState.Count -eq 0) { $skipped++; continue }

            # Emit one row per change, oldest first. Script 7 keeps only the
            # last row per test case (most-recent change) to show in the report.
            #
            # Use System.ChangedDate.newValue for the date rather than
            # revisedDate. The updates API sets revisedDate to 9999-01-01 for
            # the current live revision (not yet superseded), while
            # System.ChangedDate.newValue always holds the real timestamp of
            # the revision itself. Both fields appear in the same revision entry.
            foreach ($entry in $withState) {
                $st = $entry.fields.'System.State'
                $cd = $entry.fields.'System.ChangedDate'
                $revDate = if ($cd -and $cd.newValue) {
                    [string]$cd.newValue
                } elseif ($entry.revisedDate -and ([string]$entry.revisedDate -notlike '9999*')) {
                    [string]$entry.revisedDate
                } else { '' }
                $results.Add([pscustomobject][ordered]@{
                    TestCaseId   = $id
                    RevisionDate = $revDate
                    OldState     = if ($null -ne $st.oldValue) { [string]$st.oldValue } else { '' }
                    NewState     = if ($null -ne $st.newValue) { [string]$st.newValue } else { '' }
                })
            }
            $found++
        } catch {
            Write-Warning ("Skipped test case {0}: {1}" -f $id, $_.Exception.Message)
            $skipped++
        }
    }

    Write-Host ''
    Write-Host ("Test cases with state changes: {0}" -f $found)
    Write-Host ("Total change rows written     : {0}" -f $results.Count)
    Write-Host ("No changes / skipped          : {0}" -f $skipped)

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Wrote {0} ({1} rows)" -f $OutputPath, $results.Count)
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
