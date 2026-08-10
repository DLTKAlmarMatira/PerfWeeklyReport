<#
.SYNOPSIS
    Pulls the "Related" links hanging off each performance PBI and records what
    is on the other end - specifically, which of them are Bugs.

.DESCRIPTION
    Bugs are attached to a PBI with a plain "Related" link. That link type is
    NOT returned by the saved query behind step 2 (which asks only for
    Parent/Child), so bugs never appear in pbi_task_links.csv no matter how
    often the extract is re-run. Rather than require the ADO query to be
    edited, this script walks each PBI's relations through the REST API.

    TWO THINGS THIS SCRIPT IS CAREFUL ABOUT

    1. "Related" does not mean "Bug". On this project the Related links point
       at a mix of Bugs, other PBIs (reciprocal cross-links), and Tasks.
       Counting every Related link as a bug would be wrong, so every target's
       work item type is recorded and the filtering is left to the consumer.

    2. Some targets live in ANOTHER project and are not readable with these
       credentials - the server returns the item as an empty stub, confirming
       it exists while withholding every field. Those are written out with
       Readable=False rather than dropped, so an unreadable bug is visible as
       a gap instead of silently counting as zero.

.PARAMETER CsvDir
    Folder holding pbi_task_links.csv (the PBI list) and receiving the output.
    Default: the "csv" folder next to this script.

.PARAMETER OutputPath
    Default: <CsvDir>\pbi_bug_links.csv

.PARAMETER Pat
    Personal Access Token. If omitted: the shared encrypted cache, then
    $env:AZURE_DEVOPS_PAT, then a prompt (Enter = Windows auth).

.EXAMPLE
    .\8-Get-PbiBugLinks.ps1

.NOTES
    Read-only. Writes nothing to ADO.
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

# $PSScriptRoot is empty inside param() defaults under `powershell.exe -File`,
# so resolve here in the body. See the same note in scripts 6 and 7.
$ScriptDir = $PSScriptRoot
if (-not $ScriptDir)  { $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $CsvDir)     { $CsvDir = Join-Path $ScriptDir 'csv' }
if (-not $OutputPath) { $OutputPath = Join-Path $CsvDir 'pbi_bug_links.csv' }

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

function Get-DisplayName {
    param($Identity)
    if ($null -eq $Identity) { return '' }
    if ($Identity -is [string]) {
        if ($Identity -match 'displayName=([^;}]*)') { return $Matches[1].Trim() }
        return ($Identity -split ' <')[0].Trim()
    }
    return [string]$Identity.displayName
}

try {
    $pbiPath = Join-Path $CsvDir 'pbi_task_links.csv'
    if (-not (Test-Path -LiteralPath $pbiPath)) {
        throw "Not found: $pbiPath`nRun Run-AdoExtracts.bat steps 1-3 first."
    }

    $pbiRows = @(Import-Csv -LiteralPath $pbiPath |
                 Where-Object { $_.'Source System.WorkItemType' -eq 'Product Backlog Item' })
    $pbiTitle = @{}
    foreach ($r in $pbiRows) { $pbiTitle[$r.'Source ID'] = $r.'Source System.Title' }
    $pbiIds = @($pbiTitle.Keys | Sort-Object)
    if ($pbiIds.Count -eq 0) { throw "No PBIs found in $pbiPath" }
    Write-Host ("PBIs to inspect: {0}" -f $pbiIds.Count)

    # --- api-version probe, same approach as the other scripts.
    if (-not $ApiVersion) {
        foreach ($v in '7.1','7.0','6.0','5.1','5.0','4.1','3.2','2.0') {
            try {
                $probe = @{ ids = @($pbiIds[0]); '$expand' = 'relations' } | ConvertTo-Json
                $null = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workitemsbatch?api-version=$v" -Method Post -Body $probe
                $ApiVersion = $v; break
            } catch { }
        }
        if (-not $ApiVersion) { throw 'Could not reach the work item batch endpoint at any api-version.' }
        Write-Host "Using api-version=$ApiVersion"
    }

    # --- 1. collect every Related link off every PBI
    $links = @()          # PBI ID -> target ID
    for ($i = 0; $i -lt $pbiIds.Count; $i += 20) {
        $chunk = @($pbiIds[$i..([Math]::Min($i + 19, $pbiIds.Count - 1))])
        $body = @{ ids = $chunk; '$expand' = 'relations' } | ConvertTo-Json
        $resp = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workitemsbatch?api-version=$ApiVersion" -Method Post -Body $body
        foreach ($item in @($resp.value)) {
            foreach ($rel in @($item.relations)) {
                if ($rel.rel -notlike 'System.LinkTypes.Related*') { continue }
                $targetId = ($rel.url -split '/')[-1]
                if ($targetId -notmatch '^\d+$') { continue }
                # The project GUID sits in the relation URL; it is how we can
                # tell a cross-project link before even trying to read it.
                $projGuid = ''
                if ($rel.url -match '/([0-9a-fA-F-]{36})/_apis/') { $projGuid = $Matches[1] }
                $links += [pscustomobject]@{
                    PbiId = [string]$item.id; TargetId = $targetId; ProjectGuid = $projGuid; Rel = $rel.rel
                }
            }
        }
    }
    Write-Host ("Related links found: {0} across {1} PBI(s)" -f $links.Count,
        (@($links | Select-Object -ExpandProperty PbiId -Unique)).Count)

    if ($links.Count -eq 0) {
        Write-Warning 'No Related links on any PBI - nothing to write.'
        @() | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        return
    }

    # --- 2. resolve each distinct target. Project-scoped batch returns full
    # --- fields for in-project items; anything it refuses is recorded as
    # --- unreadable rather than dropped.
    $targets = @($links | Select-Object -ExpandProperty TargetId -Unique)
    $info = @{}
    for ($i = 0; $i -lt $targets.Count; $i += 100) {
        $chunk = @($targets[$i..([Math]::Min($i + 99, $targets.Count - 1))])
        $body = @{ ids = $chunk
                   fields = @('System.Id','System.WorkItemType','System.Title','System.State','System.AssignedTo')
                   errorPolicy = 'omit' } | ConvertTo-Json
        try {
            $resp = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workitemsbatch?api-version=$ApiVersion" -Method Post -Body $body
            foreach ($item in @($resp.value)) {
                if (-not $item.fields.'System.WorkItemType') { continue }   # empty stub = no read access
                $info[[string]$item.id] = $item.fields
            }
        } catch {
            # errorPolicy=omit is not honoured everywhere; fall back to one
            # call per id so a single unreadable item can't blank the batch.
            foreach ($id in $chunk) {
                try {
                    $one = Invoke-Ado -Uri "$BaseUrl/$Project/_apis/wit/workitems/$id`?api-version=$ApiVersion"
                    if ($one.fields.'System.WorkItemType') { $info[[string]$id] = $one.fields }
                } catch { }
            }
        }
    }

    $rows = foreach ($l in $links) {
        $f = $info[$l.TargetId]
        [pscustomobject][ordered]@{
            'PBI ID'            = $l.PbiId
            'PBI Title'         = $pbiTitle[$l.PbiId]
            'Target ID'         = $l.TargetId
            'Target Type'       = if ($f) { [string]$f.'System.WorkItemType' } else { '' }
            'Target Title'      = if ($f) { [string]$f.'System.Title' } else { '' }
            'Target State'      = if ($f) { [string]$f.'System.State' } else { '' }
            'Target AssignedTo' = if ($f) { Get-DisplayName $f.'System.AssignedTo' } else { '' }
            'Readable'          = if ($f) { 'True' } else { 'False' }
            'Project GUID'      = $l.ProjectGuid
            'Link Type'         = $l.Rel
        }
    }
    $rows = @($rows)

    $rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Wrote {0} ({1} rows)" -f $OutputPath, $rows.Count)

    $bugs      = @($rows | Where-Object { $_.'Target Type' -eq 'Bug' })
    $unreadable= @($rows | Where-Object { $_.Readable -eq 'False' })
    Write-Host ''
    Write-Host ("  Bugs linked          : {0} (across {1} PBI)" -f $bugs.Count,
        (@($bugs | Select-Object -ExpandProperty 'PBI ID' -Unique)).Count)
    Write-Host  '  Related targets by type:'
    $rows | Group-Object 'Target Type' | Sort-Object Count -Descending | ForEach-Object {
        $label = if ($_.Name) { $_.Name } else { '(unreadable)' }
        Write-Host ("     {0,-22} {1}" -f $label, $_.Count)
    }
    if ($unreadable.Count -gt 0) {
        Write-Warning ("{0} related target(s) are in another project and not readable with these credentials - they are recorded with Readable=False and will NOT be counted as bugs." -f $unreadable.Count)
    }
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    exit 1
}
