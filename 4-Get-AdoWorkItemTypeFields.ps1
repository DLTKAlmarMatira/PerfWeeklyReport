<#
.SYNOPSIS
    Lists every field defined on one or more work item types (e.g. Task,
    Product Backlog Item) in this ADO/TFS project, and flags any that look
    like a "Blocked" or "Target/Due Date" field.

.DESCRIPTION
    One-time check to confirm whether custom fields we'd want for weekly
    reporting (a Blocked flag, a Target/Due Date) actually exist on this
    org's work item types before wiring them into the extraction scripts
    via -ExtraFields. It doesn't change or create anything.

    v2: the first version of this script hardcoded api-version=7.1 and got
    an empty field list back for every type - which, given the same server
    needed API-version auto-detection for the Test Plan and Query scripts,
    is almost certainly a version mismatch rather than proof the fields
    don't exist. This version probes several API versions and both known
    "list fields for a work item type" endpoint shapes, and if every combo
    comes back empty, it dumps the raw response so we can see what's
    actually happening instead of guessing.

    Shares the same encrypted PAT cache as the other scripts.

.PARAMETER Organization
    Base URL of the ADO/TFS server + collection.
    Default: https://tfs.deltek.com/tfs/Deltek

.PARAMETER Project
    Project name. Default: QEAutomation

.PARAMETER WorkItemTypes
    One or more work item type names to inspect.
    Default: "Task","Product Backlog Item"  (confirmed exact names from the
    existing CSV exports' System.WorkItemType values)

.PARAMETER OutputPath
    Where to write the full field list as CSV. Default: workitemtype_fields.csv

.PARAMETER Pat
.PARAMETER ApiVersion
    Optional. If specified, skips probing and uses this version only.
.PARAMETER ResetPat
#>

[CmdletBinding()]
param(
    [string]$Organization = "https://tfs.deltek.com/tfs/Deltek",
    [string]$Project = "QEAutomation",
    [string[]]$WorkItemTypes = @("Task", "Product Backlog Item"),
    [string]$OutputPath = "workitemtype_fields.csv",
    [string]$Pat,
    [string]$ApiVersion,
    [switch]$ResetPat
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"
$BaseUrl = $Organization.TrimEnd('/')

$CredentialDir  = Join-Path $env:LOCALAPPDATA "AdoTestPlanExtractor"
$CredentialFile = Join-Path $CredentialDir "pat.dat"

function Get-SavedPat {
    if (-not (Test-Path $CredentialFile)) { return $null }
    try {
        $secure = Get-Content -Path $CredentialFile -ErrorAction Stop | ConvertTo-SecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch {
        Write-Warning "Could not read saved PAT. Ignoring it."
        return $null
    }
}

function Save-Pat {
    param([string]$PlainPat)
    if (-not (Test-Path $CredentialDir)) { New-Item -ItemType Directory -Path $CredentialDir -Force | Out-Null }
    $PlainPat | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -Path $CredentialFile
    Write-Host "PAT saved (encrypted, this Windows account only) to $CredentialFile"
}

function Get-SecretOrPrompt {
    param([string]$Value, [string]$Label)
    if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $saved = Get-SavedPat
    if (-not [string]::IsNullOrWhiteSpace($saved)) {
        Write-Host "Using saved PAT from $CredentialFile (run with -ResetPat to clear it)."
        return $saved
    }
    if (-not [string]::IsNullOrWhiteSpace($env:AZURE_DEVOPS_PAT)) { return $env:AZURE_DEVOPS_PAT }
    $secure = Read-Host -Prompt $Label -AsSecureString
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    if (-not [string]::IsNullOrWhiteSpace($plain)) { Save-Pat -PlainPat $plain }
    return $plain
}

if ($ResetPat -and (Test-Path $CredentialFile)) {
    Remove-Item -Path $CredentialFile -Force
    Write-Host "Cleared saved PAT. You will be prompted for a new one."
}
$Pat = Get-SecretOrPrompt -Value $Pat -Label "Personal Access Token (press Enter to use your Windows login instead)"

$UseWindowsAuth = [string]::IsNullOrWhiteSpace($Pat)
if ($UseWindowsAuth) {
    Write-Host "No PAT provided - using your current Windows login (NTLM/Kerberos)."
    $AuthHeader = @{}
} else {
    $AuthHeader = @{ Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat")) }
}

function Invoke-AdoGetRaw {
    # Returns $null on failure instead of throwing, so probing can keep going.
    param([string]$Uri, [string]$Version)
    $fullUri = $Uri + "?" + "api-version=$Version"
    try {
        if ($UseWindowsAuth) {
            $resp = Invoke-WebRequest -Uri $fullUri -Method Get -UseDefaultCredentials -UseBasicParsing
        } else {
            $resp = Invoke-WebRequest -Uri $fullUri -Method Get -Headers $AuthHeader -UseBasicParsing
        }
        return @{ Uri = $fullUri; StatusCode = $resp.StatusCode; Content = $resp.Content }
    } catch {
        $status = $null
        if ($_.Exception.Response) { $status = [int]$_.Exception.Response.StatusCode }
        return @{ Uri = $fullUri; StatusCode = $status; Content = $null; Error = $_.Exception.Message }
    }
}

$CandidateVersions = if ($ApiVersion) { @($ApiVersion) } else { @("7.1", "7.0", "6.0", "5.1", "5.0", "4.1", "3.2", "3.0", "2.3", "2.2", "2.1", "2.0") }

function Try-GetFields {
    param([string]$Wit)
    $encodedType = [Uri]::EscapeDataString($Wit)
    $encodedProject = [Uri]::EscapeDataString($Project)

    $endpointA = "$BaseUrl/$encodedProject/_apis/wit/workitemtypes/$encodedType"
    $endpointB = "$BaseUrl/$encodedProject/_apis/wit/workitemtypes/$encodedType/fields"

    $attempts = @()
    foreach ($v in $CandidateVersions) {
        foreach ($ep in @($endpointA, $endpointB)) {
            $raw = Invoke-AdoGetRaw -Uri $ep -Version $v
            $attempts += $raw
            Write-Host "  [$v] GET $ep -> HTTP $($raw.StatusCode)"
            if ($raw.Content) {
                try {
                    $obj = $raw.Content | ConvertFrom-Json
                } catch {
                    continue
                }
                $fieldsArr = $null
                if ($obj.PSObject.Properties.Name -contains 'fields' -and $obj.fields) { $fieldsArr = $obj.fields }
                elseif ($obj.PSObject.Properties.Name -contains 'value' -and $obj.value) { $fieldsArr = $obj.value }
                if ($fieldsArr -and $fieldsArr.Count -gt 0) {
                    Write-Host "  SUCCESS with api-version=$v at $ep ($($fieldsArr.Count) fields)"
                    return @{ Fields = $fieldsArr; Attempts = $attempts }
                }
            }
        }
    }
    return @{ Fields = @(); Attempts = $attempts }
}

$rows = @()
$matchesFound = @()

foreach ($wit in $WorkItemTypes) {
    Write-Host "----------------------------------------"
    Write-Host "Work item type: $wit"
    Write-Host "----------------------------------------"

    $result = Try-GetFields -Wit $wit
    $fields = $result.Fields

    if (-not $fields -or $fields.Count -eq 0) {
        Write-Warning "Still no fields for '$wit' after probing $($CandidateVersions.Count) API versions x 2 endpoint shapes."
        $lastGood = $result.Attempts | Where-Object { $_.StatusCode -eq 200 } | Select-Object -Last 1
        if ($lastGood) {
            Write-Host "  Last HTTP 200 response body (first 1000 chars) for manual inspection:"
            Write-Host "  $($lastGood.Content.Substring(0, [Math]::Min(1000, $lastGood.Content.Length)))"
        } else {
            $lastAttempt = $result.Attempts | Select-Object -Last 1
            Write-Host "  No HTTP 200 responses at all. Last attempt: HTTP $($lastAttempt.StatusCode) - $($lastAttempt.Error)"
        }
        continue
    }

    Write-Host "  $($fields.Count) field(s) found."
    foreach ($f in $fields) {
        $rows += [PSCustomObject]@{
            WorkItemType   = $wit
            ReferenceName  = $f.referenceName
            Name           = $f.name
            AlwaysRequired = $f.alwaysRequired
        }
        if ($f.name -match 'block|target|due|deadline|impediment' -or $f.referenceName -match 'block|target|due|deadline|impediment') {
            $matchesFound += [PSCustomObject]@{ WorkItemType = $wit; ReferenceName = $f.referenceName; Name = $f.name }
        }
    }
}

$rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host ""
Write-Host "Full field list written to: $OutputPath ($($rows.Count) rows)"
Write-Host ""

if ($matchesFound.Count -gt 0) {
    Write-Host "========================================"
    Write-Host "Possible Blocked / Target-Due-Date fields found:"
    Write-Host "========================================"
    $matchesFound | Format-Table -AutoSize
} elseif ($rows.Count -gt 0) {
    Write-Host "========================================"
    Write-Host "Got a real field list this time, but nothing matched 'block', 'target',"
    Write-Host "'due', 'deadline', or 'impediment'. See $OutputPath for the full list -"
    Write-Host "this org's Task/PBI types likely just don't have one."
    Write-Host "========================================"
} else {
    Write-Host "========================================"
    Write-Host "Could not retrieve a field list at all for any type/version combo."
    Write-Host "Paste the console output above back so we can see what's actually"
    Write-Host "coming back from the server (auth issue, wrong route, etc.)."
    Write-Host "========================================"
}
