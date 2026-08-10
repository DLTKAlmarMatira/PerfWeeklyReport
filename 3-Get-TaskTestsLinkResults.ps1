<#
.SYNOPSIS
    Runs your existing saved "Tests" link query (Task --Tests--> Test Case /
    Test Plan) and exports the results to CSV or Excel.

.DESCRIPTION
    This is a dedicated copy of Get-AdoQueryResults.ps1, defaulted to your own
    saved query rather than a generic one you have to paste each time. Same
    engine underneath: figures out org/project/query ID from the URL, pulls
    the query's configured columns, runs it, and - since this is a
    "Work items and direct links" query - writes one row per link (Source =
    Task, Target = Test Case/Test Plan, plus the Link Type).

    Shares the same encrypted PAT cache as the other scripts in this folder.

.PARAMETER QueryUrl
    Defaults to your saved Task Tests Link query. Override with -QueryUrl if
    you ever want to point this at a different saved query.

.PARAMETER OutputPath
    Output file path. .csv or .xlsx decided by extension.
    Default: task_tests_link_results.csv

.PARAMETER ExtraFields
    Optional. Extra field reference names to fetch beyond the query's own
    configured columns, e.g. -ExtraFields "System.IterationPath"

.PARAMETER Pat
    Personal Access Token. If omitted: checks the saved encrypted PAT, then
    $env:AZURE_DEVOPS_PAT, then prompts (press Enter to fall back to your
    Windows login instead).

.PARAMETER ApiVersion
    REST API version. If not specified, auto-detected against this query's
    own endpoint.

.PARAMETER ResetPat
    Clears the saved encrypted PAT so you're prompted for a fresh one.

.EXAMPLE
    .\Get-TaskTestsLinkResults.ps1

.EXAMPLE
    .\Get-TaskTestsLinkResults.ps1 -OutputPath "C:\CLAUDEPROJ\perfprocess\ADO\csv\task_tests_links.csv"
#>

[CmdletBinding()]
param(
    [string]$QueryUrl = "https://tfs.deltek.com/tfs/Deltek/QEAutomation/_queries/query-edit/629fda22-d2ec-4af7-9279-15b8b31f5fdb/?action=new",
    [string]$OutputPath = "task_tests_link_results.csv",
    [string[]]$ExtraFields = @(),
    [string]$Pat,
    [string]$ApiVersion = "7.1",
    [switch]$ResetPat
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = "Stop"
$ApiVersionWasSpecified = $PSBoundParameters.ContainsKey('ApiVersion')

# -ExtraFields is [string[]], but when this script is invoked from an external
# process (cmd.exe / a .bat file) rather than from within PowerShell, there's
# no PowerShell array-literal parsing at that boundary - only ONE token ends
# up bound here, and multi-value attempts either merge into one garbled
# string or overflow into other positional parameters. To support that
# calling style reliably, also accept a single comma-separated string and
# split it here.
$ExtraFields = @($ExtraFields | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

function Get-ValueOrPrompt {
    param([string]$Value, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return Read-Host -Prompt $Label
    }
    return $Value
}

# Same cache location/format as the other scripts in this folder.
$CredentialDir  = Join-Path $env:LOCALAPPDATA "AdoTestPlanExtractor"
$CredentialFile = Join-Path $CredentialDir "pat.dat"

function Get-SavedPat {
    if (-not (Test-Path $CredentialFile)) { return $null }
    try {
        $secure = Get-Content -Path $CredentialFile | ConvertTo-SecureString
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try {
            return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        } finally {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    } catch {
        Write-Warning "Could not read saved PAT (it may have been saved by a different Windows account/machine). Ignoring it."
        return $null
    }
}

function Save-Pat {
    param([string]$PlainPat)
    if (-not (Test-Path $CredentialDir)) {
        New-Item -ItemType Directory -Path $CredentialDir -Force | Out-Null
    }
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
    try {
        $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }

    if (-not [string]::IsNullOrWhiteSpace($plain)) {
        Save-Pat -PlainPat $plain
    }
    return $plain
}

function Get-Prop {
    param($Obj, [string[]]$Path)
    $cur = $Obj
    foreach ($p in $Path) {
        if ($null -eq $cur) { return $null }
        $cur = $cur.$p
    }
    return $cur
}

function Parse-QueryUrl {
    param([string]$Url)

    $guidMatch = [regex]::Match($Url, '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')
    if (-not $guidMatch.Success) {
        throw "Could not find a query ID (GUID) in the URL you provided: $Url"
    }

    $pathMatch = [regex]::Match($Url, '^(https?://.+?)/([^/]+)/_queries/')
    if (-not $pathMatch.Success) {
        throw "Could not parse organization/project from the URL you provided: $Url. Expected something like https://<server>/<project>/_queries/query/<guid>"
    }

    return @{
        BaseUrl = $pathMatch.Groups[1].Value.TrimEnd('/')
        Project = [Uri]::UnescapeDataString($pathMatch.Groups[2].Value)
        QueryId = $guidMatch.Value
    }
}

$QueryUrl = Get-ValueOrPrompt -Value $QueryUrl -Label "Paste the saved query's URL from your browser"
if ($ResetPat -and (Test-Path $CredentialFile)) {
    Remove-Item -Path $CredentialFile -Force
    Write-Host "Cleared saved PAT. You will be prompted for a new one."
}
$Pat = Get-SecretOrPrompt -Value $Pat -Label "Personal Access Token (press Enter to use your Windows login instead)"

$parsed  = Parse-QueryUrl -Url $QueryUrl
$BaseUrl = $parsed.BaseUrl
$Project = $parsed.Project
$QueryId = $parsed.QueryId

Write-Host "----------------------------------------"
Write-Host "Base URL  : $BaseUrl"
Write-Host "Project   : $Project"
Write-Host "Query ID  : $QueryId"
Write-Host "----------------------------------------"

$UseWindowsAuth = [string]::IsNullOrWhiteSpace($Pat)
if ($UseWindowsAuth) {
    Write-Host "No PAT provided - using your current Windows login (NTLM/Kerberos)."
    $AuthHeader = @{}
} else {
    $AuthHeader = @{
        Authorization = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    }
}

function Invoke-AdoRequest {
    param(
        [string]$UriWithoutVersion,
        [string]$Version,
        [string]$Method = "Get",
        [string]$Body = $null
    )
    $sep = "?"
    if ($UriWithoutVersion -match '\?') { $sep = "&" }
    $fullUri = $UriWithoutVersion + $sep + "api-version=$Version"
    Write-Host "  $Method [$fullUri]"
    try {
        if ($Body) {
            if ($UseWindowsAuth) {
                return Invoke-RestMethod -Uri $fullUri -UseDefaultCredentials -Method $Method -Body $Body -ContentType "application/json"
            }
            return Invoke-RestMethod -Uri $fullUri -Headers $AuthHeader -Method $Method -Body $Body -ContentType "application/json"
        }
        if ($UseWindowsAuth) {
            return Invoke-RestMethod -Uri $fullUri -UseDefaultCredentials -Method $Method
        }
        return Invoke-RestMethod -Uri $fullUri -Headers $AuthHeader -Method $Method
    } catch {
        Write-Host "  FAILED calling: $fullUri" -ForegroundColor Red
        throw
    }
}

function Test-Endpoint {
    param([string]$UriWithoutVersion, [string]$Version)
    try {
        $null = Invoke-AdoRequest -UriWithoutVersion $UriWithoutVersion -Version $Version
        return $true
    } catch {
        return $false
    }
}

function Resolve-WorkingApiVersion {
    param([string]$UriWithoutVersion, [string]$PreferredVersion)
    $candidates = @($PreferredVersion) + @("7.1","7.0","6.0","5.1","5.0","4.1","3.2","3.0","2.3","2.2","2.0","1.0") |
        Select-Object -Unique
    foreach ($v in $candidates) {
        if (Test-Endpoint -UriWithoutVersion $UriWithoutVersion -Version $v) {
            return $v
        }
    }
    return $null
}

$queryDefUri = "$BaseUrl/$Project/_apis/wit/queries/$QueryId"

if (-not $ApiVersionWasSpecified) {
    Write-Host "No -ApiVersion specified. Auto-detecting a working REST API version against the query endpoint..."
    $working = Resolve-WorkingApiVersion -UriWithoutVersion $queryDefUri -PreferredVersion $ApiVersion
    if ($working) {
        $ApiVersion = $working
        Write-Host "Using api-version=$ApiVersion"
    } else {
        Write-Warning "Could not confirm a working api-version automatically. Continuing with default $ApiVersion."
    }
}

function Get-QueryDefinition {
    return Invoke-AdoRequest -UriWithoutVersion $queryDefUri -Version $ApiVersion -Method Get
}

function Invoke-SavedQuery {
    $uri = "$BaseUrl/$Project/_apis/wit/wiql/$QueryId"
    return Invoke-AdoRequest -UriWithoutVersion $uri -Version $ApiVersion -Method Get
}

function Get-WorkItemFields {
    param([int[]]$Ids, [string[]]$Fields)
    $result = @{}
    if ($Ids.Count -eq 0) { return $result }
    $uri = "$BaseUrl/$Project/_apis/wit/workitemsbatch"
    for ($i = 0; $i -lt $Ids.Count; $i += 200) {
        $endIndex = [Math]::Min($i + 199, $Ids.Count - 1)
        $chunk = @($Ids[$i..$endIndex])
        $body = @{ ids = $chunk; fields = $Fields } | ConvertTo-Json
        $resp = Invoke-AdoRequest -UriWithoutVersion $uri -Version $ApiVersion -Method Post -Body $body
        foreach ($item in @($resp.value)) {
            $result[[int]$item.id] = $item.fields
        }
    }
    return $result
}

function Write-Results {
    param([array]$Rows, [string]$Path)
    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Warning "No rows to write."
        return
    }
    if ($Path -like "*.xlsx") {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel
            $Rows | Export-Excel -Path $Path -AutoSize
            Write-Host "Wrote $($Rows.Count) rows to $Path"
            return
        }
        Write-Warning "ImportExcel module not found. Install with: Install-Module ImportExcel -Scope CurrentUser"
        Write-Warning "Falling back to CSV."
        $Path = [IO.Path]::ChangeExtension($Path, ".csv")
    }
    $Rows | Export-Csv -Path $Path -NoTypeInformation
    Write-Host "Wrote $($Rows.Count) rows to $Path"
}

# ---- Main ----

try {
    Write-Host "Fetching query definition..."
    $queryDef = Get-QueryDefinition
    Write-Host "Query name : $($queryDef.name)"
    Write-Host "Query type : $($queryDef.queryType)"

    $columnFields = @($queryDef.columns | ForEach-Object { $_.referenceName } | Where-Object { $_ })
    if ($columnFields.Count -eq 0) {
        Write-Warning "Query definition did not return usable column names. Falling back to a default field set."
        $columnFields = @("System.WorkItemType", "System.Title", "System.State", "System.Tags", "System.AssignedTo")
    }
    if ($columnFields -notcontains "System.Id") { $columnFields = @("System.Id") + $columnFields }
    foreach ($f in $ExtraFields) {
        if ($columnFields -notcontains $f) { $columnFields += $f }
    }
    Write-Host "Fields     : $($columnFields -join ', ')"

    Write-Host "Running query..."
    $queryResult = Invoke-SavedQuery

    if ($queryResult.workItemRelations) {
        Write-Host "This is a link-based query (tree / direct links). Building relation rows..."
        $relations = @($queryResult.workItemRelations)
        $allIds = [System.Collections.Generic.HashSet[int]]::new()
        foreach ($r in $relations) {
            if ($r.source) { [void]$allIds.Add([int]$r.source.id) }
            if ($r.target) { [void]$allIds.Add([int]$r.target.id) }
        }
        Write-Host "Found $($relations.Count) relation(s) across $($allIds.Count) work item(s). Fetching fields..."
        $fieldsById = Get-WorkItemFields -Ids @($allIds) -Fields $columnFields

        $rows = @()
        foreach ($r in $relations) {
            $row = [ordered]@{}
            $row["Link Type"] = if ($r.rel) { $r.rel } else { "(root)" }

            if ($r.source) {
                $srcId = [int]$r.source.id
                $row["Source ID"] = $srcId
                $srcFields = $fieldsById[$srcId]
                foreach ($f in $columnFields) {
                    $row["Source " + $f] = Get-Prop $srcFields @($f)
                }
            } else {
                $row["Source ID"] = ""
                foreach ($f in $columnFields) { $row["Source " + $f] = "" }
            }

            $tgtId = [int]$r.target.id
            $row["Target ID"] = $tgtId
            $tgtFields = $fieldsById[$tgtId]
            foreach ($f in $columnFields) {
                $row["Target " + $f] = Get-Prop $tgtFields @($f)
            }

            $rows += [PSCustomObject]$row
        }
        Write-Results -Rows $rows -Path $OutputPath
    } else {
        $items = @($queryResult.workItems)
        Write-Host "Flat list query. Found $($items.Count) work item(s). Fetching fields..."
        $ids = @($items | ForEach-Object { [int]$_.id })
        $fieldsById = Get-WorkItemFields -Ids $ids -Fields $columnFields

        $rows = @()
        foreach ($id in $ids) {
            $f = $fieldsById[$id]
            $row = [ordered]@{}
            foreach ($col in $columnFields) {
                $row[$col] = Get-Prop $f @($col)
            }
            $rows += [PSCustomObject]$row
        }
        Write-Results -Rows $rows -Path $OutputPath
    }

    Write-Host "Done."
}
catch {
    Write-Error "Failed: $($_.Exception.Message)"
    if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
        Write-Error "Details: $($_.ErrorDetails.Message)"
    }
    exit 1
}
