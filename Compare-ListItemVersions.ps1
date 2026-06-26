<#
.SYNOPSIS
    Read-only check: compare per-version metadata (Modified/Created/Editor) between the source
    items in List A and the migrated items in List B, version by version.

.DESCRIPTION
    Confirms whether Copy-ListItemsWithHistory-RestOnly.ps1 reproduced the ACTUAL dates of each
    version (not "now"). Items are paired by creation order: source items ordered by Id are matched
    to target items ordered by Id (the migration created targets in source-Id order).

    This intentionally reads item versions the same way the copy script does: CSOM
    item.Versions[].FieldValues. Author/Editor are compared using Email when available, falling
    back to LookupValue.

.NOTES
    Requires PnP.PowerShell. Read-only.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,

    [string]$ListAName = "Migration Source A",
    [string]$ListBName = "Migration Target B"
)

$ErrorActionPreference = 'Stop'

function Get-RestCollection {
    param([Parameter(Mandatory)] [string]$RelativeUrl)
    $result = Invoke-PnPSPRestMethod -Method Get -Url $RelativeUrl
    if ($null -ne $result.value) { return @($result.value) }
    return @($result)
}

function Get-ItemIds {
    param([Parameter(Mandatory)] [string]$ListName)
    Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListName')/items?`$select=Id&`$orderby=Id&`$top=5000" |
        ForEach-Object { [int]$_.Id }
}

function Get-ItemVersions {
    param(
        [Parameter(Mandatory)] [string]$ListName,
        [Parameter(Mandatory)] [int]$ItemId
    )

    $item = Get-PnPListItem -List $ListName -Id $ItemId
    Get-PnPProperty -ClientObject $item -Property Versions | Out-Null

    $versions = @($item.Versions)
    [array]::Reverse($versions)

    $createdUtc = Format-Stamp $item.FieldValues["Created"]
    $author = Get-UserValue -UserValue $versions[0].FieldValues["Author"]

    for ($i = 0; $i -lt $versions.Count; $i++) {
        $fv = $versions[$i].FieldValues
        [pscustomobject]@{
            VersionLabel = $versions[$i].VersionLabel
            Title        = $fv["Title"]
            Notes        = $fv["Notes"]
            Created      = $createdUtc
            Modified     = Format-Stamp $fv["Modified"]
            Author       = $author
            Editor       = Get-UserValue -UserValue $fv["Editor"]
        }
    }
}

function Get-UserValue {
    param($UserValue)
    if (-not $UserValue) { return "<none>" }
    if ($UserValue.Email) { return $UserValue.Email.ToLowerInvariant() }
    if ($UserValue.LookupValue) { return $UserValue.LookupValue }
    return [string]$UserValue
}

function Format-Stamp {
    param($Value)
    if (-not $Value) { return "<none>" }
    ([datetime]$Value).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Green
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$sourceIds = Get-ItemIds -ListName $ListAName
$targetIds = Get-ItemIds -ListName $ListBName

if ($sourceIds.Count -ne $targetIds.Count) {
    Write-Host "WARNING: source item count ($($sourceIds.Count)) != target item count ($($targetIds.Count)). Pairing by position anyway." -ForegroundColor Yellow
}

$pairCount = [Math]::Min($sourceIds.Count, $targetIds.Count)
$totalVersions = 0
$matchedVersions = 0

for ($p = 0; $p -lt $pairCount; $p++) {
    $srcId = $sourceIds[$p]
    $tgtId = $targetIds[$p]

    $srcVersions = @(Get-ItemVersions -ListName $ListAName -ItemId $srcId)
    $tgtVersions = @(Get-ItemVersions -ListName $ListBName -ItemId $tgtId)

    Write-Host "`n=== Source #$srcId  ->  Target #$tgtId ===" -ForegroundColor Cyan
    if ($srcVersions.Count -ne $tgtVersions.Count) {
        Write-Host "  Version COUNT differs: source=$($srcVersions.Count) target=$($tgtVersions.Count)" -ForegroundColor Yellow
    }

    $rows = [Math]::Max($srcVersions.Count, $tgtVersions.Count)
    for ($i = 0; $i -lt $rows; $i++) {
        $s = if ($i -lt $srcVersions.Count) { $srcVersions[$i] } else { $null }
        $t = if ($i -lt $tgtVersions.Count) { $tgtVersions[$i] } else { $null }

        $sMod = if ($s) { $s.Modified } else { "<none>" }
        $tMod = if ($t) { $t.Modified } else { "<none>" }
        $sCr  = if ($s) { $s.Created } else { "<none>" }
        $tCr  = if ($t) { $t.Created } else { "<none>" }
        $sAu  = if ($s) { $s.Author } else { "<none>" }
        $tAu  = if ($t) { $t.Author } else { "<none>" }
        $sEd  = if ($s) { $s.Editor } else { "<none>" }
        $tEd  = if ($t) { $t.Editor } else { "<none>" }

        $isMatch = ($s -and $t -and
            $sMod -eq $tMod -and
            $sCr  -eq $tCr  -and
            $sAu  -eq $tAu  -and
            $sEd  -eq $tEd)
        if ($s -and $t) {
            $totalVersions++
            if ($isMatch) { $matchedVersions++ }
        }

        $flag  = if ($isMatch) { "OK " } else { "DIFF" }
        $color = if ($isMatch) { "Gray" } else { "Red" }

        $line = "  [{0}] v{1,-4} src(created={2} mod={3} author={4} editor={5})  tgt(created={6} mod={7} author={8} editor={9})" -f `
            $flag, ($s.VersionLabel ?? $t.VersionLabel), $sCr, $sMod, $sAu, $sEd, $tCr, $tMod, $tAu, $tEd
        Write-Host $line -ForegroundColor $color
    }
}

Write-Host "`nSummary: $matchedVersions / $totalVersions paired versions have matching Created, Modified, Author and Editor values." -ForegroundColor Green
if ($totalVersions -gt 0 -and $matchedVersions -eq $totalVersions) {
    Write-Host "All per-version metadata was preserved on the target." -ForegroundColor Green
}
elseif ($matchedVersions -eq 0) {
    Write-Host "No paired versions matched -> target versions were stamped differently from the source." -ForegroundColor Yellow
}

Disconnect-PnPOnline
