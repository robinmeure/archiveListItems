<#
.SYNOPSIS
    Copy list items from List A to List B (same site collection) preserving version history and
    the ACTUAL per-version metadata (Author/Editor/Created/Modified).

.DESCRIPTION
    Reads and writes use PnP/CSOM (the driver lists item ids via REST).

    Reading : Get-PnPListItem + item.Versions[].FieldValues  (Author/Editor as FieldUserValue)
    Writing : Add-PnPListItem + Set-PnPListItem -UpdateType   (CSOM)

    Why CSOM for writes: a plain REST entity POST/MERGE auto-stamps Modified=now and
    Editor=current user, so the source dates are lost. CSOM UpdateOverwriteVersion writes the
    system fields (Created/Modified/Author/Editor) WITHOUT auto-stamping and WITHOUT adding an
    extra version. Per source version we therefore:
        v1      : Add-PnPListItem, then UpdateOverwriteVersion to stamp v1's real metadata.
        v2..vN  : Update (creates the next version) then UpdateOverwriteVersion to set that
                  version's real Created/Modified/Author/Editor.
    Net result: target version count == source version count, each carrying its original dates.

.NOTES
    Requires PnP.PowerShell. Connects interactively with a registered Entra ID app (ClientId).
    Run Copy-ListItemsWithHistory.ps1 first to provision List A items with version history.

    Dates are passed to CSOM as UTC (DateTimeKind.Utc) so the stored instant matches the source.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,

    [string]$ListAName = "Migration Source A",
    [string]$ListBName = "Migration Target B"
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# REST helpers (SharePoint REST only)
# ---------------------------------------------------------------------------

function Get-RestCollection {
    param([Parameter(Mandatory)] [string]$RelativeUrl)
    $result = Invoke-PnPSPRestMethod -Method Get -Url $RelativeUrl
    if ($null -ne $result.value) { return @($result.value) }
    return @($result)
}

function Get-RestObject {
    param([Parameter(Mandatory)] [string]$RelativeUrl)
    Invoke-PnPSPRestMethod -Method Get -Url $RelativeUrl
}

function ConvertTo-Utc {
    param([Parameter(Mandatory)] $Value)
    ([datetime]$Value).ToUniversalTime()
}

# ---------------------------------------------------------------------------
# Migration of a single item (replays all versions oldest -> newest)
# ---------------------------------------------------------------------------

function Copy-ItemWithVersionDates {
    param(
        [Parameter(Mandatory)] [int]$SourceItemId
    )

    # Read the source item's full version history via CSOM. Loading Versions populates each
    # version's FieldValues dictionary; Author/Editor are FieldUserValue objects whose .LookupId
    # is a typed property -- reliable, unlike REST JSON where nested-object dot access returns blank.
    $item = Get-PnPListItem -List $ListAName -Id $SourceItemId
    Get-PnPProperty -ClientObject $item -Property Versions | Out-Null

    $ordered = @($item.Versions)
    [array]::Reverse($ordered)   # CSOM returns newest-first -> reverse to oldest-first
    if ($ordered.Count -eq 0) { return $null }

    # Original creation date is the item-level Created (constant); author = editor of oldest version.
    # NOTE: a version's FieldValues has no "Created" key (it's "Created_x0020_Date"); use the item.
    $createdUtc = ConvertTo-Utc -Value $item.FieldValues["Created"]
    $authorUser = $ordered[0].FieldValues["Author"]
    $authorValue = if ($authorUser.Email) { $authorUser.Email } else { $authorUser.LookupValue }

    $newItemId = $null

    for ($i = 0; $i -lt $ordered.Count; $i++) {
        $fv          = $ordered[$i].FieldValues
        $modifiedUtc = ConvertTo-Utc -Value $fv["Modified"]   # this version's Modified timestamp
        $editorUser  = $fv["Editor"]
        $editorValue = if ($editorUser.Email) { $editorUser.Email } else { $editorUser.LookupValue }

        # Full system-field stamp for this version (Author/Editor by resolvable user value, dates as UTC).
        $stampValues = @{
            Title    = $fv["Title"]
            Notes    = $fv["Notes"]
            Created  = $createdUtc
            Modified = $modifiedUtc
            Author   = $authorValue
            Editor   = $editorValue
        }

        if ($i -eq 0) {
            # Create v1, then overwrite the current version's stamps (no extra version created).
            $newItem   = Add-PnPListItem -List $ListBName -Values @{ Title = $fv["Title"]; Notes = $fv["Notes"] }
            $newItemId = $newItem.Id
            Set-PnPListItem -List $ListBName -Identity $newItemId -Values $stampValues -UpdateType UpdateOverwriteVersion | Out-Null
        }
        else {
            # Update creates the next version; UpdateOverwriteVersion fixes that version's stamps.
            Set-PnPListItem -List $ListBName -Identity $newItemId -Values @{ Title = $fv["Title"]; Notes = $fv["Notes"] } -UpdateType Update | Out-Null
            Set-PnPListItem -List $ListBName -Identity $newItemId -Values $stampValues -UpdateType UpdateOverwriteVersion | Out-Null
        }
    }

    return $newItemId
}

# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Green
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$sourceIds = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListAName')/items?`$select=Id&`$orderby=Id&`$top=5000" |
    ForEach-Object { [int]$_.Id }

Write-Host "Migrating $($sourceIds.Count) item(s) '$ListAName' -> '$ListBName' (with per-version dates)..." -ForegroundColor Green

foreach ($id in $sourceIds) {
    $newId = Copy-ItemWithVersionDates -SourceItemId $id
    Write-Host "  Copied source #$id -> target #$newId" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------------------
# Verify history landed
# ---------------------------------------------------------------------------

Write-Host "`nVerifying target history..." -ForegroundColor Cyan
$targetIds = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListBName')/items?`$select=Id&`$orderby=Id&`$top=5000" |
    ForEach-Object { [int]$_.Id }

foreach ($id in $targetIds) {
    $vers = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListBName')/items($id)/versions?`$select=VersionLabel,Modified,Title"
    Write-Host "  Target #$id : $($vers.Count) versions (latest '$($vers[0].Title)')" -ForegroundColor Gray
}

Write-Host "`nDone." -ForegroundColor Green
Disconnect-PnPOnline
