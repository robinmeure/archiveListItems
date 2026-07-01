<#
.SYNOPSIS
    Copy list items from List A to List B (same site collection) preserving version history and
    the ACTUAL per-version metadata (Author/Editor/Created/Modified).

.DESCRIPTION
    Reads and writes use PnP/CSOM (the driver lists item ids via REST).

    Reading : Get-PnPListItem + item.Versions[].FieldValues  (Author/Editor as FieldUserValue)
    Writing : Add-PnPListItem + Set-PnPListItem -UpdateType   (CSOM)

    Before copying items, the script builds an inventory of fields to copy: Title plus every
    non-hidden, non-read-only, non-sealed, non-base source field that also exists in the target
    list. Those fields are then copied for every replayed version.

    The source item id is not preserved when items are re-created (the target assigns new ids), so
    it is stored in a dedicated numeric column ($OriginalIdFieldName, default 'OriginalItemId').
    The column is created on the target list when missing and written on every replayed version.

    Person-field users that no longer exist cannot be written back (SharePoint rejects them with
    "The specified user could not be found"). Such users are dropped from the write and their
    original value is recorded in a text column ($UnresolvedUserFieldName, default 'UnresolvedUsers')
    so the copy continues. Any item that still fails is skipped and reported; the run does not abort.

    Why CSOM for writes: a plain REST entity POST/MERGE auto-stamps Modified=now and
    Editor=current user, so the source dates are lost. CSOM UpdateOverwriteVersion writes the
    system fields (Created/Modified/Author/Editor) WITHOUT auto-stamping and WITHOUT adding an
    extra version. Per source version we therefore:
        v1      : Add-PnPListItem, then UpdateOverwriteVersion to stamp v1's real metadata.
        v2..vN  : Update (creates the next version) then UpdateOverwriteVersion to set that
                  version's real Created/Modified/Author/Editor.
    Net result: target version count == source version count, each carrying its original dates.

    After the version history is replayed, the item's attachments are streamed from the source to
    the target. Attachments are not versioned in SharePoint, so adding them bumps the target item's
    version once; that version is restamped with the latest source metadata.

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
    [string]$ListBName = "Migration Target B",

    [string]$OriginalIdFieldName = "OriginalItemId",
    [string]$UnresolvedUserFieldName = "UnresolvedUsers"
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

function Convert-FieldValueForPnP {
    param($Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [array]) {
        return @($Value | ForEach-Object { Convert-FieldValueForPnP -Value $_ })
    }

    $typeName = $Value.GetType().FullName
    if ($typeName -eq 'Microsoft.SharePoint.Client.FieldUserValue') {
        # Person fields are lookups into the site collection's User Information List. Return the
        # numeric LookupId so the write sets the field by list id instead of resolving an
        # email/login against the directory -- that resolution throws "The specified user could not
        # be found" for accounts that have since been deleted. Source and target share the same site
        # collection, so the id is valid on the target.
        return $Value.LookupId
    }

    if ($typeName -eq 'Microsoft.SharePoint.Client.FieldLookupValue') {
        return $Value.LookupId
    }

    return $Value
}

function Get-CopyFieldInventory {
    param(
        [Parameter(Mandatory)] [string]$SourceListName,
        [Parameter(Mandatory)] [string]$TargetListName
    )

    $sourceList = Get-PnPList -Identity $SourceListName
    $targetList = Get-PnPList -Identity $TargetListName
    Get-PnPProperty -ClientObject $sourceList -Property Fields | Out-Null
    Get-PnPProperty -ClientObject $targetList -Property Fields | Out-Null

    $targetFieldNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($field in $targetList.Fields) {
        [void]$targetFieldNames.Add($field.InternalName)
    }

    $fields = [System.Collections.Generic.List[object]]::new()

    $titleField = $sourceList.Fields | Where-Object { $_.InternalName -eq 'Title' } | Select-Object -First 1
    if ($titleField -and $targetFieldNames.Contains('Title')) {
        $fields.Add([pscustomobject]@{
            InternalName = 'Title'
            Title        = $titleField.Title
            TypeAsString = $titleField.TypeAsString
            Category     = 'Built-in'
        })
    }

    foreach ($field in $sourceList.Fields) {
        if ($field.InternalName -eq 'Title') { continue }
        if ($field.Hidden -or $field.ReadOnlyField -or $field.Sealed -or $field.FromBaseType) { continue }

        if (-not $targetFieldNames.Contains($field.InternalName)) {
            Write-Warning "Skipping field '$($field.InternalName)' because it does not exist on target list '$TargetListName'."
            continue
        }

        $fields.Add([pscustomobject]@{
            InternalName = $field.InternalName
            Title        = $field.Title
            TypeAsString = $field.TypeAsString
            Category     = 'Custom'
        })
    }

    return @($fields)
}

function Confirm-ListColumn {
    param(
        [Parameter(Mandatory)] [string]$ListName,
        [Parameter(Mandatory)] [string]$FieldInternalName,
        [Parameter(Mandatory)] [ValidateSet('Number', 'Note')] [string]$FieldType
    )

    # Ensure a helper column exists on the target list before the copy loop runs (idempotent):
    #   Number -> holds the original source item id (ids are reassigned on the target).
    #   Note   -> holds the original value of any person field whose user could not be resolved.
    $existing = Get-PnPField -List $ListName -Identity $FieldInternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "Column '$FieldInternalName' already exists on '$ListName'." -ForegroundColor Gray
        return
    }

    Write-Host "Creating '$FieldType' column '$FieldInternalName' on '$ListName'..." -ForegroundColor Yellow
    Add-PnPField -List $ListName -DisplayName $FieldInternalName -InternalName $FieldInternalName -Type $FieldType -AddToDefaultView | Out-Null
}

function Test-SiteUser {
    param([Parameter(Mandatory)] [int]$LookupId)

    # Person-field values are User Information List lookups. A row that still exists can be written
    # back by its numeric id; a deleted account whose row is gone cannot, and the write then fails
    # with "The specified user could not be found". Cache the site's current user ids once and test
    # membership so unresolvable users are detected up front instead of aborting the copy.
    if ($null -eq $script:ValidUserIds) {
        $script:ValidUserIds = [System.Collections.Generic.HashSet[int]]::new()
        foreach ($siteUser in Get-PnPUser) { [void]$script:ValidUserIds.Add([int]$siteUser.Id) }
    }

    if (-not $LookupId) { return $false }
    return $script:ValidUserIds.Contains([int]$LookupId)
}

function New-FieldValueMap {
    param(
        [Parameter(Mandatory)] $FieldValues,
        [Parameter(Mandatory)] [object[]]$CopyFields
    )

    # Returns the values to write plus a note of any person-field users that could not be resolved
    # (so the caller can record them and keep going instead of failing the whole item).
    $values = @{}
    $unresolved = [System.Collections.Generic.List[string]]::new()

    foreach ($field in $CopyFields) {
        if (-not $FieldValues.ContainsKey($field.InternalName)) { continue }
        $raw = $FieldValues[$field.InternalName]
        if ($null -eq $raw) { continue }

        if ($field.TypeAsString -eq 'User' -or $field.TypeAsString -eq 'UserMulti') {
            $resolvedIds = [System.Collections.Generic.List[int]]::new()
            foreach ($user in @($raw)) {
                if ($null -eq $user) { continue }
                if (Test-SiteUser -LookupId $user.LookupId) {
                    $resolvedIds.Add([int]$user.LookupId)
                }
                else {
                    $label = if ($user.Email) { $user.Email } elseif ($user.LookupValue) { $user.LookupValue } else { "id $($user.LookupId)" }
                    $unresolved.Add("$($field.InternalName): $label")
                }
            }
            if ($resolvedIds.Count -gt 0) {
                if ($field.TypeAsString -eq 'UserMulti') { $values[$field.InternalName] = $resolvedIds.ToArray() }
                else { $values[$field.InternalName] = $resolvedIds[0] }
            }
        }
        else {
            $values[$field.InternalName] = Convert-FieldValueForPnP -Value $raw
        }
    }

    return [pscustomobject]@{
        Values     = $values
        Unresolved = if ($unresolved.Count -gt 0) { $unresolved -join '; ' } else { $null }
    }
}

function Copy-ItemAttachments {
    param(
        [Parameter(Mandatory)] [int]$SourceItemId,
        [Parameter(Mandatory)] [int]$TargetItemId
    )

    # Attachments are not versioned in SharePoint -- they hang off the current item -- so they are
    # copied once after the version history has been replayed. Each attachment is streamed straight
    # from the source item's AttachmentFiles to the target item without touching the local disk.
    $attachments = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListAName')/items($SourceItemId)/AttachmentFiles?`$select=FileName,ServerRelativeUrl"
    if ($attachments.Count -eq 0) { return 0 }

    $copied = 0
    foreach ($attachment in $attachments) {
        $stream = Get-PnPFile -Url $attachment.ServerRelativeUrl -AsMemoryStream
        Add-PnPListItemAttachment -List $ListBName -Identity $TargetItemId -FileName $attachment.FileName -Stream $stream | Out-Null
        $copied++
    }

    return $copied
}

function Set-VersionStamp {
    <#
      Overwrite the target item's current version with the source version's real system fields
      (Created/Modified/Author/Editor + business values) without creating an extra version.

      Author/Editor are lookup fields into the site collection's User Information List. The normal
      stamp passes a login/email that SharePoint resolves against the directory, but when the
      original account has been deleted that resolution fails with "The specified user could not be
      found" and the whole stamp is rejected. In that case retry with Author/Editor set to their
      User Information List lookup ids: passing the numeric id sets the field by lookup without a
      directory call, so the original user is preserved as long as their row in that list survives.
    #>
    param(
        [Parameter(Mandatory)] [int]$TargetItemId,
        [Parameter(Mandatory)] [hashtable]$StampValues,
        $AuthorUser,
        $EditorUser
    )

    try {
        Set-PnPListItem -List $ListBName -Identity $TargetItemId -Values $StampValues -UpdateType UpdateOverwriteVersion | Out-Null
    }
    catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException) { $message += " " + $_.Exception.InnerException.Message }
        if ($message -notmatch 'could not be found|cannot be found') { throw }

        # Directory lookup failed (deleted account); fall back to the User Information List row ids.
        $retryValues = $StampValues.Clone()
        if ($null -ne $AuthorUser) { $retryValues["Author"] = $AuthorUser.LookupId }
        if ($null -ne $EditorUser) { $retryValues["Editor"] = $EditorUser.LookupId }

        Write-Warning "Item #$TargetItemId : Author/Editor could not be resolved by login/email (deleted account); restamping Author=$($AuthorUser.LookupId), Editor=$($EditorUser.LookupId) via the User Information List."

        Set-PnPListItem -List $ListBName -Identity $TargetItemId -Values $retryValues -UpdateType UpdateOverwriteVersion | Out-Null
    }
}

# ---------------------------------------------------------------------------
# Migration of a single item (replays all versions oldest -> newest)
# ---------------------------------------------------------------------------

function Copy-ItemWithVersionDates {
    param(
        [Parameter(Mandatory)] [int]$SourceItemId,
        [Parameter(Mandatory)] [object[]]$CopyFields
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
    $unresolvedWarned = $false

    for ($i = 0; $i -lt $ordered.Count; $i++) {
        $fv          = $ordered[$i].FieldValues
        $modifiedUtc = ConvertTo-Utc -Value $fv["Modified"]   # this version's Modified timestamp
        $editorUser  = $fv["Editor"]
        $editorValue = if ($editorUser.Email) { $editorUser.Email } else { $editorUser.LookupValue }
        $mapped = New-FieldValueMap -FieldValues $fv -CopyFields $CopyFields
        $businessValues = $mapped.Values
        $businessValues[$OriginalIdFieldName] = $SourceItemId   # preserve the source item id in the extra numeric column
        if ($mapped.Unresolved) {
            # A person column referenced a user that no longer exists; keep the copy going and note
            # the original value(s) in the fallback text column instead of failing this item.
            $businessValues[$UnresolvedUserFieldName] = $mapped.Unresolved
            if (-not $unresolvedWarned) {
                Write-Warning "Source #$SourceItemId : unresolved user(s) [$($mapped.Unresolved)] stored in column '$UnresolvedUserFieldName'."
                $unresolvedWarned = $true
            }
        }

        # Full system-field stamp for this version (Author/Editor by resolvable user value, dates as UTC).
        $stampValues = $businessValues.Clone()
        $stampValues["Created"] = $createdUtc
        $stampValues["Modified"] = $modifiedUtc
        $stampValues["Author"] = $authorValue
        $stampValues["Editor"] = $editorValue

        if ($i -eq 0) {
            # Create v1, then overwrite the current version's stamps (no extra version created).
            $newItem   = Add-PnPListItem -List $ListBName -Values $businessValues
            $newItemId = $newItem.Id
            Set-VersionStamp -TargetItemId $newItemId -StampValues $stampValues -AuthorUser $authorUser -EditorUser $editorUser
        }
        else {
            # Update creates the next version; the stamp then fixes that version's system fields.
            Set-PnPListItem -List $ListBName -Identity $newItemId -Values $businessValues -UpdateType Update | Out-Null
            Set-VersionStamp -TargetItemId $newItemId -StampValues $stampValues -AuthorUser $authorUser -EditorUser $editorUser
        }
    }

    # Copy attachments onto the finished item. Adding an attachment bumps the item version, so we
    # restamp that version with the latest source metadata ($stampValues from the final loop pass)
    # to keep its Author/Editor/Modified correct instead of the current user/now.
    $attachmentCount = Copy-ItemAttachments -SourceItemId $SourceItemId -TargetItemId $newItemId
    if ($attachmentCount -gt 0) {
        Set-VersionStamp -TargetItemId $newItemId -StampValues $stampValues -AuthorUser $authorUser -EditorUser $editorUser
    }

    return $newItemId
}

# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Green
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

Confirm-ListColumn -ListName $ListBName -FieldInternalName $OriginalIdFieldName -FieldType Number
Confirm-ListColumn -ListName $ListBName -FieldInternalName $UnresolvedUserFieldName -FieldType Note

$copyFields = Get-CopyFieldInventory -SourceListName $ListAName -TargetListName $ListBName
if ($copyFields.Count -eq 0) {
    throw "No copyable fields were found. The target list must at least contain the Title field or matching custom fields."
}

Write-Host "`nField inventory to copy:" -ForegroundColor Cyan
foreach ($field in $copyFields) {
    Write-Host "  [$($field.Category)] $($field.InternalName) ('$($field.Title)', $($field.TypeAsString))" -ForegroundColor Gray
}

$sourceIds = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListAName')/items?`$select=Id&`$orderby=Id&`$top=5000" |
    ForEach-Object { [int]$_.Id }

Write-Host "Migrating $($sourceIds.Count) item(s) '$ListAName' -> '$ListBName' (with per-version dates)..." -ForegroundColor Green

$failedIds = [System.Collections.Generic.List[int]]::new()
foreach ($id in $sourceIds) {
    try {
        $newId = Copy-ItemWithVersionDates -SourceItemId $id -CopyFields $copyFields
        Write-Host "  Copied source #$id -> target #$newId" -ForegroundColor DarkGray
    }
    catch {
        $failedIds.Add($id)
        Write-Warning "Skipped source #$id : $($_.Exception.Message)"
    }
}

if ($failedIds.Count -gt 0) {
    Write-Warning "Completed with $($failedIds.Count) item(s) skipped due to errors: $($failedIds -join ', ')."
}

# ---------------------------------------------------------------------------
# Verify history landed
# ---------------------------------------------------------------------------

Write-Host "`nVerifying target history..." -ForegroundColor Cyan
$targetIds = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListBName')/items?`$select=Id&`$orderby=Id&`$top=5000" |
    ForEach-Object { [int]$_.Id }

foreach ($id in $targetIds) {
    $vers = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListBName')/items($id)/versions?`$select=VersionLabel,Modified,Title"
    $attachments = Get-RestCollection -RelativeUrl "/_api/web/lists/getbytitle('$ListBName')/items($id)/AttachmentFiles?`$select=FileName"

    # Read the current (latest) version metadata via CSOM -- Author/Editor are FieldUserValue, which
    # the version replay (and the post-attachment restamp) is supposed to have set to the source's
    # real values. This proves the attachment-add version did NOT keep the migration account / now.
    $targetItem = Get-PnPListItem -List $ListBName -Id $id
    $editorUser = $targetItem.FieldValues["Editor"]
    $editorName = if ($editorUser.Email) { $editorUser.Email } else { $editorUser.LookupValue }
    $modified   = $targetItem.FieldValues["Modified"]
    $originalId = $targetItem.FieldValues[$OriginalIdFieldName]

    Write-Host "  Target #$id : $($vers.Count) versions (latest '$($vers[0].Title)'), $($attachments.Count) attachment(s)" -ForegroundColor Gray
    Write-Host "             current metadata -> Editor: $editorName, Modified: $modified, OriginalItemId: $originalId" -ForegroundColor DarkGray
}

Write-Host "`nDone." -ForegroundColor Green
Disconnect-PnPOnline
