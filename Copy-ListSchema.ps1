<#
.SYNOPSIS
    Clone a SharePoint Online list schema into a new list, including custom/list-specific columns.

.DESCRIPTION
    Creates a target list using the source list's base template, then recreates cloneable fields
    from the source list by using each field's SchemaXml. Built-in fields already present on the
    new list (ID, Created, Modified, Author, Editor, Title, etc.) are not recreated. The Title
    field's display/required/hidden settings are copied separately.

    This script is intended to prepare List B so item/version migration can use matching internal
    field names. It does not copy list items.

.NOTES
    Requires PnP.PowerShell. Connects interactively with a registered Entra ID app (ClientId).
    Some SharePoint field types can have dependencies or restrictions (for example taxonomy,
    publishing, sealed/system fields). Those are skipped or reported as warnings.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$SiteUrl,
    [Parameter(Mandatory = $true)] [string]$ClientId,

    [Parameter(Mandatory = $true)] [string]$SourceListName,
    [Parameter(Mandatory = $true)] [string]$TargetListName,

    [switch]$Overwrite,
    [switch]$IncludeHidden
)

$ErrorActionPreference = 'Stop'

function Get-ListSafeTemplate {
    param([Parameter(Mandatory)] $List)

    try {
        return [Microsoft.SharePoint.Client.ListTemplateType]$List.BaseTemplate
    }
    catch {
        Write-Warning "Could not cast base template '$($List.BaseTemplate)' to ListTemplateType. Falling back to GenericList."
        return [Microsoft.SharePoint.Client.ListTemplateType]::GenericList
    }
}

function ConvertTo-CloneFieldXml {
    param(
        [Parameter(Mandatory)] $Field,
        [Parameter(Mandatory)] [guid]$SourceListId,
        [Parameter(Mandatory)] [guid]$TargetListId
    )

    [xml]$doc = $Field.SchemaXml
    $fieldNode = $doc.Field

    # Let SharePoint generate a new field GUID and storage details, while preserving Name/StaticName.
    foreach ($attribute in @('ID', 'SourceID', 'Version', 'ColName', 'RowOrdinal', 'WebId')) {
        if ($fieldNode.HasAttribute($attribute)) {
            $fieldNode.RemoveAttribute($attribute)
        }
    }

    # If the source field is a self-lookup, repoint it to the newly-created target list.
    if ($fieldNode.HasAttribute('List')) {
        $lookupList = $fieldNode.GetAttribute('List').Trim('{}')
        if ($lookupList -ieq $SourceListId.ToString('D')) {
            $fieldNode.SetAttribute('List', "{$($TargetListId.ToString('D'))}")
        }
    }

    return $doc.OuterXml
}

function Copy-TitleFieldSettings {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] $SourceList,
        [Parameter(Mandatory)] $TargetList
    )

    $sourceTitle = $SourceList.Fields.GetByInternalNameOrTitle('Title')
    $targetTitle = $TargetList.Fields.GetByInternalNameOrTitle('Title')
    $Context.Load($sourceTitle)
    $Context.Load($targetTitle)
    $Context.ExecuteQuery()

    $targetTitle.Title = $sourceTitle.Title
    $targetTitle.Required = $sourceTitle.Required
    $targetTitle.Hidden = $sourceTitle.Hidden
    $targetTitle.Update()
    $Context.ExecuteQuery()
}

function Copy-ListSettings {
    param(
        [Parameter(Mandatory)] $Context,
        [Parameter(Mandatory)] $SourceList,
        [Parameter(Mandatory)] $TargetList
    )

    $TargetList.ContentTypesEnabled = $SourceList.ContentTypesEnabled
    $TargetList.EnableVersioning = $SourceList.EnableVersioning
    $TargetList.EnableMinorVersions = $SourceList.EnableMinorVersions
    $TargetList.MajorVersionLimit = $SourceList.MajorVersionLimit
    $TargetList.MajorWithMinorVersionsLimit = $SourceList.MajorWithMinorVersionsLimit
    $TargetList.ForceCheckout = $SourceList.ForceCheckout
    $TargetList.Update()
    $Context.ExecuteQuery()
}

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Green
Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId

$ctx = Get-PnPContext
$sourceList = Get-PnPList -Identity $SourceListName
Get-PnPProperty -ClientObject $sourceList -Property Fields, ContentTypesEnabled, EnableVersioning, EnableMinorVersions, MajorVersionLimit, MajorWithMinorVersionsLimit, ForceCheckout, BaseTemplate, Id | Out-Null

$existingTarget = Get-PnPList -Identity $TargetListName -ErrorAction SilentlyContinue
if ($existingTarget) {
    if (-not $Overwrite) {
        throw "Target list '$TargetListName' already exists. Re-run with -Overwrite to delete and recreate it."
    }

    Write-Host "Removing existing target list '$TargetListName'..." -ForegroundColor DarkYellow
    Remove-PnPList -Identity $TargetListName -Force
}

$template = Get-ListSafeTemplate -List $sourceList
Write-Host "Creating target list '$TargetListName' using template '$template'..." -ForegroundColor Cyan
New-PnPList -Title $TargetListName -Template $template -OnQuickLaunch | Out-Null

$targetList = Get-PnPList -Identity $TargetListName
Get-PnPProperty -ClientObject $targetList -Property Fields, Id | Out-Null

Write-Host "Copying list settings..." -ForegroundColor Cyan
Copy-ListSettings -Context $ctx -SourceList $sourceList -TargetList $targetList

Write-Host "Copying Title field settings..." -ForegroundColor Cyan
try {
    Copy-TitleFieldSettings -Context $ctx -SourceList $sourceList -TargetList $targetList
}
catch {
    Write-Warning "Could not copy Title field settings: $($_.Exception.Message)"
}

# Refresh target fields after settings changes.
$targetList = Get-PnPList -Identity $TargetListName
Get-PnPProperty -ClientObject $targetList -Property Fields, Id | Out-Null
$targetFieldNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($field in $targetList.Fields) {
    [void]$targetFieldNames.Add($field.InternalName)
}

$copied = 0
$skipped = 0
$failed = 0

Write-Host "Copying fields..." -ForegroundColor Cyan
foreach ($field in $sourceList.Fields) {
    if ($targetFieldNames.Contains($field.InternalName)) {
        $skipped++
        continue
    }

    if (-not $IncludeHidden -and $field.Hidden) {
        $skipped++
        continue
    }

    if ($field.ReadOnlyField -or $field.Sealed -or $field.FromBaseType) {
        $skipped++
        continue
    }

    try {
        $fieldXml = ConvertTo-CloneFieldXml -Field $field -SourceListId $sourceList.Id -TargetListId $targetList.Id
        Add-PnPFieldFromXml -List $TargetListName -FieldXml $fieldXml | Out-Null
        [void]$targetFieldNames.Add($field.InternalName)
        $copied++
        Write-Host "  Copied field '$($field.InternalName)' ($($field.TypeAsString))" -ForegroundColor Gray
    }
    catch {
        $failed++
        Write-Warning "Failed to copy field '$($field.InternalName)' ($($field.TypeAsString)): $($_.Exception.Message)"
    }
}

Write-Host "`nDone cloning list schema." -ForegroundColor Green
Write-Host "  Source : $SourceListName" -ForegroundColor Gray
Write-Host "  Target : $TargetListName" -ForegroundColor Gray
Write-Host "  Copied : $copied field(s)" -ForegroundColor Gray
Write-Host "  Skipped: $skipped field(s)" -ForegroundColor Gray
Write-Host "  Failed : $failed field(s)" -ForegroundColor Gray

Disconnect-PnPOnline
