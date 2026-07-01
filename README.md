# Archive SharePoint List Items With Version History

This repository contains PowerShell scripts that help copy SharePoint Online list items from one list to another while preserving version history and metadata.

The main goal is to support this scenario:

- You have a source list, for example `Migration Source A`.
- You want a target/archive list, for example `Migration Target B`.
- You want to copy the items, including their version history.
- You want each copied version to keep its original `Created`, `Modified`, `Author`, and `Editor` metadata.
- You want custom columns, such as `Notes`, to be copied as well.
- You want each item's attachments to be copied to the target as well.
- You want the original source item id preserved in a column, because the target list assigns new item ids.

## What The Scripts Do

| Script | Purpose |
| --- | --- |
| `Copy-ListSchema.ps1` | Creates a new list based on an existing list and copies its custom/list-specific columns. Run this first when your target list does not already have matching columns. |
| `Copy-ListItemsWithHistory-RestOnly.ps1` | Copies list items from source to target, recreates each version, copies all included field values, copies each item's attachments, preserves the original source item id in a numeric column, and stamps the original metadata on each version. |
| `Compare-ListItemVersions.ps1` | Verifies that source and target versions match for `Created`, `Modified`, `Author`, and `Editor`. |

> Note: The item copy script name still includes `RestOnly` from an earlier experiment. The final working version uses PnP/CSOM for version and metadata writes because that is what reliably preserves per-version system fields in SharePoint Online.

## Requirements

You need:

1. A Windows machine with PowerShell 7 or Windows PowerShell.
2. Access to the SharePoint Online site.
3. Permission to create lists and edit list items on the site.
4. The `PnP.PowerShell` module installed.
5. An Entra ID app registration client id that can be used with PnP PowerShell interactive sign-in.

## Install PnP PowerShell

Open PowerShell and run:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

If PowerShell asks whether you trust the repository, answer `Y` and press Enter.

To confirm it installed:

```powershell
Get-Module PnP.PowerShell -ListAvailable
```

## Get Your Site URL

Your site URL looks like this:

```text
https://contoso.sharepoint.com/sites/MySite
```

In the examples below, replace the URL with your real site URL.

## Get Your Client ID

The scripts connect using PnP PowerShell interactive authentication:

```powershell
Connect-PnPOnline -Url <site-url> -Interactive -ClientId <client-id>
```

If you already have a client id that works with PnP PowerShell, use that.

If you do not have one, ask your Microsoft 365 administrator to provide or register an Entra ID app for PnP PowerShell interactive delegated access to SharePoint Online. The account you sign in with must have permissions on the SharePoint site.

Common delegated permissions used for this kind of operation are:

- `AllSites.FullControl` for SharePoint delegated access, or
- an equivalent admin-approved permission model that allows list creation and item updates.

Your tenant may have stricter policies, so use whatever your administrator approves.

## Recommended Workflow

Use this order:

1. Set variables for the site and client id.
2. Clone the source list schema to create the target list.
3. Copy the items and their version history.
4. Compare source and target versions.

## Step 1: Open PowerShell In This Folder

Open PowerShell and go to this repository folder:

```powershell
cd D:\repos\lipton
```

If your repository is in a different location, use that folder instead.

## Step 2: Set Your Variables

Paste this into PowerShell and update the values:

```powershell
$siteurl = "https://contoso.sharepoint.com/sites/MySite"
$clientId = "00000000-0000-0000-0000-000000000000"
$sourceList = "Migration Source A"
$targetList = "Migration Target B"
```

Replace:

- `https://contoso.sharepoint.com/sites/MySite` with your SharePoint site URL.
- `00000000-0000-0000-0000-000000000000` with your app registration client id.
- `Migration Source A` with your source list name.
- `Migration Target B` with your target/archive list name.

## Step 3: Create Or Clone The Target List

Create or synchronize the target list schema:

```powershell
.\Copy-ListSchema.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -SourceListName $sourceList `
  -TargetListName $targetList
```

If the target list does not exist, this creates it and copies cloneable custom/list-specific columns.

If the target list already exists, this reuses it, checks whether matching columns already exist, adds missing cloneable columns, and warns about existing columns whose basic settings differ from the source.

If the target list already exists and you want to delete and recreate it, use `-Overwrite`:

```powershell
.\Copy-ListSchema.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -SourceListName $sourceList `
  -TargetListName $targetList `
  -Overwrite
```

Warning: `-Overwrite` removes the existing target list and its data.

## Step 4: Copy Items And Version History

Run:

```powershell
.\Copy-ListItemsWithHistory-RestOnly.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -ListAName $sourceList `
  -ListBName $targetList
```

During the run, the script prints a field inventory, for example:

```text
Field inventory to copy:
  [Built-in] Title ('Title', Text)
  [Custom] Notes ('Notes', Text)
```

This inventory tells you which fields will be copied for every version.

The script copies:

- `Title`
- custom/non-default fields that exist on both lists
- each item's version history
- each version's original `Created`
- each version's original `Modified`
- each version's original `Author`
- each version's original `Editor`
- each item's attachments
- the original source item id, stored in a numeric column (`OriginalItemId` by default)

After the version history is replayed, the script streams each source attachment to the matching target item. Because SharePoint does not version attachments, adding them bumps the target item's version once; the script restamps that version with the latest source metadata so its `Author`, `Editor`, and `Modified` stay correct instead of showing the migration account and current time.

The verify output at the end of the run reports the version count, attachment count, and the current `Editor`, `Modified`, and `OriginalItemId` for each target item, for example:

```text
  Target #1 : 3 versions (latest 'Item one'), 2 attachment(s)
             current metadata -> Editor: user@contoso.com, Modified: 1/2/2026 10:15:00 AM, OriginalItemId: 1
```

## Step 5: Verify The Copy

Run:

```powershell
.\Compare-ListItemVersions.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -ListAName $sourceList `
  -ListBName $targetList
```

The compare script checks source and target versions side by side.

A successful comparison looks like this:

```text
[OK ] v1.0  src(created=... mod=... author=... editor=...)  tgt(created=... mod=... author=... editor=...)
[OK ] v2.0  src(created=... mod=... author=... editor=...)  tgt(created=... mod=... author=... editor=...)
```

At the end, you want to see:

```text
All per-version metadata was preserved on the target.
```

## Important Notes

### Run On A Clean Target List

For easiest verification, run the copy into an empty target list.

If you run the copy multiple times into the same target list, the target will contain duplicate copied items. The compare script pairs source and target items by order, so duplicate previous runs can make the output confusing.

For repeated tests, either:

- delete items from the target list first, or
- recreate the target list with `Copy-ListSchema.ps1 -Overwrite`.

### Target Columns Must Exist

The item copy script only copies custom fields that exist on both source and target.

If a source field is missing on the target list, the script prints a warning and skips that field.

Use `Copy-ListSchema.ps1` first if you want the target list to have matching custom columns.

### Attachments Add One Extra Version

SharePoint does not version attachments, so they are copied once after the version history is replayed. Adding an attachment to an existing item always creates a new version, so a target item that has attachments will end up with one more version than the source. The script restamps that extra version with the latest source metadata, so its `Author`, `Editor`, and `Modified` are correct (they reflect the source's latest edit, not the migration account).

### Original Item Id Is Stored In A Column

When items are recreated in the target list, SharePoint assigns them new item ids, so the source `ID` cannot be reused. To keep the link back to the source, the copy script writes the source item id into a dedicated numeric column on the target list.

- The column is named `OriginalItemId` by default. Change it with `-OriginalIdFieldName` if you already use that name for something else.
- The script creates the column automatically if it does not exist, so you do not have to add it to the target list yourself.
- The value is written on every replayed version, so it is present throughout the item's history.

### Some Field Types May Need Extra Handling

The scripts handle common field values, including:

- text
- note/multiline text
- number
- date/time
- choice
- yes/no
- person fields
- lookup fields

Some complex SharePoint field types may require extra work, such as:

- managed metadata/taxonomy fields
- publishing fields
- location fields
- custom field types
- fields depending on unavailable lookup lists

If one of those fields fails, check the warning or error message and decide whether to skip that field or add field-specific conversion logic.

### This Does Not Use A Native SharePoint Move

SharePoint Online does not provide a supported native move operation for generic list items that preserves the entire item and version history.

Document libraries have real file/folder move APIs, but generic list items do not. That is why this solution recreates each version in the target list and then stamps the original metadata onto that version.

## Troubleshooting

### PowerShell Says Scripts Are Blocked

If PowerShell blocks script execution, run this once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Then run the script again.

### PnP.PowerShell Is Not Found

Install it:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Then restart PowerShell.

### Login Window Does Not Appear

Make sure you are running the script in an interactive PowerShell session, not a non-interactive job.

Also confirm your client id is valid for interactive PnP PowerShell login.

### Target List Already Exists

If the target list already exists, `Copy-ListSchema.ps1` now reuses it by default. It checks existing fields, adds missing cloneable fields, and warns if an existing field differs from the source.

Use `-Overwrite` only when you intentionally want to delete and recreate the target list:

```powershell
-Overwrite
```

Warning: `-Overwrite` deletes and recreates the target list.

### A Custom Column Is Skipped

The most common reason is that the field exists on the source list but not on the target list.

Run `Copy-ListSchema.ps1` first, or manually create the missing column on the target list with the same internal name.

### A User Cannot Be Found

The copy script writes `Author` and `Editor` using the user's email or login where available. When an account has been deleted from the directory, SharePoint can no longer resolve it and returns `The specified user could not be found`.

The script handles this automatically. `Author` and `Editor` are lookup fields into the site's User Information List, so when the login/email cannot be resolved the script retries using the user's existing User Information List id. This preserves the original author/editor even for deleted accounts, as long as their entry still exists in the User Information List (which is the normal case within the same site collection).

If you see a warning like `Author/Editor could not be resolved by login/email (deleted account); restamping ... via the User Information List`, that is the fallback working as intended, not an error.

## Example End-To-End Run

This is a complete example you can paste after changing the first two variables:

```powershell
cd D:\repos\lipton

$siteurl = "https://contoso.sharepoint.com/sites/MySite"
$clientId = "00000000-0000-0000-0000-000000000000"
$sourceList = "Migration Source A"
$targetList = "Migration Target B"

.\Copy-ListSchema.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -SourceListName $sourceList `
  -TargetListName $targetList `
  -Overwrite

.\Copy-ListItemsWithHistory-RestOnly.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -ListAName $sourceList `
  -ListBName $targetList

.\Compare-ListItemVersions.ps1 `
  -SiteUrl $siteurl `
  -ClientId $clientId `
  -ListAName $sourceList `
  -ListBName $targetList
```

## File Summary

### `Copy-ListSchema.ps1`

Use this when the target list does not exist yet or does not have matching custom columns.

Main parameters:

- `-SiteUrl`
- `-ClientId`
- `-SourceListName`
- `-TargetListName`
- `-Overwrite` optional
- `-IncludeHidden` optional

### `Copy-ListItemsWithHistory-RestOnly.ps1`

Use this to copy items and version history.

Main parameters:

- `-SiteUrl`
- `-ClientId`
- `-ListAName`
- `-ListBName`
- `-OriginalIdFieldName` optional (numeric column for the source item id, default `OriginalItemId`)

### `Compare-ListItemVersions.ps1`

Use this to verify the result.

Main parameters:

- `-SiteUrl`
- `-ClientId`
- `-ListAName`
- `-ListBName`
