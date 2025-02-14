# 📂 README: PowerShell Script for Counting Subfolders in SharePoint Online Document Library

## 📝 Overview
This PowerShell script allows users to *count the number of subfolders* in a SharePoint Online document library and update a custom column named FolderCount. The script works recursively and ensures accurate statistics at each folder level.

### *✨ Key Features*:
- ✅ *Counts only folders* (ignoring files) in a SharePoint document library.
- 🔄 *Recursively processes subfolders* to update counts at all levels.
- 🔐 *Uses PnP PowerShell*, ensuring a secure and efficient connection to SharePoint Online.
- 📊 *Updates SharePoint metadata* in real time by setting the FolderCount field.

---
## 🔧 *Requirements*
Before running the script, ensure you have the following:

### *Prerequisites*:
1. 📌 *PnP PowerShell Module*
   powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   
2. 🔑 *SharePoint Online Access*
   - You must have the necessary permissions to connect and update the document library.
3. 🗂 *FolderCount Column*
   - Ensure that your SharePoint document library has a column named FolderCount (Type: Number).

---
## 🚀 *PowerShell Script*

powershell
# Function to get the number of subfolders (ignoring files) recursively
Function Get-SPOFolderStats
{
    [cmdletbinding()]
    param
    (
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)][Microsoft.SharePoint.Client.Folder]$Folder
    )
    
    # Get Sub-folders of the folder
    Get-PnPProperty -ClientObject $Folder -Property ServerRelativeUrl, Folders | Out-Null

    # Get the SiteRelativeUrl
    $Web = Get-PnPWeb -Includes ServerRelativeUrl
    $SiteRelativeUrl = $Folder.ServerRelativeUrl -replace "$($web.ServerRelativeUrl)", [string]::Empty

    # Calculate subfolder count only (no file count)
    $SubFolderCount = Get-PnPFolderItem -FolderSiteRelativeUrl $SiteRelativeUrl -ItemType Folder | Measure-Object | Select -ExpandProperty Count

    # Fetch the List Item corresponding to the folder
    $ListItem = Get-PnPListItem -List $ListName | Where-Object { $_["FileRef"] -eq $Folder.ServerRelativeUrl }

    # Check if the list item exists and update the FolderCount field
    if ($ListItem) {
        Set-PnPListItem -List $ListName -Identity $ListItem.Id -Values @{"FolderCount" = $SubFolderCount}
        Write-Host "✅ Updated FolderCount for $($Folder.ServerRelativeUrl): $SubFolderCount"
    } else {
        Write-Host "⚠ List item for folder $($Folder.ServerRelativeUrl) not found."
    }

    # Process Sub-folders recursively
    ForEach($SubFolder in $Folder.Folders)
    {
        Get-SPOFolderStats -Folder $SubFolder
    }
}

# Set the SharePoint Site URL and List Name
$SiteURL = "https://futurrizoninterns.sharepoint.com/sites/MentalHealthCareWebApplication1"
$ListName = "CustomDocumentLibrary"  # Name of your document library

# Connect to SharePoint Online using Web Login
Connect-PnPOnline -URL $SiteURL -UseWebLogin

# Fetch the Document Library and start counting subfolders
$Library = Get-PnPList -Identity $ListName -Includes RootFolder

# Call the Function to Get the Library Statistics - Number of subfolders at each level
$Library.RootFolder | Get-SPOFolderStats

Write-Host "🎉 Folder-wise statistics updated successfully!"


---
## 🔍 *How to Use the Script*

1. 🛠 *Modify the variables*:
   - Set your *SharePoint Online site URL* in $SiteURL.
   - Define your *document library name* in $ListName.
2. 🔐 *Ensure you have the necessary permissions* to update SharePoint metadata.
3. ▶ *Run the script* in PowerShell:
   powershell
   .\YourScriptName.ps1
   
4. 📊 *Check your SharePoint library*:
   - The FolderCount column will be updated with the number of subfolders.

---
## ❓ *Troubleshooting & FAQs*

### ❌ **1. What if FolderCount column is missing?**
- ➡ Create a *Number column* in your document library and name it FolderCount.

### 🔑 *2. What if the script fails to connect?*
- ➡ Ensure you have *PnP PowerShell installed* and use -UseWebLogin to authenticate.

### 📁 *3. Can I use this on multiple document libraries?*
- ✅ Yes! Modify $ListName and re-run the script for each library.

### ⚠ **4. What if I see List item for folder not found?**
- ➡ Ensure the folders are properly indexed in SharePoint.

---
## 🌎 *What is this script for?*
This script is designed for *SharePoint Online Folder Count, **PnP PowerShell Get Subfolders, and **SharePoint Document Library Folder Statistics*. If you are searching for:
- 🔍 "How to count folders in SharePoint using PowerShell"
- 🔍 "PnP PowerShell get subfolder count in SharePoint"
- 🔍 "Update SharePoint folder metadata automatically"
- 🔍 "Recursive folder count in SharePoint Online"

Then this script is the *perfect solution!* 🚀
