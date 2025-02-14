$SiteURL = "https://futurrizoninterns.sharepoint.com/sites/lookUpDataTesting"
$ListName = "Document Management Library 2"

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

    # Update the FolderCount column in SharePoint (update folder's "FolderCount" field)
    if ($ListItem) {
        Set-PnPListItem -List $ListName -Identity $ListItem.Id -Values @{"FolderCount" = $SubFolderCount}
        Write-Host "Updated FolderCount for $($Folder.ServerRelativeUrl): $SubFolderCount"
    } else {
        Write-Host "List item for folder $($Folder.ServerRelativeUrl) not found."
    }

    # Process Sub-folders
    ForEach($SubFolder in $Folder.Folders)
    {
        Get-SPOFolderStats -Folder $SubFolder
    }
}

# Connect to SharePoint Online using Web Login
Connect-PnPOnline $SiteURL -UseWebLogin

# Call the Function to Get the Library Statistics - Number of subfolders at each level
$FolderStats = Get-PnPList -Identity $ListName -Includes RootFolder | Select -ExpandProperty RootFolder | Get-SPOFolderStats
