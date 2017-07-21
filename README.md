# Simple Sharepoint List Cmdlets
Powershell cmdlets for manipulating Sharepoint lists. See function comments in code for parameter details.

# Usage
Sharepoint Online SDK must be installed
https://www.microsoft.com/en-us/download/details.aspx?id=42038

Sharepoint Client DLLs must be loaded from main script before calling any of these cmdlets. Use the following two commands (update paths as needed):
```powershell
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
```
Then import the module and you're good to go!
```powershell
Import-Module "C:\Path\To\SharepointCmdlets.psm1"
```

# Cmdlets

## New-SPContext
Connects to Sharepoint and returns context

Example:
```powershell
$c = New-SPContext -URL "https://mytenant.sharepoint.com/mysite" -UserID "MyUser@MyTenant.com" -Password $SecurePassword
```

## Get-SPList
Returns a list object

Example:
```powershell
$myList = Get-SPList -Context $c -ListName "My List"
```

## Get-SPListFields
Returns all available fields in a list. Fields may not necessarily be accessed via PowerShell by their friendly names, so this cmdlet is useful for discovering how to reference your fields.

Example:
```powershell
Get-SPListFields -List $myList
```

## Get-SPItemArray
Converts an array of list items into an array useable by Sharepoint. You must specify which list fields you want in the array. This function is mainly used by Get-SPListItems, it doesn't have much use on its own.

Example:
```powershell
$objArray = Get-SPItemArray -Items $myItems -fields $myFields
```

## Get-SPListItems
Gets all items in a Sharepoint list and returns them as a Powershell array. Fields to include in the array must be specified. It is recommended to always include ID and Title at a minimum.

Example:
```powershell
$myArray = Get-SPListItems -Context $c -List $myList -Fields "ID","Title","Description","Notes"
```

## Get-SPListItem
Searches a Sharepoint list for a single item based on the provided value for a specified field. Fields to include in the result must be specified. It is recommend to always include ID and Title at a minimum

Example:
```powershell
$myItem = Get-SPListItem -Context $c -List $myList -SearchField "Title" -Value "Item 1" -Fields "ID","Title","Description","Notes"
```

## Set-SPListItem
Sets the field of a specific list item to a value. Item must be referenced by its internal Sharepoint ID.

Example:
```powershell
Set-SPListItem -Context $c -List $myList -ItemID $myItem.ID -Field "Description" -Value "This is my item. There are many others like it, but this one belongs to me."
```

## New-SPListItem
Creates a new list item.

Example:
```powershell
New-SPListItem -Context $c -List $myList -ItemTitle "Item 2"
```
