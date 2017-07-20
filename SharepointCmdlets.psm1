
###
# Simple Sharepoint List Cmdlets
# Created by Chris Thayer - 2017
# http://www.tacoisland.net
###

### Prerequisites ###
# Sharepoint Online SDK
# https://www.microsoft.com/en-us/download/details.aspx?id=42038
#
# Sharepoint Online SDK DLLs must be loaded by main script before these cmdlets can be used
# Use the following commands (update paths as needed):
# Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll" 
# Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
###


Function New-SPContext($URL, $UserID, $Password) {
<#
    New-SPContext

    Connects to Sharepoint and returns context

    URL        The Sharepoint site URL
    UserID     User ID to connect to sharepoint with
    Password   Password for UserID as SecureString
#>
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($URL)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserID, $Password)
    $context
}



Function Get-SPList($Context, $ListName) {
<#
    Get-SPList

    Returns a list object

    Context    The Sharepoint context (returned by New-SPContext)
    ListName   Name of the list to return
#>
    $list = $Context.Web.Lists.GetByTitle($ListName)
    $Context.Load($list)
    $Context.Load($list.Fields)
    $Context.ExecuteQuery()
    $list
}



Function Get-SPListFields($list) {
<#
    Get-SPListFields

    Returns list of all available fields (columns) in a SharePoint list

    List     The list object (returned by Get-SPList)
#>
    ForEach ($f in $list.fields) {
        $f.InternalName
    }    
}



Function Get-SPItemArray ($items, $fields) {
<#
    Get-SPItemArray

    Converts SharePoint list items into a useable array

    Items   Array of all items in the list (returned by list.GetItems)
    Fields  Array of fields to return
#>
    # Create empty array to contain all list items as objects
    $objArray = @()

    # Create empty array to contain all fields to return
    $fieldArray = @()

    # Read all fields that have been passed to the function into our fieldArray
    ForEach ($f in $fields) {
        $fieldArray += $f
    }

    # Loop through each SharePoint item in the list and add to our objArray
    ForEach ($i in $items) {

        # Create new empty object. This will be our item with an array of fields.
        $obj = new-object PSObject
        
        # Loop through each field and add to object array
        For ($x = 0; $x -lt $fields.count; $x++) {
            
            # Add field to single object
            # $objArray[x] is the field we are dealing with
            # $i[$fields[$x]] is the data from the array element in the individual $i object we are working with
            $obj | add-member NoteProperty $fieldArray[$x]($i[$fields[$x]])

        }

        # Add our new object to the array of objects
        $objArray += $obj
    }

    $objArray
}



Function Get-SPListItems($Context, $List, $Fields) {
<#
    Get-SPListItems

    Returns all items in a list as an array of fields

    Context The Sharepoint context (returned by New-SPContext)
    List    The list to convert to an array
    Fields  The fields to return for each item in the list.
            See Get-SPListFields to discover all available fields for the list.
#>
    $query = new-object Microsoft.SharePoint.Client.CamlQuery
    $items = $List.GetItems($query)
    $Context.load($items)
    $Context.ExecuteQuery()

    $objArray = Get-SPItemArray -items $items -fields $fields

    $objArray
}



Function Get-SPListItem($Context, $List, $SearchField, $Value, $Fields) {
<#
    Get-SPListItem

    Returns a single item from a Sharepoint list

    Context     The Sharepoint context (returned by New-SPContext)
    List        The list to search
    SearchField The field (column) to search
    Value       The value to match in SearchField
    Fields      Array of fields to return
                See Get-SPListFields to discover all available fields for the list.
#>
    $query = new-object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXML = "<View><Query><Where><Eq><FieldRef Name='$SearchField'/><Value Type='Text'>$Value</Value></Eq></Where></Query></View>"

    $items = $list.GetItems($query)
    $Context.load($items)
    $Context.ExecuteQuery()

    $objArray = Get-SPItemArray -items $items -fields $fields

    $objArray
}



Function Set-SPListItem($Context, $List, $ItemID, $Field, $Value) {
<#
    Set-SPListItem

    Sets the field of a specific list item to a value

    Context    The Sharepoint context (returned by New-SPContext)
    List       List object containing the item to update
    ItemID     The ID of the list item to update
               It is recommended to always include "ID" as one of the fields to return from Get-SPListItem(s)
    Field      Name of the field to set
    Value      Value to assign to the field
#>
    $item = $List.GetItemByID($ItemID)
    $Context.Load($item)
    $Context.ExecuteQuery()
  
    $Item[$Field] = $Value
    $Item.Update()
  
    try {
        $Context.ExecuteQuery()
    }
    catch [Net.WebException] { 
        $_.Exception.ToString()
    }
}



Function New-SPListItem($Context, $List, $ItemTitle) {
<#
	New-SPListItem
	
	Creates a new item in the Sharepoint list
	
	Context    The Sharepoint context (returned by New-SPContext)
	List       List object where the new item will be created
	ItemTitle  Identifier for new item
#>
	$newItem = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	
	$item = $List.AddItem($newItem)
	$item['Title'] = $ItemTitle
	
	$item.Update()
	$list.Update()
	
	try {
		$Context.ExecuteQuery()
	}
    catch [Net.WebException] { 
        $_.Exception.ToString()
    }
}
