---
layout: post
title: "The SMlets Exchange Connector at its core"
author: Adam
---
  
If you're entering PowerShell, SMLets, or no less this connector for the first time a question you may have is simply "Where's the part that processes the mail?" In short - it's at the bottom of the script. But even still it's filled with a lot of conditional processing to handle the different Work Item types as well as different email types. The following are all of the same variables and logic from the original SMLets Exchange Connector just condensed down into this version again for readability and examining the whole premise of said connector. So let's dive in and see how this works *without* talking about Service Manager.

```powershell
#define the variables we'll need to connect, if you use impersonation instead of windows authentication you'll need to supply credentials
$exchangeEWSAPIPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
$exchangeAuthenticationType = "windows"
$workflowEmailAddress = "some.address@corp.net"
$username = ""
$password = ""
$domain = ""
$UseAutodiscover = $true
$ExchangeEndpoint = ""

#load Exchange Web Services assembly and define connection to Exchange
[void] [Reflection.Assembly]::LoadFile("$exchangeEWSAPIPath")
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
$exchangeService.UseDefaultCredentials = $true
$exchangeService.AutodiscoverUrl($workflowEmailAddress)

#determine how to connect to Exchange based on the variables that were initially defined
switch ($exchangeAuthenticationType)
{
    "impersonation" {$exchangeService.Credentials = New-Object Net.NetworkCredential($username, $password, $domain)}
    "windows" {$exchangeService.UseDefaultCredentials = $true}
}
if ($UseAutoDiscover -eq $true) {
    $exchangeService.AutodiscoverUrl($workflowEmailAddress)
}
else {
    $exchangeService.Url = [System.Uri]$ExchangeEndpoint
}

#define search parameters, search on the defined classes and get messages that are older than the current time
$inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
$inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)
$itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
$mimeContentSchema = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
$dateTimeItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
$now = get-date
$searchFilter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo -ArgumentList $dateTimeItem,$now

#load the items in the Inbox using the the $searchFilter from above, in this case - load ANY kind of message (calendar, encrypted, regular email, OOO, etc.)
$inbox = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView)

#loop through the inbox
foreach ($message in $inbox)
{
    $message
}
```

It probably goes without saying, but the loop through inbox part is the area that contains all of the logic for how email is processed in the actual connector. Does the subject contain something like that looks like [IRxxxxx]? Then go find the Work Item and update it. Does the subject not contain something like [IRxxxxx]? Then create a New Work Item.
  
This is the basis of how the whole thing works. Connect to Exchange, get the items in the inbox that you want to work with, and loop through them one by one. The rest of the functions in the script call one another so as to create or update Work Items in Service Manager. Not only can you use this script to test your connection to Exchange and said account, but you can also use it to begin testing scenarios for your own modifications to the SMLets Exchange Connector.
