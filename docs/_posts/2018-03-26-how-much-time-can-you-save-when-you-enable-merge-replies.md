---
layout: post
title: "How much time can you save when you enable Merge Replies?"
---
Here's something you may have been wondering - if you enable the connector's Merge Replies feature, that is, when emails come in that are a Reply to another email that *do not* have a Work Item in the subject how much time are you really saving your analysts? Because once you turn it on, you'll never know it is happening. Emails are simply appended to the correct Work Item from that point forward and everyone benefits as there are no more duplicates to micro-manage. But maybe you're the curious type and/or looking to impress your boss with some time savings statistics on the SMLets Exchange connector you recently implemented. Look no further than the following PowerShell.

```powershell
<#
.SYNOPSIS
    Show how many Work Items were not created by virtue of enabling the SMLets Exchange Connector's Merge Reply feature
.DESCRIPTION
    If you've enabled the $mergeReplies feature on the SMLets Exchange Connector so that emails
    that come into SCSM that are a Reply and DO NOT feature a [Work Item] in the subject. The connector understands
    how to bind the message back to the correct Work Item through the Exchange Conversation ID. This script
    pulls all Incidents from the first of the month, whose source is Email, and then evaluates each Incident that features
    at least two email attachments. The script then evalutes the Subject of each attached email to see if it meets
    the criteria that would have resulted in a merge.

    Each merge occurance is printed out to the screen, the total time to execute this script is shown, followed by the number
    of occurances. The number of occurances available in the $mergedReplies can be used to see how many duplicate
    Work Items WOULD have been created without this feature as well as a multiplier against the minutes it normally
    takes analysts to handle duplicates.
.NOTES
    Author: Adam Dzyacky
#>

$startDate = get-date

#get all the Incidents created this month and whose Source is Email
$scsmMGMTserver = "localhost"
$firstDayOfMonth = Get-Date -day 1 -hour 0 -minute 0 -second 0
$sourceEmailEnum = (get-scsmenumeration -name "IncidentSourceEnum.Email$" -ComputerName $scsmMGMTserver).id
$irClass = get-scsmclass -name "system.workitem.incident$" -computername $scsmMGMTserver
$cType = "Microsoft.EnterpriseManagement.Common.EnterpriseManagementObjectCriteria"
$cString = "Source = '$sourceEmailEnum' and CreatedDate > '$firstDayOfMonth'"
$crit = new-object $cType $cString, $irClass
$incidents = Get-SCSMObject -criteria $crit -computername $scsmMGMTserver

#define the file attachment relationship class and initialize a counter for number occurances
$fileAttachmentRelClass = Get-SCSMRelationshipClass -name "System.WorkItemHasFileAttachment$" -ComputerName $scsmMGMTserver
$mergedReplies = 0

#loop through the Incidents this month
foreach ($incident in $incidents)
{
    #process only the email attachments
    $attachments = get-scsmrelatedobject -SMObject $incident -Relationship $fileAttachmentRelClass -ComputerName $scsmMGMTserver | ?{$_.extension -eq "eml"}

    #as long as we have at least 2 email attachments, evaluate
    if ($attachments.count -ge 2)
    {
        foreach ($attachment in $attachments)
        {
            $ms = New-Object System.IO.MemoryStream($attachment.Content.data,0,$attachment.Content.data.Length)
            $email = [System.Text.Encoding]::ASCII.GetString($ms.ToArray())
            if ($email | Select-String -Pattern "Subject: ([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" | foreach-object { $_.Matches })
            {
                write-output "$($incident.name) had a merge performed"
                $mergedReplies++
            }
        }
    }
}

$endDate = get-date
$endDate - $startDate

write-output "Number of replies merged: $mergedReplies"
```

This script grabs all of your Incidents that came in from Email and then parses each of the attachments (Related Items of that Incident) that are of an email type (so skip over things like jpgs, pngs, and any other attachment type). What's more, making sure that we have at least two email attachments as if we had no check in place we'd be grabbing Incidents that simply were created via email and perhaps no follow up emails took place. Then once we have this new set of data, evaluate each email attachment by looking for a Subject that is a Reply (RE:) but is *missing* a [Work Item] in brackets using the regular expression seen above and no less the identical regular expression used in the actual SMLets Exchange Connector. If we find a match, just increment our $mergedReplies counter by 1. Finally show us the total time it took to run this script and the number of emails it merged into correct Work Items. Be prepared, this could take awhile to run so if you are even the least bit concerned about performance run this after hours or just perform an initial GET of the incidents without processing the attachments by taking the initial block of code and looking at $incidents.count property to get an estimate.

With this value now available in the $mergedReplyCount variable you can multiply this by the number of minutes it takes your analysts to perform the work of identifying duplicates and/or working with each other to confirm duplicates, opening the Incident form, resolving it, possibly marking the resolution in a special way so you know it is in fact a duplicate to avoid reporting on, and finally only possibly relating this new duplicate IR to the original IR for reporting.

Some quick numbers though - let's say this feature engaged 150 times in a month and it takes your analysts 5 minutes to the do the aforementioned work I just described. Not only is it 150 Work Items the connector prevented from being created. That's 12.5 hours of Work Item micro-management in SCSM. That's 12.5 hours not Resolving Incidents or Fulfilling Service Requests. So whose got two thumbs pointed towards them and just removed some SCSM administrative overhead? *insert your name here*
