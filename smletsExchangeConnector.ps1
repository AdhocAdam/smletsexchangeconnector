<#
.SYNOPSIS
Provides SCSM Exchange Connector functionality through PowerShell

.DESCRIPTION
This PowerShell script/runbook aims to address shortcomings and wants in the
out of box SCSM Exchange Connector as well as help enable new functionality around
Work Item scenarios, runbook automation, 3rd party customizations, and
enabling other organizational level processes via email

.NOTES
Author: Adam Dzyacky
Contributors: Martin Blomgren, Leigh Kilday
Reviewers: Tom Hendricks, Brian Weist
Inspiration: The Cireson Community, Anders Asp, Stefan Roth, and (of course) Travis Wright for SMlets examples
Requires: PowerShell 4+, SMlets, and Exchange Web Services API (already installed on SCSM workflow server).
3rd party option: If you're a Cireson customer and make use of their paid SCSM Portal with HTML Knowledge Base this will work as is
    if you aren't, you'll need to create your own Type Projection for Change Requests for the Add-ChangeRequestComment
    function. Navigate to that function to read more. If you don't make use of their HTML KB, you'll want to keep $searchCiresonHTMLKB = $false
Misc: The Release Record functionality does not exist in this as no out of box (or 3rd party) Type Projection exists to serve this purpose.
    You would have to create your own Type Projection in order to leverage this.
Version: 1.2 = created Send-EmailFromWorkflowAccount for future functions to leverage the SCSM workflow account defined therein
                updated Search-CiresonKnowledgeBase to use Send-EmailFromWorkflowAccount
                created $exchangeAuthenticationType so as to introduce Windows Authentication or Impersonation to bring to closer parity with stock EC connector
                expanded email processing loop to prepare for things other than IPM.Note message class (i.e. Calendar appointments, custom message classes per org.)
                created Schedule-WorkItem function to enable setting Scheduled Start/End Dates on Work Items based on the Calendar Start/End times.
                    introduced configuration variable for this feature ($processCalendarAppointment)
                created config variable for LoggingLevel
                    could just reference the same registry key for the native Exchange Connector
                    need to build what the levels of logging represent and create said functions
                updated Attach-EmailToWorkItem so the Exchange Conversation ID is written into the Description as "ExchangeConversationID:$id;"
                issue on Attach-EmailToWorkItem/Attach-FileToWorkItem where the "AttachedBy" relationship was using the wrong variable
                created Verify-WorkItem to attempt to begin identifying potentially quickly responded messages/append them to
                    current/recently created Work Items. Trying to address scenario where someone emails WF and CC's others. If the others
                    reply back before a notification about the current Work Item ID goes out (or they ignore it), they queue more messages
                    for the connector and in turn create more default work items rather than updating the "original" thread/Work Item that
                    was created in the same/previous processing loop. Also looking to use this function to potentially address 3rd party
                    ticketing systems.
                    Introduced configuration variable for this feature ($mergeReplies)
Version: 1.1 = GitHub issue raised on updating work items. Per discussion was pinpointed to the
                Get-WorkItem function wherein passed in values were including brackets in the search (i.e. [IRxxxx] instead of IRxxxx). Also
                updated the email subject matching regex, so that the Update-WorkItem took the $result.id instead of the $matches[0]. Again, this
                ensures the brackets aren't passed when performing the search/update.
#>

#region #### Configuration ####
#define the an SCSM management server, this could be a remote name or localhost
$scsmMGMTServer = ""

#define/use SCSM WF credentials
#$exchangeAuthenticationType - "windows" or "impersonation" are valid inputs here.
    #Windows will use the credentials that start this script in order to authenticate to Exchange and retrieve messages
        #choosing this option only requires the $workflowEmailAddress variable to be defined
        #this is ideal if you'll be using Task Manager or SMA to initiate this
    #Impersonation will use the credentials that are defined here to connect to Exchange and retrieve messages
        #choosing this option requires the $workflowEmailAddress, $username, $password, and $domain variables to be defined
$exchangeAuthenticationType = "windows"
$workflowEmailAddress = ""
$username = ""
$password = ""
$domain = ""

#defaultNewWorkItem = set to either "ir" or "sr"
#minFileSizeInKB = Set the minimum file size in kilobytes to be attached to work items
#createUsersNotInCMDB = If someone from outside your org emails into SCSM this allows you to take that email and create a User in your CMDB
#includeWholeEmail = If long chains get forwarded into SCSM, you can choose to write the whole email to a single action log entry OR the beginning to the first finding of "From:"
#attachEmailToWorkItem = If $true, attach email as an *.eml to each work item. Additionally, write the Exchange Conversation ID into the Description of the Attachment object
#fromKeyword = If $includeWholeEmail is set to true, messages will be parsed UNTIL they find this word
$defaultNewWorkItem = "ir"
$defaultIRTemplate = Get-SCSMObjectTemplate -DisplayName "IR Template Name Goes Here" -computername $scsmMGMTServer
$defaultSRTemplate = Get-SCSMObjectTemplate -DisplayName "SR Template Name Goes Here" -computername $scsmMGMTServer
$minFileSizeInKB = "25"
$createUsersNotInCMDB = $true
$includeWholeEmail = $false
$attachEmailToWorkItem = $false
$fromKeyword = "From"

#processCalendarAppointment = If $true, scheduling appointments with the Workflow Inbox where a [WorkItemID] is in the Subject will
    #set the Scheduled Start and End Dates on the Work Item per the Start/End Times of the calendar appointment
#mergeReplies = If $true, emails that are Replies (signified by RE: in the subject) will attempt to be matched to a Work Item in SCSM by their
    #Exchange Conversation ID and will also override $attachEmailToWorkItem to be $true if set to $false
$processCalendarAppointment = $false
$mergeReplies = $false

#optional, enable KB search of your Cireson HTML KB
#this uses the now depricated Cireson KB API Search by Text, it works as of v7.x but should be noted it could be entirely removed in future portals
#$ciresonPortalServer = URL that will be used to search for KB articles via invoke-webrequest. Make sure to leave the "/" after your tld!
#$ciresonAccountGUID = This is the GUID/ID of a User within the ServiceManagement DB that will be used to search the knowledge base from the CI$User table
#$ciresonPortalWindowsAuth = how invoke-webrequest should attempt to auth to your portal server.
    #Leave true if your portal uses Windows Auth, change to False for Forms authentication.
    #If using forms, you'll need to set the ciresonPortalUsername and Password variables. For ease, you could set this equal to the username/password defined above
$searchCiresonHTMLKB = $false
$ciresonPortalServer = "https://portalserver.domain.tld/"
$ciresonAccountGUID = "11111111-2222-3333-4444-AABBCCDDEEFF"
$ciresonKBLanguageCode = "ENU"
$ciresonPortalWindowsAuth = $true
$ciresonPortalUsername = ""
$ciresonPortalPassword = ""

#optional, enable SCOM functionality
#To prevent unintended 3rd parties from gaining knowledge of your environment via SCSM about SCOM, you must choose to set either an AD Group that the sender must be a part of
#or you must manually define email addresses of users allowed to make these requests
#enableSCOMIntegration = set to $true or $false to enable this functionality
#scomMGMTServer = set equal to the name of your scom management server
#approvedMemberTypeForSCOM = set to either "users" or "group"
#approvedADGroupForSCOM = if approvedUsersForSCOM = group, set this to the AD Group that contains groups/members that are allowed to make SCOM email requests
    #this approach allows you control access through Active Directory
#approvedUsersForSCOM = if approvedUsersForSCOM = users, set this to a comma seperated list of email addresses that are allowed to make SCOM email requests
    #this approach allows you to control through this script
$enableSCOMIntegration = $false
$scomMGMTServer = ""
$approvedMemberTypeForSCOM = "group"
$approvedADGroupForSCOM = "my custom AD SCOM group"
$approvedUsersForSCOM = "myfirst.email@domain.com", "mysecond.address@domain.com"

#define SCSM Work Item keywords to be used
$acknowledgedKeyword = "acknowledge"
$reactivateKeyword = "reactivate"
$resolvedKeyword = "resolved"
$closedKeyword = "closed"
$holdKeyword = "hold"
$cancelledKeyword = "cancelled"
$takeKeyword = "take"
$completedKeyword = "completed"
$skipKeyword = "skipped"
$approvedKeyword = "approved"
$rejectedKeyword = "rejected"

#define the path to the Exchange Web Services API. the following is the default install directory for EWS API
$exchangeEWSAPIPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"

#enable logging per standard Exchange Connector registry keys
#valid options on that registry key are 1 to 7 where 7 is the most verbose
#$loggingLevel = (Get-ItemProperty "HKLM:\Software\Microsoft\System Center Service Manager Exchange Connector" -ErrorAction SilentlyContinue).LoggingLevel
$loggingLevel = 1

#endregion

#region #### SCSM Classes ####
$irClass = get-scsmclass "System.WorkItem.Incident$" -computername $scsmMGMTServer
$srClass = get-scsmclass "System.WorkItem.ServiceRequest$" -computername $scsmMGMTServer
$prClass = get-scsmclass "System.WorkItem.Problem$" -computername $scsmMGMTServer
$crClass = get-scsmclass "System.Workitem.ChangeRequest$" -computername $scsmMGMTServer
$rrClass = get-scsmclass "System.Workitem.ReleaseRecord$" -computername $scsmMGMTServer
$maClass = get-scsmclass "System.WorkItem.Activity.ManualActivity$" -computername $scsmMGMTServer
$raClass = get-scsmclass "System.WorkItem.Activity.ReviewActivity$" -computername $scsmMGMTServer
$paClass = get-scsmclass "System.WorkItem.Activity.ParallelActivity$" -computername $scsmMGMTServer
$saClass = get-scsmclass "System.WorkItem.Activity.SequentialActivity$" -computername $scsmMGMTServer
$daClass = get-scsmclass "System.WorkItem.Activity.DependentActivity$" -computername $scsmMGMTServer

$raHasReviewerRelClass = Get-SCSMRelationshipClass "System.ReviewActivityHasReviewer$" -computername $scsmMGMTServer
$raReviewerIsUserRelClass = Get-SCSMRelationshipClass "System.ReviewerIsUser$" -computername $scsmMGMTServer
$raVotedByUserRelClass = Get-SCSMRelationshipClass "System.ReviewerVotedByUser$" -computername $scsmMGMTServer

$userClass = get-scsmclass "System.User$" -computername $scsmMGMTServer
$domainUserClass = get-scsmclass "System.Domain.User$" -computername $scsmMGMTServer
$notificationClass = get-scsmclass "System.Notification.Endpoint$" -computername $scsmMGMTServer

$irLowImpact = Get-SCSMEnumeration "System.WorkItem.TroubleTicket.ImpactEnum.Low$" -computername $scsmMGMTServer
$irLowUrgency = Get-SCSMEnumeration "System.WorkItem.TroubleTicket.ImpactEnum.Low$" -computername $scsmMGMTServer
$irActiveStatus = Get-SCSMEnumeration "IncidentStatusEnum.Active$" -computername $scsmMGMTServer

$affectedUserRelClass = get-scsmrelationshipclass "System.WorkItemAffectedUser$" -computername $scsmMGMTServer
$assignedToUserRelClass  = Get-SCSMRelationshipClass "System.WorkItemAssignedToUser$" -computername $scsmMGMTServer
$createdByUserRelClass = Get-SCSMRelationshipClass "System.WorkItemCreatedByUser$" -computername $scsmMGMTServer
$workResolvedByUserRelClass = Get-SCSMRelationshipClass "System.WorkItem.TroubleTicketResolvedByUser$" -computername $scsmMGMTServer
$wiRelatesToCIRelClass = Get-SCSMRelationshipClass "System.WorkItemRelatesToConfigItem$" -computername $scsmMGMTServer
$wiRelatesToWIRelClass = Get-SCSMRelationshipClass "System.WorkItemRelatesToWorkItem$" -computername $scsmMGMTServer
$wiContainsActivityRelClass = Get-SCSMRelationshipClass "System.WorkItemContainsActivity$" -computername $scsmMGMTServer
$sysUserHasPrefRelClass = Get-SCSMRelationshipClass "System.UserHasPreference$" -ComputerName $scsmMGMTServer

$fileAttachmentClass = Get-SCSMClass -Name "System.FileAttachment$" -computername $scsmMGMTServer
$fileAttachmentRelClass = Get-SCSMRelationshipClass "System.WorkItemHasFileAttachment$" -computername $scsmMGMTServer
$fileAddedByUserRelClass = Get-SCSMRelationshipClass "System.FileAttachmentAddedByUser$" -ComputerName $scsmMGMTServer
$managementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup $scsmMGMTServer

$irTypeProjection = Get-SCSMTypeProjection "system.workitem.incident.projectiontype$" -computername $scsmMGMTServer
$srTypeProjection = Get-SCSMTypeProjection "system.workitem.servicerequestprojection$" -computername $scsmMGMTServer

$userHasPrefProjection = Get-SCSMTypeProjection "System.User.Preferences.Projection$" -computername $scsmMGMTServer
#endregion

#region #### Exchange Connector Functions ####
function New-WorkItem ($message, $wiType, $returnWIBool) 
{
    $from = $message.From
    $to = $message.To
    $cced = $message.CC
    $title = $message.subject
    $description = $message.body

    #if the message is longer than 4000 characters take only the first 4000.
    if ($description.length -ge "4000")
    {
        $description = $description.substring(0,4000)
    }

    #find Affected User from the From Address
    $relatedUsers = @()
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$from'"  -computername $scsmMGMTServer
    if ($userSMTPNotification) 
    { 
        $affectedUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    }
    else
    {
        if ($createUsersNotInCMDB -eq $true)
        {
            $affectedUser = create-userincmdb $from
        }
    }

    #find Related Users (To)       
    if ($to -gt $0)
    {
        $x = 0
        while ($x -lt $to.count)
        {
            $ToSMTP = $to[$x]
            $userToSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($ToSMTP.address)'"  -computername $scsmMGMTServer
            if ($userToSMTPNotification) 
            { 
                $relatedUser = (Get-SCSMRelationshipObject -ByTarget $userToSMTPNotification -computername $scsmMGMTServer).sourceObject 
                $relatedUsers += $relatedUser
            }
            else
            {
                if ($createUsersNotInCMDB -eq $true)
                {
                    $newUser = create-userincmdb $to[$x]
                    $relatedUsers += $relatedUser
                }
            }
            $x++
        }
    }
    
    #find Related Users (Cc)         
    if ($cced -gt $0)
    {
        $x = 0
        while ($x -lt $cced.count)
        {
            $ccSMTP = $cced[$x]
            $userCCSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($ccSMTP.address)'" -computername $scsmMGMTServer
            if ($userCCSMTPNotification) 
            { 
                $relatedUser = (Get-SCSMRelationshipObject -ByTarget $userCCSMTPNotification -computername $scsmMGMTServer).sourceObject 
                $relatedUsers += $relatedUser
            }
            else
            {
                if ($createUsersNotInCMDB -eq $true)
                {
                    $newUser = create-userincmdb $cced[$x]
                    $relatedUsers += $relatedUser
                }
            }
            $x++
        }
    }

    #create the Work Item based on the globally defined Work Item type and Template
    switch ($defaultNewWorkItem) 
    {
        "ir" {
                    $newWorkItem = New-SCSMObject -Class $irClass -PropertyHashtable @{"ID" = "IR{0}"; "Status" = $irActiveStatus; "Title" = $title; "Description" = $description; "Classification" = $null; "Impact" = $irLowImpact; "Urgency" = $irLowUrgency; "Source" = "IncidentSourceEnum.Email$"} -PassThru -computername $scsmMGMTServer
                    $irProjection = Get-SCSMObjectProjection -ProjectionName $irTypeProjection.Name -Filter "Name -eq $($newWorkItem.Name)" -computername $scsmMGMTServer
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Set-SCSMObjectTemplate -Projection $irProjection -Template $defaultIRTemplate -computername $scsmMGMTServer
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                        New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk -computername $scsmMGMTServer
                        }
                    }
                    if ($searchCiresonHTMLKB -eq $true)
                    {
                        Search-CiresonKnowledgeBase $message $newWorkItem
                    }
                }
        "sr" {
                    $newWorkItem = new-scsmobject -class $srClass -propertyhashtable @{"ID" = "SR{0}"; "Title" = $title; "Description" = $description; "Status" = "ServiceRequestStatusEnum.Submitted$"} -PassThru -computername $scsmMGMTServer
                    $srProjection = Get-SCSMObjectProjection -ProjectionName $srTypeProjection.Name -Filter "Name -eq $($newWorkItem.Name)" -computername $scsmMGMTServer
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Set-SCSMObjectTemplate -projection $srProjection -Template $defaultSRTemplate -computername $scsmMGMTServer
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                        New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk -computername $scsmMGMTServer
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk -computername $scsmMGMTServer
                        }
                    }
                    if ($searchCiresonHTMLKB -eq $true)
                    {
                        Search-CiresonKnowledgeBase $message $newWorkItem
                    }
                } 
    }

    if ($returnWIBool -eq $true)
    {
        return $newWorkItem
    }
}

function Update-WorkItem ($message, $wiType, $workItemID) 
{
    #determine the comment to add and ensure it's less than 4000 characters
    if ($includeWholeEmail -eq $true)
    {
        $commentToAdd = $message.body
        if ($commentToAdd.length -ge "4000")
        {
            $commentToAdd.substring(0, 4000)
        }
    }
    else
    {
        $fromKeywordPosition = $message.Body.IndexOf("$fromKeyword" + ":")
        if (($fromKeywordPosition -eq $null) -or ($fromKeywordPosition -eq -1))
        {
            $commentToAdd = $message.body
            if ($commentToAdd.length -ge "4000")
            {
                $commentToAdd.substring(0, 4000)
            }
        }
        else
        {
            $commentToAdd = $message.Body.substring(0, $fromKeywordPosition)
            if ($commentToAdd.length -ge "4000")
            {
                $commentToAdd.substring(0, 4000)
            }
        }
    }

    #determine who left the comment
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.From)'" -computername $scsmMGMTServer
    if ($userSMTPNotification) 
    { 
        $commentLeftBy = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
    }
    else
    {
        if ($createUsersNotInCMDB -eq $true)
        {
            $commentLeftBy = create-userincmdb $from
        }
    }

    #add any attachments
    if ($message.Attachments)
    {
        Attach-FileToWorkItem $message $workItemID
    }

    #update the work item with the comment and/or action
    switch ($wiType) 
    {
        #### primary work item types ####
        "ir" {
                    $workItem = get-scsmobject -class $irClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $affectedUserSMTP.TargetAddress {Add-IncidentComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -AnalystComment $false -isPrivate $false}
                        $assignedToSMTP.TargetAddress {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $false};Add-IncidentComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $isPrivateBool}
                        default {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $null};Add-IncidentComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $isPrivateBool}
                    }
                    #take action on the Work Item if neccesary
                    switch -Regex ($commentToAdd)
                    {
                        "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() -computername $scsmMGMTServer}}
                        "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Resolved$" -computername $scsmMGMTServer; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                        "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Closed$" -computername $scsmMGMTServer}
                        "\[$takeKeyword]" {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                        {("\[$reactivateKeyword]") -and ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved")} {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Active$" -computername $scsmMGMTServer}
                        {("\[$reactivateKeyword]") -and ($workItem.Status.Name -eq "IncidentStatusEnum.Closed")} {if($message.Subject -match "[I][R][0-9]+"){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", "")}; $returnedWorkItem = New-WorkItem $message "ir" $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk -computername $scsmMGMTServer}catch{}}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                } 
        "sr" {
                    $workItem = get-scsmobject -class $srClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo -computername $scsmMGMTServer| ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    switch ($message.From)
                    {
                        $affectedUserSMTP.TargetAddress {Add-ServiceRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -AnalystComment $false -isPrivate $false}
                        $assignedToSMTP.TargetAddress {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $false};Add-ServiceRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $isPrivateBool}
                        default {if($commentToAdd -match "#private"){$isPrivateBool = $true}else{$isPrivateBool = $null};Add-ServiceRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $isPrivateBool}
                    }
                    switch -Regex ($commentToAdd)
                    {
                        "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Completed$" -computername $scsmMGMTServer}
                        "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Cancelled$" -computername $scsmMGMTServer}
                        "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Closed$" -computername $scsmMGMTServer}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                } 
        "pr" {
                    $workItem = get-scsmobject -class $prClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $assignedToSMTP.TargetAddress {Add-ProblemComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $false}
                        default {Add-ProblemComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $true -isPrivate $null}
                    }
                    #take action on the Work Item if neccesary
                    switch -Regex ($commentToAdd)
                    {
                         "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Resolved$" -computername $scsmMGMTServer; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                         "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Closed$" -computername $scsmMGMTServer}
                         "\[$takeKeyword]" {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                        {("\[$reactivateKeyword]") -and ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved")} {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Active$" -computername $scsmMGMTServer}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                }
        "cr" {
                    $workItem = get-scsmobject -class $crClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try{$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $assignedToSMTP.TargetAddress {Add-ChangeRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -AnalystComment $true -isPrivate $false}
                        default {Add-ChangeRequestComment -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -isPrivate $false}
                    }
                    #take action on the Work Item if neccesary
                    switch -Regex ($commentToAdd)
                    {
                        "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold$" -computername $scsmMGMTServer}
                        "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled$" -computername $scsmMGMTServer}
                        "\[$takeKeyword]" {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy -computername $scsmMGMTServer -bulk}
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk -computername $scsmMGMTServer
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                }
        
        #### activities ####
        "ra" {
                    $workItem = get-scsmobject -class $raClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    $reviewers = Get-SCSMRelatedObject -SMObject $workItem -Relationship $raHasReviewerRelClass -computername $scsmMGMTServer
                    foreach ($reviewer in $reviewers)
                    {
                        $reviewingUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $reviewer -Relationship $raReviewerIsUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer
                        $reviewingUserSMTP = Get-SCSMRelatedObject -SMObject $reviewingUser -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress
                        
                        #approved
                        if (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$approvedKeyword]"))
                        {
                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} -computername $scsmMGMTServer
                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk -computername $scsmMGMTServer
                        }
                        #rejected
                        elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                        {
                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $commentToAdd} -computername $scsmMGMTServer
                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk -computername $scsmMGMTServer
                        }
                        #no keyword, add a comment to parent work item
                        elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and (($commentToAdd -notmatch "\[$approvedKeyword]") -or ($commentToAdd -notmatch "\[$rejectedKeyword]")))
                        {
                            $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                            switch ($parentWorkItem.Classname)
                            {
                                "System.WorkItem.ChangeRequest" {Add-ChangeRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                "System.WorkItem.ServiceRequest" {Add-ServiceRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                                "System.WorkItem.Incident" {Add-IncidentComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            }
                            
                        }
                    }
                }
        "ma" {
                    $workItem = get-scsmobject -class $maClass -filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
                    try {$activityImplementer = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass -computername $scsmMGMTServer).id -computername $scsmMGMTServer} catch {}
                    if ($activityImplementer){$activityImplementerSMTP = Get-SCSMRelatedObject -SMObject $activityImplementer -computername $scsmMGMTServer | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    
                    #completed
                    if (($activityImplementerSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$completedKeyword]"))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Completed$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} -computername $scsmMGMTServer
                    }
                    #skipped
                    elseif (($activityImplementerSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$skipKeyword]"))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Skipped$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} -computername $scsmMGMTServer
                    }
                    #not from the Activity Implementer, add to the MA Notes
                    elseif (($activityImplementerSMTP.TargetAddress -ne $message.From))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} -computername $scsmMGMTServer
                    }
                    #no keywords, add to the Parent Work Item
                    elseif (($activityImplementerSMTP.TargetAddress -eq $message.From) -and (($commentToAdd -notmatch "\[$completedKeyword]") -or ($commentToAdd -notmatch "\[$skipKeyword]")))
                    {
                        $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                        switch ($parentWorkItem.Classname)
                        {
                            "System.WorkItem.ChangeRequest" {Add-ChangeRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            "System.WorkItem.ServiceRequest" {Add-ServiceRequestComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                            "System.WorkItem.Incident" {Add-IncidentComment -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -AnalystComment $false -IsPrivate $false}
                        }
                            
                    }
                }
    } 
}

function Attach-EmailToWorkItem ($message, $workItemID)
{
    $messageMime = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $MemoryStream = New-Object System.IO.MemoryStream($messageMime.MimeContent.Content,0,$messageMime.MimeContent.Content.Length)

    #Create the attachment object itself and set its properties for SCSM
    $emailAttachment = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
    $emailAttachment.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
    $emailAttachment.Item($fileAttachmentClass, "DisplayName").Value = "message.eml"
    $emailAttachment.Item($fileAttachmentClass, "Description").Value = "ExchangeConversationID:$($message.ConversationID);"
    $emailAttachment.Item($fileAttachmentClass, "Extension").Value =   "eml"
    $emailAttachment.Item($fileAttachmentClass, "Size").Value =        $MemoryStream.Length
    $emailAttachment.Item($fileAttachmentClass, "AddedDate").Value =   [DateTime]::Now.ToUniversalTime()
    $emailAttachment.Item($fileAttachmentClass, "Content").Value =     $MemoryStream
    
    #Add the attachment to the work item and commit the changes
    $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemID'" -computername $scsmMGMTServer
    $WorkItemProjection.__base.Add($emailAttachment, $fileAttachmentRelClass.Target)
    $WorkItemProjection.__base.Commit()
            
    #create the Attached By relationship if possible
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.from)'" -computername $scsmMGMTServer
    if ($userSMTPNotification) 
    { 
        $attachedByUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
        New-SCSMRelationshipObject -Source $emailAttachment -Relationship $fileAddedByUserRelClass -Target $attachedByUser -Bulk
    }
}

#inspired and modified from Stefan Roth here - https://stefanroth.net/2015/03/28/scsm-passing-attachments-via-web-service-e-g-sma-web-service/
function Attach-FileToWorkItem ($message, $workItemId)
{
    foreach ($attachment in $message.Attachments)
    {
        $attachment.Load()
        $base64attachment = [System.Convert]::ToBase64String($attachment.Content)

        #Convert the Base64String back to bytes
        $AttachmentContent = [convert]::FromBase64String($base64attachment)

        #Create a new MemoryStream object out of the attachment data
        $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)

        if (([int]$MemoryStream.Length * 1000) -gt $minFileSizeInKB)
        {
            #Create the attachment object itself and set its properties for SCSM
            $NewFile = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
            $NewFile.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
            $NewFile.Item($fileAttachmentClass, "DisplayName").Value = $attachment.Name
            #$NewFile.Item($fileAttachmentClass, "Description").Value = $attachment.Description
            #$NewFile.Item($fileAttachmentClass, "Extension").Value =   $attachment.Extension
            $NewFile.Item($fileAttachmentClass, "Size").Value =        $MemoryStream.Length
            $NewFile.Item($fileAttachmentClass, "AddedDate").Value =   [DateTime]::Now.ToUniversalTime()
            $NewFile.Item($fileAttachmentClass, "Content").Value =     $MemoryStream
    
            #Add the attachment to the work item and commit the changes
            $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemId'" -computername $scsmMGMTServer
            $WorkItemProjection.__base.Add($NewFile, $fileAttachmentRelClass.Target)
            $WorkItemProjection.__base.Commit()

            #create the Attached By relationship if possible
            $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$($message.from)'" -computername $scsmMGMTServer
            if ($userSMTPNotification) 
            { 
                $attachedByUser = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification -computername $scsmMGMTServer).sourceObject.id -computername $scsmMGMTServer
                New-SCSMRelationshipObject -Source $emailAttachment -Relationship $fileAddedByUserRelClass -Target $attachedByUser -Bulk
            }
        }
    }
}

function Get-WorkItem ($workItemID, $workItemClass)
{
    #removes [] surrounding a Work Item ID if neccesary
    if ($workitemID.StartsWith("[") -and $workitemID.EndsWith("]"))
    {
        $workitemID = $workitemID.TrimStart("[").TrimEnd("]")
    }

    #get the work item
    $wi = get-scsmobject -Class $workItemClass -Filter "Name -eq '$workItemID'" -computername $scsmMGMTServer
    return $wi
}

#courtesy of Leigh Kilday. Modified.
function Get-SCSMWorkItemParent
{
    [CmdLetBinding()]
    PARAM (
        [Parameter(ParameterSetName = 'GUID', Mandatory=$True)]
        [Alias('ID')]
        $WorkItemGUID
    )
    PROCESS
    {
        TRY
        {
            If ($PSBoundParameters['WorkItemGUID'])
            {
                Write-Verbose -Message "[PROCESS] Retrieving WI with GUID"
                $ActivityObject = Get-SCSMObject -Id $WorkItemGUID -computername $scsmMGMTServer
            }
        
            #Retrieve Parent
            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name)"
            Write-Verbose -Message "[PROCESS] Retrieving WI Parent"
            $ParentRelatedObject = Get-SCSMRelationshipObject -ByTarget $ActivityObject -computername $scsmMGMTServer | ?{$_.RelationshipID -eq $wiContainsActivityRelClass.id.Guid}
            $ParentObject = $ParentRelatedObject.SourceObject

            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name) - Parent: $($ParentObject.Name)"

            If ($ParentObject.ClassName -eq 'System.WorkItem.ServiceRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ChangeRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ReleaseRecord' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Incident' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Problem')
            {
                Write-Verbose -Message "[PROCESS] This is the top level parent"
                
                #return parent object Work Item
                Return $ParentObject
            }
            Else
            {
                Write-Verbose -Message "[PROCESS] Not the top level parent. Running against this object"
                Get-SCSMWorkItemParent -WorkItemGUID $ParentObject.Id.GUID -computername $scsmMGMTServer
            }
        }
        CATCH
        {
            Write-Error -Message $Error[0].Exception.Message
        }
    }
}

#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/
function Create-UserInCMDB ($userEmail)
{
    #The ID for external users appears to be a GUID, but it can't be identified by get-scsmobject
    #The ID for internal domain users takes the form of domain_username_SMTP
    #It's unclear how this ID should be generated. Opted to take the form of an internal domain for the ID
    #By using the internal domain style (_SMTP) this means New/Update Work Item tasks will understand how to find these new external users going forward
    $username = $userEmail.Split("@")[0]
    $domainAndTLD = $userEmail.Split("@")[1]
    $domain = $domainAndTLD.Split(".")[0]
    $newID = $domain + "_" + $username + "_SMTP"

    #create the new user
    $newUser = New-SCSMObject -Class $domainUserClass -PropertyHashtable @{"domain" = "$domainAndTLD"; "username" = "$username"; "displayname" = "$userEmail"} -PassThru

    #create the user notification projection
    $userNoticeProjection = @{__CLASS = "$($domainUserClass.Name)";
                                __SEED = $newUser;
                                Notification = @{__CLASS = "$($notificationClass)";
                                                    __OBJECT = @{"ID" = $newID; "TargetAddress" = "$userEmail"; "DisplayName" = "E-mail address"; "ChannelName" = "SMTP"}
                                                }
                                }

    #create the user's email notification channel
    New-SCSMObjectProjection -Type "$($userHasPrefProjection.Name)" -Projection $userNoticeProjection

    return $newUser
}

#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/
function Add-IncidentComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the WI Object it passed to the function
    If ($WIObject.Id -ne $NULL) {

        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "AnalystComments"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "UserComments"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.ClassName)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Incident Object!"
    }
}

#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
function Add-ServiceRequestComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the SR Object it passed to the function
    If ($WIObject.Id -ne $NULL) {
         
 
        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "AnalystCommentLog"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "EndUserCommentLog"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.Classname)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Service Request Object!"
    }
}

#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
function Add-ProblemComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the SR Object it passed to the function
    If ($WIObject.Id -ne $NULL) {
         
 
        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "Comment"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "EndUserCommentLog"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.Classname)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Problem Object!"
    }
}

#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
function Add-ChangeRequestComment {
    param (
        [parameter(Mandatory=$True,Position=0)]$WIObject,
        [parameter(Mandatory=$True,Position=1)]$Comment,
        [parameter(Mandatory=$True,Position=2)]$EnteredBy,
        [parameter(Mandatory=$False,Position=3)]$AnalystComment,
        [parameter(Mandatory=$False,Position=4)]$IsPrivate
    )
 
    # Make sure that the SR Object it passed to the function
    If ($WIObject.Id -ne $NULL) {
         
 
        If ($AnalystComment -eq $true) {
            $CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"
            $CommentClassName = "AnalystComments"
        } else {
            $CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"
            $CommentClassName = "UserComments"
        }
 
        # Generate a new GUID for the comment
        $NewGUID = ([guid]::NewGuid()).ToString()
 
        # Create the object projection with properties
        $Projection = @{__CLASS = "$($WIObject.Classname)";
                        __SEED = $WIObject;
                        $CommentClassName = @{__CLASS = $CommentClass;
                                            __OBJECT = @{Id = $NewGUID;
                                                        DisplayName = $NewGUID;
                                                        Comment = $Comment;
                                                        EnteredBy = $EnteredBy;
                                                        EnteredDate = (Get-Date).ToUniversalTime();
                                                        IsPrivate = $IsPrivate;
                                            }
                        }
        }
 
        # Create the actual comment
        #NOTE: This Projection is 100% based on Cireson's CR projection as this is the ONLY projection
        #that features AssignedTo, AffectedUser, CreatedBy, and EndUser/Analyst Action Log comments
        #If you aren't a customer of Cireson, you'll need to create your own type projection
        #to use here.
        New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel" -Projection $Projection -computername $scsmMGMTServer
    } else {
        Throw "Invalid Change Request Object!"
    }
}

#search the Cireson KB based on content from a New Work Item and notify the Affected User
function Search-CiresonKnowledgeBase ($message, $workItem)
{
    $searchQuery = $workItem.Title.Trim() + " " + $workItem.Description.Trim()
    $resultsToNotify = @()

    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -UseDefaultCredentials -SessionVariable ecPortalSession
        $kbResults = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$ciresonAccountGUID&searchValue=$searchQuery&isManager=false&userLanguageCode=$ciresonKBLanguageCode") -WebSession $ecPortalSession
    }
    else
    {
        $portalLoginRequest = Invoke-WebRequest -Uri $ciresonPortalServer -Method get -SessionVariable ecPortalSession
        $loginForm = $portalLoginRequest.Forms[0]
        $loginForm.Fields["UserName"] = $ciresonPortalUsername
        $loginForm.Fields["Password"] = $ciresonPortalPassword
    
        $portalLoginPost = Invoke-WebRequest -Uri ($ciresonPortalServer + "/Login/Login?ReturnUrl=%2f") -WebSession $ecPortalSession -Method post -Body $loginForm.Fields
        $kbResults = Invoke-WebRequest -Uri ($ciresonPortalServer + "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$ciresonAccountGUID&searchValue=$searchQuery&isManager=false&userLanguageCode=$ciresonKBLanguageCode") -WebSession $ecPortalSession
    }

    $kbResults = $kbResults.Content | ConvertFrom-Json
    $kbResults =  $kbResults | ?{$_.endusercontent -ne ""} | select-object articleid, title
    
    if ($kbResults)
    {
        foreach ($kbResult in $kbResults)
        {
            $resultsToNotify += "<a href=$ciresonPortalServer" + "KnowledgeBase/View/$($kbResult.articleid)#/>$($kbResult.title)</a><br />"
        }

        #build the message and send through already established EWS connection
        #determine if IR/SR to build appropriate email response
        if ($workItem.ClassName -eq "System.WorkItem.Incident")
        {
            $resolveMailTo = "<a href=`"mailto:$email" + "?subject=" + "[" + $workItem.id + "]" + ";body=This%20can%20be%20[resolved]" + "`">resolve</a>"
        }
        else
        {
            $resolveMailTo = "<a href=`"mailto:$email" + "?subject=" + "[" + $workItem.id + "]" + ";body=This%20can%20be%20[cancelled]" + "`">cancel</a>"
        }

        $body = "We found some knowledge articles that may be of assistance to you <br/><br/>
            $resultsToNotify <br/><br />
            If one of these articles resolves your request, you can use the following
            link to $resolveMailTo your request."
    
        #add the Work Item ID, so a potential reply doesn't trigger the creation of a new work item but instead updates it
        Send-EmailFromWorkflowAccount -subject ("[" + $workItem.id + "]" - $message.subject) -body $body -bodyType "HTML" -toRecipients $message.From
    }
}

#send an email from the SCSM Workflow Account
function Send-EmailFromWorkflowAccount ($subject, $body, $bodyType, $toRecipients)
{
    $emailToSendOut = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchangeService
    $emailToSendOut.Subject = $subject
    $emailToSendOut.Body = $body
    $emailToSendOut.ToRecipients.Add($toRecipients)
    $emailToSendOut.Body.BodyType = $bodyType
    $emailToSendOut.Send()
}

function Schedule-WorkItem ($calAppt, $wiType, $workItem)
{
    #set the Scheduled Start/End dates on the Work Item
    $scheduledHashTable =  @{"ScheduledStartDate" = $calAppt.StartTime.ToUniversalTime(); "ScheduledEndDate" = $calAppt.EndTime.ToUniversalTime()}  
    switch ($wiType)
    {
        "ir" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "sr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "pr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "cr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "rr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}

        #activities
        "ma" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "pa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "sa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
        "da" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable}
    }

    #Trigger Update to update the Action log of the item
    Update-WorkItem -message $calAppt -wiType $wiType -workItemID $workItem.id
}

function Verify-WorkItem ($message)
{
    #If emails are being attached to New Work Items, filter on the File Attachment Description that equals the Exchange Conversation ID as defined in the Attach-EmailToWorkItem function
    if ($attachEmailToWorkItem -eq $true)
    {
        $emailAttachmentSearchObject = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" -ComputerName $scsmMGMTServer | select-object -first 1 
        $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $emailAttachmentSearchObject -ComputerName $scsmMGMTServer).sourceobject.id -ComputerName $scsmMGMTServer
        if ($emailAttachmentSearchObject -and $relatedWorkItemFromAttachmentSearch)
        {
            switch ($relatedWorkItemFromAttachmentSearch.ClassName)
            {
                "System.WorkItem.Incident" {Update-WorkItem -message $message -wiType "ir" -workItemID $relatedWorkItemFromAttachmentSearch.id}
                "System.WorkItem.ServiceRequest" {Update-WorkItem -message $message -wiType "sr" -workItemID $relatedWorkItemFromAttachmentSearch.id}
            }
        }
        else
        {
            #no match was found, Create a New Work Item
            New-WorkItem $message $defaultNewWorkItem
        }
    }
    else
    {
        #will never engage as Verify-WorkItem currently only works when attaching emails to work items 
    }
}
#endregion

#determine merge logic
if ($mergeReplies -eq $true)
{
    $attachEmailToWorkItem = $true
}

#define Exchange assembly and connect to EWS
[void] [Reflection.Assembly]::LoadFile("$exchangeEWSAPIPath")
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
switch ($exchangeAuthenticationType)
{
    "impersonation" {$exchangeService.Credentials = New-Object Net.NetworkCredential($username, $password, $domain)}
    "windows" {$exchangeService.UseDefaultCredentials = $true}
}
$exchangeService.AutodiscoverUrl($workflowEmailAddress)

#define search parameters and search on the defined classes
$inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
$inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)
$itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
$mimeContentSchema = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
$dateTimeItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
$now = get-date
$searchFilter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo -ArgumentList $dateTimeItem,$now

#build the Where-Object scriptblock based on defined configuration
$emailFilterString = '($_.ItemClass -eq "IPM.Note")'
$calendarFilterString = '($_.ItemClass -eq "IPM.Schedule.Meeting.Request")'
$unreadFilterString = '($_.isRead -eq $false)'
$inboxFilterString = $emailFilterString
if ($processCalendarAppointment -eq $true)
{
    $inboxFilterString = $emailFilterString + " -or " + $calendarFilterString
}

#finalize the where-object string by ensuring to look for all Unread Items
$inboxFilterString = "(" + $inboxFilterString + ")" + " -and " + $unreadFilterString
$inboxFilterString = [scriptblock]::Create("$inboxFilterString")

#filter the inbox
$inbox = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView) | where-object $inboxFilterString

#parse each message
foreach ($message in $inbox)
{
    #load the entire message
    $message.Load($propertySet)

    #Process an Email
    if ($message.ItemClass -eq "IPM.Note")
    {
        $email = New-Object System.Object 
        $email | Add-Member -type NoteProperty -name From -value $message.From.Address
        $email | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $email | Add-Member -type NoteProperty -name CC -value $message.CcRecipients
        $email | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $email | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $email | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $email | Add-Member -type NoteProperty -name ID -Value $message.ID
        $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic

        switch -Regex ($email.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
            #### activities ####
            "\[[R][A][0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

            #### 3rd party classes, work items, etc. add here ####

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "([R][E][:])(?!.*\[(([I|S|P|C][R])|([M|R][A]))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

            #### default action, create work item ####
            default {new-workitem $email $defaultNewWorkItem} 
        }

        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
        $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
    }

    #Process a Calendar Appointment
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Request")
    {
        $appointment = New-Object System.Object 
        $appointment | Add-Member -type NoteProperty -name StartTime -value $message.Start
        $appointment | Add-Member -type NoteProperty -name EndTime -value $message.End
        $appointment | Add-Member -type NoteProperty -name To -value $message.ToRecipients
        $appointment | Add-Member -type NoteProperty -name From -value $message.From.Address
        $appointment | Add-Member -type NoteProperty -name Attachments -value $message.Attachments
        $appointment | Add-Member -type NoteProperty -name Subject -value $message.Subject
        $appointment | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $appointment | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $appointment | Add-Member -type NoteProperty -name Body -value $message.Body.Text
        $appointment | Add-Member -type NoteProperty -name ID -Value $message.ID
        $appointment | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationID
        $appointment | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic

        switch -Regex ($appointment.subject) 
        { 
            #### primary work item types ####
            "\[[I][R][0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){schedule-workitem $appointment "ir" $result; $message.Accept($true)}}
            "\[[S][R][0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){schedule-workitem $appointment "sr" $result; $message.Accept($true)}}
            "\[[P][R][0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){schedule-workitem $appointment "pr" $result; $message.Accept($true)}}
            "\[[C][R][0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){schedule-workitem $appointment "cr" $result; $message.Accept($true)}}
            "\[[R][R][0-9]+\]" {$result = get-workitem $matches[0] $rrClass; if ($result){schedule-workitem $appointment "rr" $result; $message.Accept($true)}}

            #### activities ####
            "\[[M][A][0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){schedule-workitem $appointment "ma" $result; $message.Accept($true)}}
            "\[[P][A][0-9]+\]" {$result = get-workitem $matches[0] $paClass; if ($result){schedule-workitem $appointment "pa" $result; $message.Accept($true)}}
            "\[[S][A][0-9]+\]" {$result = get-workitem $matches[0] $saClass; if ($result){schedule-workitem $appointment "sa" $result; $message.Accept($true)}}
            "\[[D][A][0-9]+\]" {$result = get-workitem $matches[0] $daClass; if ($result){schedule-workitem $appointment "da" $result; $message.Accept($true)}}

            #### 3rd party classes, work items, etc. add here ####

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem $appointment $defaultNewWorkItem $true; schedule-workitem -calAppt $appointment -wiType $defaultNewWorkItem -workItem $returnedNewWorkItemToSchedule; $message.Accept($true)} 
        }
    }
}
