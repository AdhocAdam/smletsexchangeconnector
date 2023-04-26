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
Contributors: Martin Blomgren, Leigh Kilday, Tom Hendricks, nradler2, Justin Workman, Brad Zima, bennyguk, Jan Schulz, Peter Miklian, Daniel Polivka, Alexander Axberg
Reviewers: Tom Hendricks, Brian Wiest
Inspiration: The Cireson Community, Anders Asp, Stefan Roth, and (of course) Travis Wright for SMlets examples
Requires: PowerShell 4+, SMlets, and Exchange Web Services API (already installed on SCSM workflow server by virtue of stock Exchange Connector).
    3rd party option: If you're a Cireson customer and make use of their paid SCSM Portal with HTML Knowledge Base this will work as is
        if you aren't, you'll need to create your own Type Projection for Change Requests for the Add-ActionLogEntry
        function. Navigate to that function to read more. If you don't make use of their HTML KB, you'll want to keep $searchCiresonHTMLKB = $false
    Signed/Encrypted option: .NET 4.5 is required to use MimeKit.dll
Misc: The Release Record functionality does not exist in this as no out of box (or 3rd party) Type Projection exists to serve this purpose.
    You would have to create your own Type Projection in order to leverage this.
Version: 5.0.2 = #444 - Bug, Fix for Templates that contain one or many Activities
Version: 5.0.1 = #441 - Enhancement, Logging for when a user attempts to vote on a Review Activity that they are not a Reviewer on
                #440 - Enhancement, standardizing remaining functions param blocks and clarify logging event
                #439 - Enhancement, ItemClass logging event
                #438 - Documentation, Indented Review Activity code block in Update-WorkItem
                #437 - Enhancement, Additional logging for take and voting on behalf of AD group
                #436 - Documentation, Exchange notes
                #431 - Enhancement, address remaining PSScriptAnalyzer issues
                #429 - Bug, Get work item by conversation ID fails when multiple files exist
Version: 5.0.0 = #397 - Feature, Support for Multiple Keywords
                #399 - Enhancement, Use Singular Nouns
                #404 - Enhancement, Remove Unused Variables
                #405 - Enhancement, Refactor Get-TierMembership
                #406 - Enhancement, Refactor Set-SCSMTemplate and Update-SCSMPropertyCollection
                #407 - Enhancement, Variable usage for digitally signed messages
                #408 - Enhancement, Variable scope modifier for SCOM related actions
                #409 - Bug, Retrieving a user's email address for SCOM actions uses the wrong property
                #410 - Bug, Update-WorkItem does not individually attach files from an email
                #403 - Bug, Get-SCSMUserByEmailAddress fails to find the user if their email contains a PowerShell comparison operator
                #413 - Enhancement, The Description for the enum for IPM.Note.SMIME.MultipartSigned was missing the word "signed"
                #412 - Enhancement, Choosing a SCOM Group or Announcement Group no longer makes use of a label to display the selection
                #417 - Enhancement, Templates and Logging pane in the UI now use relevant iconography
                #422 - Bug, Certain environments may not load SMLets fast enough
                #423 - Enhancement, Logging is now more descriptive of the start, process, and finish stages of a run of the connector
                #424 - Enhancement, More logging and error handling for OAuth token authentication
                #425 - Enhancement, Set the DisplayName property on Work Items as though they were created in the SCSM Console
                #427 - Bug, Credentials supplied by any other means than using the SCSM Workflow engine when connecting to M365 fails email address validation
Version: 4.1.1 = #382 - Bug, When Azure Cloud Instance is defined in the Settings, the PowerShell returned a string instead of the full enum
                #378 - Bug and Enhancement, Several catch blocks did not contain logging events. Assigned User of a Closed Incident cannot correctly reactivate it which would result in a New and Related Work Item
                #379 - Enhancement, Get-CiresonPortalAPIToken function no longer makes use of ConvertTo-SecureString
Version: 4.1.0 = #355 - Enhancement, Support for defining Azure Cloud Instance (Exchange Online and/or Azure AI services)
                #366 - Enhancement, DLL form resizies as Settings UI resizes + link to wiki on configuring Custom Events
                #361 - Enhancement, Multi-mailbox, Custom Rules, History, and About forms stretch as Setting UI resizes
                #365 - Enhancement, Simplify New-SMEXCOEvent function
                #363 - Bug, Verify User availablity before Suggesting KA/RO
Version: 4.0.1 = #357 - Bug, When using Custom Rules, Work Items with a matched pattern don't correctly update
Version: 4.0.0 = #329 - Feature, Custom Rules
Version: 3.4.1 = #324 - Bug, File attachments allow illegal path characters
Version: 3.4.0 = #320 - Feature, File Attachments on Activities are pushed to the Parent Work Item
                #321 - Enhancement, Moved Cireson Search KB to a supported endpoint
                #318 - Enhancement, Searching Cireson Knowledge Base/Service Catalog could include a line break and produce erroneous results
                #325 - Enhancement, Get-CiresonSuggestionURL function performs more validation before invoking a search
                #327 - Bug, Null Reviewers on Review Activities cause the workflow to error when processing votes
                #315 - Bug, Emails with Subjects greater than 200 character cause the workflow to error
Version: 3.3.1 = #289 - Enhancement, Class naming precision
                #295 - Enhancement, Support Group functions have logging
                #294 - Bug, Users not found and not created in the CMDB when using Workflows can't be used in New-WorkItem or Update-WorkItem
                #291 - Enhancement, Dynamic Analyst Assignment related features have logging
                #293 - Enhancement, Adding Attachments features more error handling
                #290 - Bug, AML could predict an enum/CI that has been deleted since the last round of training
                #299 - Enhancement, Azure Cognitive Services related functions have logging
                #300 - Enhancement, New-SMExcoEvent calls are standardized to search them
                #302 - Enhancement, Set-AssignedToPerSupportGroup, Get-SCOMDistributedAppHealth, and Schedule-WorkItem functions have been made more efficient
                #303 - Bug, Incident/Problem Reactivation does not null neccesary values. Changing Incident Status based on replies no longer evaluates "Resolved" status
                #307 - Bug, Logging that validates the Run as Account email address for Exchange Online was used in on premise configurations
                #311 - Bug, A message that has no body would cause Add-ActionLogEntry to throw an error as the Comment would be null
Version: 3.3.0 = #90 - Bug, Cannot Convert null object to base 64 in Attach-FileToWorkItem
                #265 - Feature, Settings History
                #174 - Bug, Merge Reply Logic to handle subjects with multiple RE: patterns
Version: 3.2.0 = #233 - Feature, Mail can optionally be left in the inbox after processing
                #212 - Bug, MultiMailbox had issue processing due to an incorrect data type
                #255 - Bug, Attached emails use "eml" instead of ".eml" as their extension
                #234 - Bug, Depending on how the connector's PowerShell is opened it could render non UTF8 characters
                #263 - Documentation, typo in a contributor name
                #252 - Documentation, typo in Settings UI for Review Activities tooltip
Version: 3.1.0 = #10 - Optimization, Unable to process digitally signed AND encrypted messages
                #223 - Enhancement, Ability to ignore emails with invalid digital signature
                #224 - Feature, [pwsh] keyword for use with digital signatures
                #228 - Bug, Conversion overflow error when opening MP
                #217 - Bug, Regional decimal delimeter when validating Min File Attachment size
Version: 3.0.0 = #2 - Feature, adding support for logging regardless of deployment strategy
               = #207 - Feature, workflow support
Version: 2.4.0 = #171 - Optimization, Support for Exchange Online via OAuth 2.0 tokens
Version: 2.3.0 = #55 - Feature, Image Analysis (support for png, jpg, jpeg, bmp, and gif)
                #5 - Feature, Optical Character Recognition (support for png, jpg, jpeg, bmp, and gif)
                #54 - Feature, Speech to Text for Audio Files (support for wav and ogg)
                #188 - Bug, Dynamic work item assignment not functioning as expected
Version: 2.2.0 = #12 - Feature, Predict Affected/Impacted Configuration Item(s)
                #175 - Bug, Get-TemplatesByMailbox: incorrect variable name
                #88 - Bug, Verify-WorkItem does not handle null search results
Version: 2.1.0 = #34 - Feature, Integrate with Cireson Watchlist feature
                #169 - Bug, Corrected the logic issue for determining CiresonSuggestionURLs. Added class instantiation for URL suggestion email to include the BodyType object.
Version: 2.0.1 = #158 = Bug, Take Requires Group Membership fails on MA, CR, and PR in v2.x
Version: 2.0.0 = #49 - Enhancement, Introduce a Settings MP
Version: 1.6.1 = #142 = Bug, Schedule Outlook Meeting Task doesn't work in IE/Edge
                #44 = Enhancement, Refactor group membership check for take keyword
                #145 = Enhancement, Improve Manual Activity Actions
                #160 = Bug, KA/RO suggestions attempt to send even when disabled
Version: 1.6.0 = #135 = Bug, Dynamic Analyst Assignment not working
                #140 = Feature, Language Translation via Azure Translate
                #127 = Enhancement, Extended Support for AML returned values
                #130 = Bug, Add-ActionLogEntry has potential issue with similarly named Type Projections
                #132 = Documentation, includeWholeEmail notes incorrect
Version: 1.5.0 = #22 - Feature, Auto Assign Work Items when Created
                #112 - Feature, Predict Work Item Type, Classification and Support Group through Azure Machine Learning
                #116 - Bug, User reply that flips Incident status should not work against Closed Incidents
                #118 - Bug, Incidents that feature an Activity are missing their respective Activity Prefix (e.g. MA, RB)
                #120 - Feature, Record Scoring returned from Azure into Custom Class Extensions
                #92 - Optimization, Make it easier to configure how Related User's comments are marked on the Action Log
                #86 - Bug, Attach-FileToWorkItem Property 'Count' cannot be found
                #8 = Feature, Support for Custom Work Item prefixes as defined through SCSM
Version: 1.4.6 = #93 - Optimization, Configured Template Name may return multiple values
                #91 - Feature, Reply can change Incident Status
                #18 - Optimization, Convert "Invoke-WebRequest" to "Invoke-RestMethod" for Cireson Portal actions
                #97 - Optimization, Combine Get-SCSMAttachmentSettings and Get-SCSMObjectPrefix
                #104 - Bug, Dates that should be set by keywords are not
                #95 - Optimization, Convert Suggestion logic matrix to a Function
                #84 - Feature, Leverage keywords as an alternative to ACS Sentiment in order to drive New IR/SR creation
                #85 - Bug, Get-ScsmClassProperty error when using CR / MA support group enums
                #83 - Optimization, Process CC and BCC fields for multiple mailbox routing
                Milestone notes available at https://github.com/AdhocAdam/smletsexchangeconnector/milestone/3?closed=1
Version: 1.4.5 = #75 - Feature, Use Azure Cognitive Services Text Analytics API to set Urgency/Priority
                #68 - Feature, Redact sensitive information
                #69 - Optimization, Custom HTML Email Templates for Suggestions feature
                #71 - Optimization, Sort Suggestions back to Affected User sorts by words matched
Version: 1.4.4 = #48 - Created the ability to optionally set First Response Date on IR/SR when the connector makes Knowledge Article or Request Offering
                    suggestions to the Affected User
                #51 - Fixed issue with updating Work Items from Meeting Requests when the [Work Item] doesn't appear in the subject using an updated
                    version of the Verify-WorkItem function to perform the lookup
                #58 - Fixed issue with the connector leaving two comments when the Affected User is also the Assigned To user. Matching functionality
                    with the OOB Exchange Connector the comment is now marked as a Public End User Comment
                #56 - Introduce [take] keyword on Manual Activities
                #59 - Introduce Class Extension support for MA/CR Support Groups to support [take] enforcement
                #63 - #private is now a configurable keyword through the new $privateCommentKeyword
                #62 - Create a single Action Log function that allows multiple entry types (Actions and Comments)
                #60 - [resolved] keyword should trigger a Resolved Record Action Log entry instead of an Analyst Comment and allow
                    a default Resolution Category to be set. This was extended to include Service Requests, Problems, and their respective
                    Resolution Descriptions/Implementation Notes as well
                #61 - [reactivated] keyword should trigger a Record Reopened Action Log entry instead of an Analyst Comment
                #65 - Add minimum words to match to Knowledge Base suggestions
                #66 - Independently control Azure Cognitive Services in KA/RO Suggestion Feature
                #67 - Suggesting Request Offerings from the Cireson Portal returns no results
Version: 1.4.3 = Introduction of Azure Cognitive Services integration
Version: 1.4.2 = Fixed issue with attachment size comparison, when using SCSM size limits.
                 Fixed issue with [Take] function, if support group membership is checked.
                 Fixed issue with [approved] and [rejected] if groups or users belong to a different domain.
Version: 1.4.1 = Fixed issue raised on $irLowUrgency, was using the impact value instead of urgency.
Version: 1.4 = Changed how credentials are (optionally) added to SMLets, if provided, using $scsmMGMTparams hashtable and splatting.
                Created optional processing of mail from multiple mailboxes in addition to default mailbox. Messages must
                    be REDIRECTED (not forwarded) to the default mailbox using server or client rules, if this is enabled.
                Created optional per-mailbox IR, SR, PR, and CR template assignment, in support of multiple mailbox processing.
                Created optional configs for non-autodiscover connections to Exchange, and to provide explicit credentials.
                Created optional creation of new related ticket when comment is received on Closed ticket.
                Changed [take] behavior so that it (optionally) only functions if the email sender belongs to the currently selected support group.
                Created option to check SCSM attachment limits to determine if attachment(s) should be added or not.
                Created event handlers that trigger customizable functions in an accompanying "custom events" file. This allows proprietary
                        or custom functionality to be added at critical points throughout the script without having to merge them with this script.
                Created $voteOnBehalfOfGroups configuration variable so as to introduce the ability for users to Vote on Behalf of a Group
                Created Get-SCSMUserByEmailAddress to simplify how often users are retrieved by email address
                Changed areas that request a user object by email with the new Get-SCSMUserByEmailAddress function
                Added ability to create Problems and Change Requests as the default new work item
                Fixed issue when creating a New SR with activities, used identical approach for New CR functionality
Version: 1.3.3 = Fixed issue with [cancelled] keyword for Service Request
                 Added [take] keyword to Service Request
Version: 1.3.2 = Fixed issue when using the script other than on the SCSM Workflow server
                 Fixed issue when enabling/disabling features
Version: 1.3.1 = Fixed issue matching users when AD connector syncs users that were renamed.
                Changed how Request Offering suggestions are matched and made.
Version: 1.3 = created Set-CiresonPortalAnnouncement and Set-CoreSCSMAnnouncement to introduce announcement integration into the connector
                    by leveraging the new configurable [announcement] keyword and #low or #high tags to set priority on the announcement.
                    absence of the tag results in a normal priority announcement being created
                created Get-SCSMAuthorizedAnnouncer to verify the sender's permissions to post announcements
                created Get-CiresonPortalAnnouncements to search/update announcements
                created Read-MIMEMessage to allow parsing digitally signed or encryped emails. This feature leverages the
                    open source project known as MimeKit by Jeffrey Stedfast. It can be found here - https://github.com/jstedfast/MimeKit
                created Get-CiresonPortalUser to query for a user through the Cireson Web API to retrieve user information (object)
                created Get-CiresonPortalGroup to query for a group through the Cireson Web API to retrieve group information (object)
                created Search-AvailableCiresonPortalOfferings in order to look for relevant request offerings within a user's
                    Service Catalog scope to suggest relevant requests to the Affected User based on the content of their email
                improved/simplified Search-CiresonKnowledgeBase by use of new Get-CiresonPortalUser function
                created Get-SCOMDistributedAppHealth (SCOM integration) allows an authorized user to retrieve the health of
                    a distributed application from Operations Manager. Features configurable [keyword].
                created Get-SCOMAuthorizedRequester (SCOM integration) to verify that the individual requesting status on a SCOM Distributed Application
                    is authorized to do so
Version: 1.2 = created Send-EmailFromWorkflowAccount for future functions to leverage the SCSM workflow account defined therein
                updated Search-CiresonKnowledgeBase to use Send-EmailFromWorkflowAccount
                created $exchangeAuthenticationType so as to introduce Windows Authentication or Impersonation to bring to closer parity with stock EC connector
                expanded email processing loop to prepare for things other than IPM.Note message class (i.e. Calendar appointments, custom message classes per org.)
                created Schedule-WorkItem function to enable setting Scheduled Start/End Dates on Work Items based on the Calendar Start/End times.
                    introduced configuration variable for this feature ($processCalendarAppointment)
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

$startTime = Get-Date
#inspired and modified from Kevin Holman/Mark Manty https://kevinholman.com/2016/04/02/writing-events-with-parameters-using-powershell/
function New-SMEXCOEvent
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    [OutputType([String])]

    param (
        #The ID of the Event, e.g. 0, 1, 2, 3, etc.
        [parameter(Mandatory=$true, Position=0)]
        $EventID,
        #The message to write into the Event
        [parameter(Mandatory=$true, Position=1)]
        [string] $LogMessage,
        #The Event Log Source
        [parameter(Mandatory=$true, Position=2)]
        [ValidateSet("General", "CustomEvents", "Cryptography", "Test-EmailPattern", "New-WorkItem","Update-WorkItem","Add-EmailToSCSMObject", "Add-FileToSCSMObject", "Confirm-WorkItem",
            "Set-WorkItemScheduledTime", "Get-SCSMUserByEmailAddress", "Get-TierMembership", "Get-TierMember", "Get-AssignedToWorkItemVolume",
            "Set-AssignedToPerSupportGroup", "Get-SCSMWorkItemParent", "New-CMDBUser", "Add-ActionLogEntry", "Get-CiresonPortalAPIToken",
            "Get-CiresonPortalUser", "Get-CiresonPortalGroup", "Get-CiresonPortalAnnouncement", "Search-AvailableCiresonPortalOffering",
            "Search-CiresonKnowledgeBase", "Get-CiresonSuggestionURL", "Send-CiresonSuggestionEmail", "Add-CiresonWatchListUser",
            "Remove-CiresonWatchListUser", "Read-MIMEMessage", "Get-TemplatesByMailbox", "Get-SCSMAuthorizedAnnouncer", "Set-CoreSCSMAnnouncement",
            "Set-CiresonPortalAnnouncement", "Get-AzureEmailLanguage", "Get-SCOMAuthorizedRequester", "Get-SCOMDistributedAppHealth",
            "Send-EmailFromWorkflowAccount", "Test-KeywordsFoundInMessage", "Get-AMLWorkItemProbability", "Get-AzureEmailTranslation",
            "Get-AzureEmailKeyword", "Get-AzureEmailSentiment", "Get-AzureEmailImageAnalysis", "Get-AzureSpeechEmailAudioText",
            "Get-AzureEmailImageText", "Get-ACSWorkItemPriority")]
        [string] $Source,
        #The severity Level of the Event
        [parameter(Mandatory=$true, Position=3)]
        [ValidateSet("Information","Warning","Error")]
        [string] $Severity,
        #optional parameters to write into the Event
        [parameter(Mandatory=$false, Position=4)]
        [string] $EventParam1,
        [parameter(Mandatory=$false, Position=5)]
        [string] $EventParam2,
        [parameter(Mandatory=$false, Position=6)]
        [string] $EventParam3,
        [parameter(Mandatory=$false, Position=7)]
        [string] $EventParam4,
        [parameter(Mandatory=$false, Position=8)]
        [string] $EventParam5,
        [parameter(Mandatory=$false, Position=9)]
        [string] $EventParam6,
        [parameter(Mandatory=$false, Position=9)]
        [string] $EventParam7,
        [parameter(Mandatory=$false, Position=9)]
        [string] $EventParam8
    )

    if ($PSCmdlet.ShouldProcess($Source, "Log event $EventId"))
    {
        switch ($severity)
        {
            "Information" {$id = New-Object System.Diagnostics.EventInstance($eventID,1)}
            "Warning" {$id = New-Object System.Diagnostics.EventInstance($eventID,1,2)}
            "Error" {$id = New-Object System.Diagnostics.EventInstance($eventID,1,1)}
        }

        if ($loggingType -eq "Workflow")
        {
            try
            {
                #create the Event Log, if it already exists ignore and continue
                New-EventLog -LogName "SMLets Exchange Connector" -Source $Source -ErrorAction SilentlyContinue

                #Attempt to write to the Windows Event Log
                $evtObject = New-Object System.Diagnostics.EventLog
                $evtObject.Log = "SMLets Exchange Connector"
                $evtObject.Source = $source
                #$evtObject.Category = "custom"
                $evtObject.WriteEvent($id, @($LogMessage,$eventparam1,$eventparam2,$eventparam3,$eventparam4,$eventparam5,$eventparam6,$eventparam7,$eventparam8))
            }
            catch
            {
                #couldn't create a Windows Event Log entry
                Write-Error -Message "SMLets Exchange Connector Windows Event Log could not be created. $($_.Exception)"
            }
        }
        else
        {
            #The Event Log doesn't exist, use Write-Output/Warning/Error (SMA/Azure Automation)
            Write-Output "EventId:$EventID;Severity:$Severity;Source:$Source;:Message$LogMessage;"
            switch ($severity)
            {
                "Information" {Write-Output $EventID;$LogMessage;$Source;$Severity}
                "Warning" {Write-Warning $EventID;$LogMessage;$Source;$Severity}
                "Error" {Write-Error $EventID;$LogMessage;$Source;$Severity}
            }
        }
    }
}

#region #### Configuration and Prep SMLets ####
# Ensure SMLets is loaded in the current session.
if (-Not (Get-Module SMLets)) {
    Import-Module SMLets
    # If the import is unsuccessful and PowerShell 5+ is installed, pull SMLets from the gallery and install.
    if ($PsVersionTable.PsVersion.Major -ge 5 -And (-Not (Get-Module SMLets))) {
        Find-Module SMLets | Install-Module
        Import-Module SMLets
    }
}
#retrieve the SMLets Exchange Connector MP to define configuration
$smexcoSettingsMP = ((Get-SCSMObject -Class (Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings$")))
$smexcoSettingsMPMailboxes = ((Get-SCSMObject -Class (Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings.AdditionalMailbox$")))
$smexcoSettingsCustomRules = ((Get-SCSMObject -Class (Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings.CustomRule$")))
$scsmLFXConfigMP = Get-SCSMManagementPack -Id "50daaf82-06ce-cacb-8cf5-3950aebae0b0"

#define the SCSM management server, this could be a remote name or localhost
$scsmMGMTServer = "$($smexcoSettingsMP.SCSMmgmtServer)"
#if you are running this script in SMA or Orchestrator, you may need/want to present a credential object to the management server.  Leave empty, otherwise.
$scsmMGMTCreds = $null

#define/use SCSM WF credentials
#$exchangeAuthenticationType - "windows" or "impersonation" are valid inputs here only with a local Exchange server.
    #Windows will use the credentials that start this script in order to authenticate to Exchange and retrieve messages
        #choosing this option only requires the $workflowEmailAddress variable to be defined
        #this is ideal if you'll be using Task Scheduler or SMA to initiate this
    #Impersonation will use the credentials that are defined here to connect to Exchange and retrieve messages
        #choosing this option requires the $workflowEmailAddress, $username, $password, and $domain variables to be defined
#UseAutoDiscover = Determines whether ($true) or not ($false) to connect to Exchange using autodiscover.  If $false, provide a URL for $ExchangeEndpoint
    #ExchangeEndpoint = A URL in the format of 'https://<yourservername.domain.tld>/EWS/Exchange.asmx' such as 'https://mail.contoso.com/EWS/Exchange.asmx'
#UseExchangeOnline = When set to true the exchangeAuthenticationType is disregarded. Additionally on the General page in the Settings UI, the following should be set
    #Use AutoDiscover should be set to false
    #AutoDiscover URL should be set to https://outlook.office365.com/EWS/Exchange.asmx
$exchangeAuthenticationType = "windows"
$workflowEmailAddress = "$($smexcoSettingsMP.WorkflowEmailAddress)"
$username = ""
$password = ""
$domain = ""
$UseAutodiscover = $smexcoSettingsMP.UseAutoDiscover
$ExchangeEndpoint = "$($smexcoSettingsMP.ExchangeAutodiscoverURL)"
$UseExchangeOnline = $smexcoSettingsMP.UseExchangeOnline
$AzureClientID = "$($smexcoSettingsMP.AzureClientID)"
$AzureTenantID = "$($smexcoSettingsMP.AzureTenantID)"
$AzureCloudInstance = $($smexcoSettingsMP.AzureCloudInstance)
#determine which Azure Cloud (if any) is being used to set required URLs
switch ($AzureCloudInstance.Name)
{
    "SMLets.Exchange.Connector.AzureCloudInstanceEnum.AzurePublic"              {$azureScopeURL = "https://outlook.office.com/EWS.AccessAsUser.All"; $azureTokenURL = "https://login.microsoftonline.com/$AzureTenantID/oauth2/v2.0/token"; $azureTLD = "com"}
    "SMLets.Exchange.Connector.AzureCloudInstanceEnum.AzureUsGovernment"          {$azureScopeURL = "https://outlook.office.com/EWS.AccessAsUser.All"; $azureTokenURL = "https://login.microsoftonline.com/$AzureTenantID/oauth2/v2.0/token"; $azureTLD = "com"}
    "SMLets.Exchange.Connector.AzureCloudInstanceEnum.AzureUsGovernment.GCCHigh"  {$azureScopeURL = "https://outlook.office365.us/EWS.AccessAsUser.All"; $azureTokenURL = "https://login.microsoftonline.us/$AzureTenantID/oauth2/v2.0/token"; $azureTLD = "us"}
    "SMLets.Exchange.Connector.AzureCloudInstanceEnum.AzureUsGovernment.DOD"      {$azureScopeURL = "https://dod-outlook.office365.us/EWS.AccessAsUser.All"; $azureTokenURL = "https://login.microsoftonline.us/$AzureTenantID/oauth2/v2.0/token"; $azureTLD = "us"}
    default {$azureScopeURL = "https://outlook.office.com/EWS.AccessAsUser.All"; $azureTokenURL = "https://login.microsoftonline.com/$AzureTenantID/oauth2/v2.0/token"; $azureTLD = "com"}
}

#defaultNewWorkItem = set to either "ir", "sr", "pr", or "cr"
#default*RTemplate = define the displayname of the template you'll be using based on what you've set for $defaultNewWorkItem
#default(WORKITEM)ResolutionCategory = Optionally set the default Resolution Category for Incidents, Problems, or Service Requests when using the [resolved]
    # or [completed] keywords. Examples include:
    #$defaultIncidentResolutionCategory = "IncidentResolutionCategoryEnum.FixedByAnalyst$"
    #$defaultProblemResolutionCategory = "ProblemResolutionEnum.Documentation$"
    #$defaultServiceRequestImplementationCategory = "ServiceRequestImplementationResultsEnum.SuccessfullyImplemented$"
#checkAttachmentSettings = If $true, instructs the script to query SCSM for its attachment size and count limits per work item type.  If $false, neither is restricted.
#minFileSizeInKB = Set the minimum file size in kilobytes to be attached to work items
#createUsersNotInCMDB = If someone from outside your org emails into SCSM this allows you to take that email and create a User in your CMDB
#includeWholeEmail = If long chains get forwarded into SCSM, you can choose to write the whole email to a single action log entry OR the beginning to the first finding of "From:"
#attachEmailToWorkItem = If $true, attach email as an *.eml to each work item. Additionally, write the Exchange Conversation ID into the Description of the Attachment object
#voteOnBehalfOfGroups = If $true, Review Activities featuring an AD group can be Voted on Behalf of if one of the groups members Approve or Reject
#fromKeyword = If $includeWholeEmail is set to false, messages will be parsed UNTIL they find this word
#UseMailboxRedirection = Emails redirected to the mailbox this script connects to can have different templates applied, based on the address of another mailbox that redirects to it.
    # Mailboxes = This is a list of mailboxes that redirect to your primary workflow mailbox, and their properties. You DO NOT need to add the mailbox that you are connecting to.
        # DefaultWiType = Mail not associated with an existing ticket should be processed as a new work item of this type.
        # IRTemplate = The displaystring for the template that should be applied to new IR tickets
        # SRTemplate = The displaystring for the template that should be applied to new SR tickets
        # PRTemplate = The displaystring for the template that should be applied to new PR tickets
        # CRTemplate = The displaystring for the template that should be applied to new CR tickets
#CreateNewWorkItemWhenClosed = When set to $true, replies to a closed work item will create a new work item.
#takeRequiresGroupMembership = When set to $true, the [take] keyword only functions when the sender belongs to the ticket's support group.
    #This functionality requires the Cireson Analyst Portal for Service Manager.  Set $false if you do not use the Cireson Analyst Portal.
#crSupportGroupEnumGUID = Enter the GUID of the Enum for your custom CR Support Group to be leveraged with items such as takeRequiresGroupMembership. The best way to verify this is to
    #perform a: (Get-SCSMEnumeration -name "CR Support Group List name" | select-object id, name, displayname) to verify the ID of the Support Group property name for your Change Requests.
    #This is value needs to be set to the ID/GUID of your result.
    #This functionality requires the Cireson Analyst Portal for Service Manager.
#maSupportGroupEnumGUID = Enter the Name of the Enum for your custom MA Support Group to be leveraged with items such as takeRequiresGroupMembership. The best way to verify this is to
    #perform a: (Get-SCSMEnumeration -name "MA Support Group List name" | select-object id, name, displayname) to verify the ID of the Support Group property name for your Manual Activities. This is value needs to be
    #This is value needs to be set to the ID/GUID of your result.
    #This functionality requires the Cireson Analyst Portal for Service Manager.
#redactPiiFromMessage = if $true, instructs the script to remove personally identifiable information from message description in work-items.
    #The PII being redacted is listed as regex's in a text file located in the same directory as the script. The file is called "pii_regex.txt", has 1 regex per line, written without quotes.
#changeIncidentStatusOnReply = if $true, updates to Incidents will change their status based on who replied
#changeIncidentStatusOnReplyAffectedUser = If changeIncidentStatusOnReply is $true, the Status enum an Incident should change to when the Affected User updates the Incident via email
    #perform a: Get-SCSMChildEnumeration -Enumeration (Get-SCSMEnumeration -name "IncidentStatusEnum$") | Where-Object {$_.displayname -eq "myCustomStatusHere"}
    #to verify your Incident Status enum value Name if not using the out of box enums
#changeIncidentStatusOnReplyAssignedTo = If changeIncidentStatusOnReply is $true, The Status enum an Incident should change to when the Assigned To updates the Incident via email
    #perform a: Get-SCSMChildEnumeration -Enumeration (Get-SCSMEnumeration -name "IncidentStatusEnum$") | Where-Object {$_.displayname -eq "myCustomStatusHere"}
    #to verify your Incident Status enum value Name if not using the out of box enums
#changeIncidentStatusOnReplyRelatedUser = If changeIncidentStatusOnReply is $true, The Status enum an Incident should change to when a Related User updates the Incident via email
    #perform a: Get-SCSMChildEnumeration -Enumeration (Get-SCSMEnumeration -name "IncidentStatusEnum$") | Where-Object {$_.displayname -eq "myCustomStatusHere"}
    #to verify your Incident Status enum value Name if not using the out of box enums
#DynamicWorkItemAssignment = This functionality requires the Cireson Analyst Portal for Service Manager.
    #When variable is set to one of the following values, on New Work Item creation the Assigned To will be set based on:
    #"random" - Get all of the Analysts within a Support Group and randomly assign one of them to the New Work Item
    #"volume" - Get all of the Analysts within a Support Group, get the Analyst with the least amount of Assigned Work Items and assign them the New Work Item
    #"OOOrandom" - Same as above but doesn't assign when Out of Office using the SCSM Out of Office Management pack. https://github.com/AdhocAdam/scsmoutofoffice
    #"OOOvolume" - Same as above but doesn't assign when Out of Office using the SCSM Out of Office Management pack. https://github.com/AdhocAdam/scsmoutofoffice
#ExternalPartyCommentPrivacy*R = Control Comment Privacy when the User leaving the comment IS NOT the Affected User or Assigned To User
    #Comments can continue to be left as $null (stock connector behavior), always Private ($true), or always Public ($false). Please be mindful that this setting
    #can impact any custom Action Log notifiers you've configured and potentially expose/hide information from one party (Assigned To/Affected User).
#ExternalPartyCommentType*R = Control the type of comment that is left on Work Items when the User leaving the comment IS NOT the Affected User or Assigned To User
    #Comments can continue to be left as "AnalystComment" (stock connector behavior) or changed to "EndUserComment". Please be mindful modifying this setting in
    #conjuction with the above ExternalPartyCommentPrivacy*R. This can impact any custom Action Log notifiers you've configured and potentially expose/hide
    #information from one party (Assigned To/Affected User).
$defaultNewWorkItem = "$($smexcoSettingsMP.DefaultWorkItemType)"
$defaultIRTemplateGuid = "$($smexcoSettingsMP.DefaultIncidentTemplateGUID.Guid)"
$defaultSRTemplateGuid = "$($smexcoSettingsMP.DefaultServiceRequestTemplateGUID.Guid)"
$defaultPRTemplateGuid = "$($smexcoSettingsMP.DefaultProblemTemplateGUID.Guid)"
$defaultCRTemplateGuid = "$($smexcoSettingsMP.DefaultChangeRequestTemplateGUID.Guid)"
$defaultIncidentResolutionCategory = "$($smexcoSettingsMP.IncidentResolutionCategory.Name + "$")"
$defaultProblemResolutionCategory = "$($smexcoSettingsMP.ProblemResolutionCategory.Name + "$")"
$defaultServiceRequestImplementationCategory = "$($smexcoSettingsMP.ServiceRequestImplementationCategory.Name + "$")"
$checkAttachmentSettings = $smexcoSettingsMP.EnforceFileAttachmentSettings
$minFileSizeInKB = "$($smexcoSettingsMP.MinimumFileAttachmentSize)"
$createUsersNotInCMDB = $smexcoSettingsMP.CreateUsersNotInCMDB
$includeWholeEmail = $smexcoSettingsMP.IncludeWholeEmail
$attachEmailToWorkItem = $smexcoSettingsMP.AttachEmailToWorkItem
$deleteAfterProcessing = $smexcoSettingsMP.DeleteMessageAfterProcessing
$voteOnBehalfOfGroups = $smexcoSettingsMP.VoteOnBehalfOfADGroup
$fromKeyword = "$($smexcoSettingsMP.SCSMKeywordFrom)"
$UseMailboxRedirection = $smexcoSettingsMP.UseMailboxRedirection
if ($smexcoSettingsMPMailboxes) {$smexcoSettingsMPMailboxes | foreach-object {$Mailboxes += @{$_.MailboxAddress = @{"DefaultWiType"="$($_.MailboxTemplateWorkItemType)";"IRTemplate"="$($_.MailboxIRTemplateGUID)";"SRTemplate"="$($_.MailboxSRTemplateGUID)";"PRTemplate"="$($_.MailboxPRTemplateGUID)";"CRTemplate"="$($_.MailboxCRTemplateGUID)"}}}}
$CreateNewWorkItemWhenClosed = $smexcoSettingsMP.CreateNewWorkItemIfWorkItemClosed
$takeRequiresGroupMembership = $smexcoSettingsMP.TakeRequiresSupportGroupMembership
$crSupportGroupEnumGUID = "$($smexcoSettingsMP.CRSupportGroupGUID.Guid)"
$maSupportGroupEnumGUID = "$($smexcoSettingsMP.MASupportGroupGUID.Guid)"
$prSupportGroupEnumGUID = "$($smexcoSettingsMP.PRSupportGroupGUID.Guid)"
$redactPiiFromMessage = $smexcoSettingsMP.RemovePII
$changeIncidentStatusOnReply = $smexcoSettingsMP.ChangeIncidentStatusOnReply
$changeIncidentStatusOnReplyAffectedUser = "$($smexcoSettingsMP.IncidentStatusOnAffectedUserReply.Name + "$")"
$changeIncidentStatusOnReplyAssignedTo = "$($smexcoSettingsMP.IncidentStatusOnAssignedToReply.Name + "$")"
$changeIncidentStatusOnReplyRelatedUser = "$($smexcoSettingsMP.IncidentStatusOnRelatedUserReply.Name + "$")"
$DynamicWorkItemAssignment = $smexcoSettingsMP.DynamicAnalystAssignmentType
$ExternalPartyCommentPrivacyIR = Get-Variable -Name $smexcoSettingsMP.ExternalPartyCommentPrivacyIR -ValueOnly
$ExternalPartyCommentPrivacySR = Get-Variable -Name $smexcoSettingsMP.ExternalPartyCommentPrivacySR -ValueOnly
$ExternalPartyCommentTypeIR = "$($smexcoSettingsMP.ExternalPartyCommentTypeIR)"
$ExternalPartyCommentTypeSR = "$($smexcoSettingsMP.ExternalPartyCommentTypeSR)"

#processCalendarAppointment = If $true, scheduling appointments with the Workflow Inbox where a [WorkItemID] is in the Subject will
    #set the Scheduled Start and End Dates on the Work Item per the Start/End Times of the calendar appointment
    #and will also override $attachEmailToWorkItem to be $true if set to $false
#processDigitallySignedMessages = If $true, MimeKit will parse digitally signed email messages will also be processed in accordance with
    #settings defined for normal email processing
#processEncryptedMessage = If $true, MimeKit will parse encrypted email messages in accordance with settings defined for normal email processing.
    #In order for this to work, the correct decrypting certificate must be placed in either the Current User or Local Machine store
#certStore = If you will be processing encrypted email, you must define where the decrypting certificate is located. This takes the values
    #of either "user" or "machine"
#mergeReplies = If $true, emails that are Replies (signified by RE: in the subject) will attempt to be matched to a Work Item in SCSM by their
    #Exchange Conversation ID and will also override $attachEmailToWorkItem to be $true if set to $false
$processCalendarAppointment = $smexcoSettingsMP.ProcessCalendarAppointments
$processDigitallySignedMessages = $smexcoSettingsMP.ProcessDigitallySignedMessages
$processEncryptedMessages = $smexcoSettingsMP.ProcessDigitallyEncryptedMessages
$ignoreInvalidDigitalSignature = $smexcoSettingsMP.IgnoreInvalidDigitalSignature
$certStore = "$($smexcoSettingsMP.CertificateStore)"
$mergeReplies = $smexcoSettingsMP.MergeReplies

#optional, enable integration with Cireson Portal for Knowledge Base/Service Catalog suggestions or Watchlist integration
#this uses the now depricated Cireson KB API Search by Text, it works as of v7.x but should be noted it could be entirely removed in future portals
#enableCiresonIntegration = In order to use the Watchlist feature this must be flipped to true. Then the URL and keywords must be defined. This value
    #currently has no bearing on $searchAvailableCiresonPortalOfferings or $enableSetFirstResponseDateOnSuggestions values
#$numberOfWordsToMatchFromEmailToRO = defines the minimum number of words that must be matched from an email/new work item before Request Offerings will be
    #suggested to the Affected User about them
#$numberOfWordsToMatchFromEmailToKA = defines the minimum number of words that must be matched from an email/new work item before Knowledge Articles will be
    #suggested to the Affected User about them
#searchAvailableCiresonPortalOfferings = search available Request Offerings within the Affected User's permission scope based words matched in
    #their email/new work item
#enableSetFirstResponseDateOnSuggestions = When Knowledge Article or Request Offering suggestions are made to the Affected User, you can optionally
    #set the First Response Date value on a New Work Item
#$ciresonPortalServer = URL that will be used to search for KB articles, Request Offerings, and for the Watchlist via invoke-restmethod. Make sure to leave the "/" after your tld!
#$ciresonPortalWindowsAuth = how invoke-restmethod should attempt to authenticate to your portal server.
    #Leave true if your portal uses Windows Auth, change to False for Forms authentication.
    #If using forms, you'll need to set the ciresonPortalUsername and Password variables. For ease, you could set this equal to the username/password defined above
#add/removeWatchlistKeywords = the keywords to use to add/remove an Incident, Service Request, Problem, or Change to the Watchlist when updating a work item
$enableCiresonIntegration = $smexcoSettingsMP.EnableCiresonIntegration
$searchCiresonHTMLKB = $smexcoSettingsMP.CiresonSearchKnowledgeBase
$numberOfWordsToMatchFromEmailToRO = $smexcoSettingsMP.NumberOfWordsToMatchFromEmailToCiresonRequestOffering
$numberOfWordsToMatchFromEmailToKA = $smexcoSettingsMP.NumberOfWordsToMatchFromEmailToCiresonKnowledgeArticle
$searchAvailableCiresonPortalOfferings = $smexcoSettingsMP.CiresonSearchRequestOfferings
$enableSetFirstResponseDateOnSuggestions = $smexcoSettingsMP.EnableSetFirstResponseDateOnSuggestions
$ciresonPortalServer = "$($smexcoSettingsMP.CiresonPortalURL)"
$ciresonPortalWindowsAuth = $true
$ciresonPortalUsername = ""
$ciresonPortalPassword = ""
$addWatchlistKeywords = "$($smexcoSettingsMP.CiresonKeywordWatchlistAdd)" | Foreach-Object {$_.Split(",")}
$removeWatchlistKeywords = "$($smexcoSettingsMP.CiresonKeywordWatchlistRemove)" | Foreach-Object {$_.Split(",")}
$addWatchlistKeyword = "(" + [string]::Join('|', $addWatchlistKeywords) + ")"
$removeWatchlistKeyword = "(" + [string]::Join('|', $removeWatchlistKeywords) + ")"

#optional, enable Announcement control in SCSM/Cireson portal from email
#enableSCSMAnnouncements/enableCiresonPortalAnnouncements: You can create/update announcements
    #in Core SCSM or the Cireson Portal by changing these values from $false to $true
#announcementKeyword = if this [keyword] is in the message body, a new announcement will be created
#approved users, groups, type = control who is authorized to post announcements to SCSM/Cireson Portal
    #you can configure individual users by email address or use an Active Directory group
#priority keywords: These are the words that you can also include in the body of your message to further
    #define an announcment by setting it's priority.  For example the body of your message could be "Patching systems this weekend. [announcement] #low"
#priorityExpirationInHours: Since both SCSM and the Cireson require an announcement expiration date, when announcements are created
    #this is the number of hours added to the current time to set the announcement to expire. If you send Calendar Meetings which by definition
    #have a start and end time, these expirationInHours style variables are ignored
$enableSCSMAnnouncements = $smexcoSettingsMP.EnableSCSMAnnouncements
$enableCiresonPortalAnnouncements = $smexcoSettingsMP.EnableCiresonSCSMAnnouncements
$announcementKeywords = $smexcoSettingsMP.SCSMKeywordAnnouncement | Foreach-Object {$_.Split(",")}
$announcementKeyword = "(" + [string]::Join('|', $announcementKeywords) + ")"
$approvedADGroupForSCSMAnnouncements = "$($smexcoSettingsMP.SCSMApprovedAnnouncementGroupDisplayName)"
$approvedUsersForSCSMAnnouncements = "$($smexcoSettingsMP.SCSMApprovedAnnouncementUsers)"
$approvedMemberTypeForSCSMAnnouncer = "$($smexcoSettingsMP.SCSMAnnouncementApprovedMemberType)"
$lowAnnouncemnentPriorityKeywords = $smexcoSettingsMP.AnnouncementKeywordLow | Foreach-Object {$_.Split(",")}
$criticalAnnouncemnentPriorityKeywords = $smexcoSettingsMP.AnnouncementKeywordHigh | Foreach-Object {$_.Split(",")}
$lowAnnouncemnentPriorityKeyword = "(" + [string]::Join('|', $lowAnnouncemnentPriorityKeywords) + ")"
$criticalAnnouncemnentPriorityKeyword = "(" + [string]::Join('|', $criticalAnnouncemnentPriorityKeywords) + ")"
$lowAnnouncemnentExpirationInHours = $smexcoSettingsMP.AnnouncementPriorityLowExpirationInHours
$normalAnnouncemnentExpirationInHours = $smexcoSettingsMP.AnnouncementPriorityNormalExpirationInHours
$criticalAnnouncemnentExpirationInHours = $smexcoSettingsMP.AnnouncementPriorityCriticalExpirationInHours

<#ARTIFICIAL INTELLIGENCE OPTION 1, enable AI through Azure Cognitive Services
#PLEASE NOTE: HIGHLY EXPERIMENTAL!
By enabling this feature, the entire body of the email on New Work Item creation will be sent to the Azure
subscription provided and parsed by Cognitive Services. This information is collected and stored by Microsoft.
Use of this feature will vary between work items and isn't something that can be refined/configured.
#### SENTIMENT ANALYSIS ####
The information returned to this script is a percentage estimate of the perceived sentiment of the email in a
range from 0% to 100%. With 0% being negative and 100% being positive. Using this range, you can customize at
which percentage threshold an Incident or Service Request should be created. For example anything 95% or greater
creates an Service Request, while anything less than this creates an Incident. As such, this feature is
invoked only on New Work Item creation.
#### KEYWORD ANALYSIS ####
#It also returns what Azure Cognitive Services perceives to be keywords from the message. This is used to speed up searches
#against the Cireson Knowledge Base and/or Service Catalog for recommendations back to the Affected User.
#### !!!! WARNING !!!! ####
Use of this feature and in turn Azure Cognitive services has the possibility of incurring monthly Azure charges.
Please ensure you understand pricing model as seen at the followng URL
https://azure.microsoft.com/en-us/pricing/details/cognitive-services/text-analytics/
Using this URL, you can better plan for possible monetary charges and ensure you understand the potential financial
cost to your organization before enabling this feature.#>

#### requires Azure subscription and Cognitive Services Text Analytics API deployed ####
#enableAzureCognitiveServicesForKA = If enabled, Azure Cognitive Services Text Analytics API will extract keywords from the email to
    #search your Cireson Knowledge Base
#enableAzureCognitiveServicesForRO = If enabled, Azure Cognitive Services Text Analytics API will extract keywords from the email to
    #search your Cireson Service Catalog based on permissions scope of the Sender
#enableAzureCognitiveServicesForNewWI = If enabled, Azure Cognitive Services Text Analytics API will perform Sentiment Analysis
    #to either create an Incident or Service Request
#enableAzureCognitiveServicesPriorityScoring = If enabled, the Sentiment Score will be used
    #to set the Impact & Urgency and/or Urgency $ Priority on Incidents or Service Requests. Bounds can be edited within
    #the Get-ACSWorkItemPriority function. This feature can also be used even when using AI Option #3 described below.
#acsSentimentScore*RClassExtensionName = You can choose to write the returned Sentiment Score into the New Work Item.
    #This requires you to have extended the Incident AND Service Request classes with a custom Decimal value and then
    #enter the name of that property here.
#azureRegion = where Cognitive Services is deployed as seen in it's respective settings pane,
    #i.e. ukwest, eastus2, westus, northcentralus
#azureCogSvcTextAnalyticsAPIKey = API key for your cognitive services text analytics deployment. This is found in the settings pane for Cognitive Services in https://portal.azure.com
#minPercentToCreateServiceRequest = The minimum sentiment rating required to create a Service Request, a number less than this will create an Incident
$enableAzureCognitiveServicesForNewWI = $smexcoSettingsMP.EnableACSForNewWorkItem
$minPercentToCreateServiceRequest = "$($smexcoSettingsMP.MinACSSentimentToCreateSR)"
$enableAzureCognitiveServicesForKA = $smexcoSettingsMP.EnableACSForCiresonKASuggestion
$enableAzureCognitiveServicesForRO = $smexcoSettingsMP.EnableACSForCiresonROSuggestion
$enableAzureCognitiveServicesPriorityScoring = $smexcoSettingsMP.EnableACSPriorityScoring
$acsSentimentScoreIRClassExtensionName = "$($smexcoSettingsMP.ACSSentimentScoreIncidentClassExtensionGUID.Guid)"
$acsSentimentScoreSRClassExtensionName = "$($smexcoSettingsMP.ACSSentimentScoreServiceRequestClassExtensionGUID.Guid)"
$azureRegion = "$($smexcoSettingsMP.ACSTextAnalyticsRegion)"
$azureCogSvcTextAnalyticsAPIKey = "$($smexcoSettingsMP.ACSTextAnalyticsAPIKey)"

#ARTIFICIAL INTELLIGENCE OPTION 2, enable AI through pre-defined keywords
#If Azure Cognitive Services isn't an option for you can alternatively enable this more controlled mechanism
#that you configure with specific keywords in order to create either an Incident or Service Request when those keywords are present.
#For example, you could set the default Work Item type near the top of the configuration to be a Service Request but
#if any of these words are found, then an Incident would be created.
#enableKeywordMatchForNewWI = Indicates whether or not to use a list of keywords, which if found will force a different work item type to be used.
    #     NOTE: This will only function if Azure Cognitive Services is not also enabled.  ACS supersedes this functionality if enabled.
#workItemTypeOverrideKeywords = A regular expression containing keywords that will cause the new work item to be created as the $workItemOverrideType if found.
    #Use the pipe ("|") character to separate key words (it is the regex "OR")
    #you can test it out directly in PowerShell with the following 2 lines of PowerShell
    #     $workItemTypeOverrideKeywords = "(?<!in )error|problem|fail|crash|\bjam\b|\bjammed\b|\bjamming\b|broke|froze|issue|unable"
    #     "i have a problem with my computer" -match $workItemTypeOverrideKeywords
#workItemOverrideType = The type of work item to create if key words are found in the message.
$enableKeywordMatchForNewWI = "$($smexcoSettingsMP.EnableKeywordMatchForNewWorkItem)"
$workItemTypeOverrideKeywords = "$($smexcoSettingsMP.KeywordMatchRegexForNewWorkItem)"
$workItemOverrideType = "$($smexcoSettingsMP.KeywordMatchWorkItemType)"

#ARTIFICIAL INTELLIGENCE OPTION 3, enable AI through Azure Machine Learning
#PLEASE NOTE: HIGHLY EXPERIMENTAL!
#While using Azure Cognitive Services introduces some intelligence to the connector, Azure Machine Learning introduces
#a feedback loop that can ensure the connector applies increasing levels of intelligence based on your own unique SCSM environment.
#This is done by taking a subset of data from your SCSM DW, uploading to Azure Machine Learning, and then training ML on said dataset.
#Details on setup/configuration can be found on the SMLets Exchange Connector Wiki.
#Once trained, you can publish an AML web service the connector can consume in order to intelligently predict the
#Work Item Type (Incident/Service Request), Work Item Support Group, Work Item classification, and Impacted Configuration Items. Once enabled, you to set a minimum
#percent threshold before these values are applied. In doing so, you can ensure high standards for incoming email classification so
#AML only engages when met otherwise it will fallback to your Default Work Item template. Finally, AML can co-exist with the ACS
#feature that defines Priority/Urgency/Impact based on Sentiment Analysis.

#### requires Azure subscription and Azure Machine Learning web service deployed ####
#enableAzureMachineLearning = If enabled, your AML Web Service will attempt to define Work Item Type, Classification, and Support Group
#amlAPIKey = This is the API key for your AML web service
#amlURL = This is the URL for your AML web service
#amlWorkItemTypeMinPercentConfidence = The minimum percentage AML must return in order to decide should an Incident or Service Request be created
#amlWorkItemClassificationMinPercentConfidence = The minimum percentage AML must return in order to set the Classification on the New Work Item
#amlWorkItemSupportGroupMinPercentConfidence = The minimum percentage AML must return in order set the Support Group on the New Work Item
#amlImpactedConfigItemMinPercentConfidence = The minimum percentage AML must return in order set the Impacted Config Items on the New Work Item
#aml*ClassificationScoreClassExtensionName = Optionally write the returned percent confidence value to a decimal class extension on Incidents or Service Requests
#aml*ClassificationEnumPredictionExtName = Optionally write the returned enum value to an enum class extension bound to Classification/Area on Incidents or Service Requests
$enableAzureMachineLearning = $smexcoSettingsMP.EnableAML
$amlAPIKey = "$($smexcoSettingsMP.AMLAPIKey)"
$amlURL = "$($smexcoSettingsMP.AMLurl)"
#minimum confidence scores before AML engages
$amlWorkItemTypeMinPercentConfidence = "$($smexcoSettingsMP.AMLMinConfidenceWorkItemType)"
$amlWorkItemClassificationMinPercentConfidence = "$($smexcoSettingsMP.AMLMinConfidenceWorkItemClassification)"
$amlWorkItemSupportGroupMinPercentConfidence = "$($smexcoSettingsMP.AMLMinConfidenceWorkItemSupportGroup)"
$amlImpactedConfigItemMinPercentConfidence = "$($smexcoSettingsMP.AMLMinConfidenceImpactedConfigItem)"
#class extension, work item type prediction (str) and work item type prediction score (dec)
$amlWITypeIncidentStringClassExtensionName = "$($smexcoSettingsMP.AMLIRWorkItemTypePredictionClassExtensionGUID)"
$amlWITypeIncidentScoreClassExtensionName = "$($smexcoSettingsMP.AMLIncidentConfidenceClassExtensionGUID)"
$amlWITypeServiceRequestStringClassExtensionName = "$($smexcoSettingsMP.AMLSRWorkItemTypePredictionClassExtensionGUID)"
$amlWITypeServiceRequestScoreClassExtensionName = "$($smexcoSettingsMP.AMLServiceRequestConfidenceClassExtensionGUID)"
#class extension, incident classification score (dec) and classification prediction (enum)
$amlIncidentClassificationScoreClassExtensionName = "$($smexcoSettingsMP.AMLIncidentClassificationConfidenceClassExtensionGUID)"
$amlIncidentClassificationEnumPredictionExtName = "$($smexcoSettingsMP.AMLIncidentClassificationPredictionClassExtensionGUID)"
#class extension, incident tier queue score (dec) and tier queue prediction (enum)
$amlIncidentTierQueueScoreClassExtensionName = "$($smexcoSettingsMP.AMLIncidentSupportGroupConfidenceClassExtensionGUID)"
$amlIncidentTierQueueEnumPredictionExtName = "$($smexcoSettingsMP.AMLIncidentSupportGroupPredictionClassExtensionGUID)"
#class extension, service request area score (dec) and area prediction (enum)
$amlServiceRequestAreaScoreClassExtensionName = "$($smexcoSettingsMP.AMLServiceRequestClassificationConfidenceClassExtensionGUID)"
$amlServiceRequestAreaEnumPredictionExtName = "$($smexcoSettingsMP.AMLServiceRequestClassificationPredictionClassExtensionGUID)"
#class extension, service request support group score (dec) and support group prediction (enum)
$amlServiceRequestSupportGroupScoreClassExtensionName = "$($smexcoSettingsMP.AMLServiceRequestSupportGroupConfidenceClassExtensionGUID)"
$amlServiceRequestSupportGroupEnumPredictionExtName = "$($smexcoSettingsMP.AMLServiceRequestSupportGroupPredictionClassExtensionGUID)"

#optional, enable Language Translation through Azure Cognitive Services
#Use Translation services from Azure in order to create New Work Items that feature a translated Description as the First Comment in the
#Action Log. This Comment is a Public End User Comment that will be Entered By "Azure Translate/AUDISPLAYNAME" where AUDISPLAYNAME is
#the Affected User's Display Name
#defaultAzureTranslateLanguage = Pick the language code to which all New Work Items should be translated into. If the IR/SR is that language
#already, then significantly less will be consumed in Azure spend. A list of support languages and their codes can be found here
#https://docs.microsoft.com/en-us/azure/cognitive-services/translator/language-support
#pricing details can be found here: https://azure.microsoft.com/en-ca/pricing/details/cognitive-services/translator-text-api/
#azureCogSvcTranslateAPIKey = The API key for your deployed Azure Translation service
$enableAzureTranslateForNewWI = $smexcoSettingsMP.EnableACSTranslate
$defaultAzureTranslateLanguage = $smexcoSettingsMP.ACSTranslateDefaultLanguageCode
$azureCogSvcTranslateAPIKey = $smexcoSettingsMP.ACSTranslateAPIKey

#optional, enable Azure Vision through Azure Cognitive Services
#use Vision services from Azure in order to populate the Description of images attached to Work Items from email. By enabling, the image is first sent to
#the Image Analysis API to attempt to describe the top 5 categories or Tags of the image. In the event one of these Tags is the word "text"
#another call to the Optical Character Recognition (OCR) API will be made and attempt to extract text/words from the image.
#Given the maximum length of the File Attachment's Description property is 255 characters, Tags will always be present but the OCR result could be chopped off.
#For example a screenshot of an Outlook error message attached to an email would have these 5 Tags and the associated Description in the file's Description property.
#Tags:screenshot,abstract,text,design,graphic;Desc:Microsoft Outlook Cannot start Microsoft Outlook. Cannot open the Outlook window.
#pricing details can be found here: https://azure.microsoft.com/en-ca/pricing/details/cognitive-services/computer-vision/
$enableAzureVision = $smexcoSettingsMP.EnableACSVision
$azureVisionRegion = $smexcoSettingsMP.ACSVisionRegion
$azureCogSvcVisionAPIKey = $smexcoSettingsMP.ACSVisionAPIKey

#optional, enable Azure Speech through Azure Cognitive Services
#use Speech services from Azure in order to populate the Description of wav/ogg files attached to Work Items from email. By enabling, the audio file is first sent to
#use the Speech to Text API in an attempt to convert the file to readable text.
#pricing details can be found here: https://azure.microsoft.com/en-us/services/cognitive-services/speech-services/
$enableAzureSpeech = $smexcoSettingsMP.EnableACSSpeech
$azureSpeechRegion = $smexcoSettingsMP.ACSSpeechRegion
$azureCogSvcSpeechAPIKey = $smexcoSettingsMP.ACSSpeechAPIKey

#optional, enable SCOM functionality
#enableSCOMIntegration = set to $true or $false to enable this functionality
#scomMGMTServer = set equal to the name of your scom management server
#approvedMemberTypeForSCOM = To prevent unapproved individuals from gaining knowledge about your SCOM environment via SCSM, you must
    #choose to set either an AD Group that the sender must be a part of or you must manually define email addresses of users allowed to
    #make these requests. This variable can be set to "users" or "group"
#approvedADGroupForSCOM = if approvedUsersForSCOM = group, set this to the AD Group that contains groups/members that are allowed to make SCOM email requests
    #this approach allows you control access through Active Directory
#approvedUsersForSCOM = if approvedUsersForSCOM = users, set this to a comma seperated list of email addresses that are allowed to make SCOM email requests
    #this approach allows you to control through this script
#distributedApplicationHealthKeyword = the keyword to use in the subject for the connector to request DA status from SCOM
$enableSCOMIntegration = $smexcoSettingsMP.EnableSCOMIntegration
$scomMGMTServer = "$($smexcoSettingsMP.SCOMmgmtServer)"
$approvedMemberTypeForSCOM = "$($smexcoSettingsMP.SCOMApprovedMemberType)"
$approvedADGroupForSCOM = if ($smexcoSettingsMP.SCOMApprovedGroupGUID) {Get-SCSMObject -id ($smexcoSettingsMP.SCOMApprovedGroupGUID.Guid) | select-object username -ExpandProperty username}
$approvedUsersForSCOM = "$($smexcoSettingsMP.SCOMApprovedUsers)"
$distributedApplicationHealthKeywords = "$($smexcoSettingsMP.SCOMKeywordHealth)" | Foreach-Object {$_.Split(",")}
$distributedApplicationHealthKeyword = "(" + [string]::Join('|', $distributedApplicationHealthKeywords) + ")"

#retrieve SCSM Work Item keywords to be used
$acknowledgedKeywords = "$($smexcoSettingsMP.SCSMKeywordAcknowledge)" | Foreach-Object {$_.Split(",")}
$reactivateKeywords = "$($smexcoSettingsMP.SCSMKeywordReactivate)" | Foreach-Object {$_.Split(",")}
$resolvedKeywords = "$($smexcoSettingsMP.SCSMKeywordResolved)" | Foreach-Object {$_.Split(",")}
$closedKeywords = "$($smexcoSettingsMP.SCSMKeywordClosed)" | Foreach-Object {$_.Split(",")}
$holdKeywords = "$($smexcoSettingsMP.SCSMKeywordHold)" | Foreach-Object {$_.Split(",")}
$cancelledKeywords = "$($smexcoSettingsMP.SCSMKeywordCancelled)" | Foreach-Object {$_.Split(",")}
$takeKeywords = "$($smexcoSettingsMP.SCSMKeywordTake)" | Foreach-Object {$_.Split(",")}
$completedKeywords = "$($smexcoSettingsMP.SCSMKeywordCompleted)" | Foreach-Object {$_.Split(",")}
$skipKeywords = "$($smexcoSettingsMP.SCSMKeywordSkipped)" | Foreach-Object {$_.Split(",")}
$approvedKeywords = "$($smexcoSettingsMP.SCSMKeywordApprove)" | Foreach-Object {$_.Split(",")}
$rejectedKeywords = "$($smexcoSettingsMP.SCSMKeywordReject)" | Foreach-Object {$_.Split(",")}
$privateCommentKeywords = "$($smexcoSettingsMP.SCSMKeywordPrivate)" | Foreach-Object {$_.Split(",")}
$powershellKeywords = "$($smexcoSettingsMP.KeywordPowerShell)" | Foreach-Object {$_.Split(",")}

#format SCSM Work Item keywords to be used within the regular expressions throughout the connector
$acknowledgedKeyword = "(" + [string]::Join('|', $acknowledgedKeywords) + ")"
$reactivateKeyword = "(" + [string]::Join('|', $reactivateKeywords) + ")"
$resolvedKeyword = "(" + [string]::Join('|', $resolvedKeywords) + ")"
$closedKeyword = "(" + [string]::Join('|', $closedKeywords) + ")"
$holdKeyword = "(" + [string]::Join('|', $holdKeywords) + ")"
$cancelledKeyword = "(" + [string]::Join('|', $cancelledKeywords) + ")"
$takeKeyword = "(" + [string]::Join('|', $takeKeywords) + ")"
$completedKeyword = "(" + [string]::Join('|', $completedKeywords) + ")"
$skipKeyword = "(" + [string]::Join('|', $skipKeywords) + ")"
$approvedKeyword = "(" + [string]::Join('|', $approvedKeywords) + ")"
$rejectedKeyword = "(" + [string]::Join('|', $rejectedKeywords) + ")"
$privateCommentKeyword = "(" + [string]::Join('|', $privateCommentKeywords) + ")"
$powershellKeyword = "(" + [string]::Join('|', $powershellKeywords) + ")"

#define the path to the Exchange Web Services API and MimeKit
#the PII regex file and HTML Suggestion Template paths will only be leveraged if these features are enabled above.
#$htmlSuggestionTemplatePath must end with a "\"
$exchangeEWSAPIPath = "$($smexcoSettingsMP.FilePathEWSDLL)"
$mimeKitDLLPath = "$($smexcoSettingsMP.FilePathMimeKitDLL)"
$piiRegexPath = "$($smexcoSettingsMP.FilePathPIIRegex)"
$htmlSuggestionTemplatePath = "$($smexcoSettingsMP.FilePathHTMLSuggestionTemplates)"

#enable logging per standard Exchange Connector registry keys
#valid options on that registry key are 1 to 7 where 7 is the most verbose
#$loggingLevel = (Get-ItemProperty "HKLM:\Software\Microsoft\System Center Service Manager Exchange Connector" -ErrorAction SilentlyContinue).LoggingLevel
[int]$loggingLevel = "$($smexcoSettingsMP.LogLevel)"
$loggingType = "$($smexcoSettingsMP.LogType)"

#$ceScripts = invoke the Custom Events script, will optionally load custom/proprietary scripts as certain events occur.
    # set this equal to empty quotes ("") to turn custom events OFF
    # if using this feature, DO NOT USE QUOTES.  Start with a period/dot and then add the path to the script/runbook.
    # If running in SMA OR as a scheduled task with the custom events script in the same folder, use this format: . .\smletsExchangeConnector_CustomEvents.ps1
    # If running as a scheduled task and you have stored the events script in another folder, use this format: . C:\otherFolder\smletsExchangeConnector_CustomEvents.ps1'
$ceScripts = if($smexcoSettingsMP.FilePathCustomEvents.EndsWith(".ps1"))
{
    try
    {
        Invoke-Expression $smexcoSettingsMP.FilePathCustomEvents
        if ($loggingLevel -ge 4)
        {
            New-SMEXCOEvent -Source "CustomEvents" -EventID 0 -Severity "Information" -LogMessage "Custom Events PowerShell loaded successfully" | out-null
        }
        $true
    }
    catch
    {
        if ($loggingLevel -ge 2)
        {
            New-SMEXCOEvent -Source "CustomEvents" -EventID 1 -Severity "Warning" -LogMessage $_.Exception | out-null
        }
        $false
    }
}
#endregion #### Configuration ####

#region #### Process User Configs ####

# Configure SMLets with -ComputerName and -Credential switches, if applicable.
$scsmMGMTParams = @{ ComputerName = $scsmMGMTServer }
if ($scsmMGMTCreds) { $scsmMGMTParams.Credential = $scsmMGMTCreds }

# Configure AD Cmdlets with -Credential switch, if applicable.  Note that the -Server switch will be determined based on objects' domain.
if ($scsmMGMTCreds) {
    $adParams = @{Credential = $scsmMGMTCreds}
}
else {
    $adParams = @{}
}

# Set default templates and mailbox settings
if ($UseMailboxRedirection -eq $true) {
    $Mailboxes.add("$($workflowEmailAddress)", @{"DefaultWiType"=$defaultNewWorkItem;"IRTemplate"=$DefaultIRTemplateGuid;"SRTemplate"=$DefaultSRTemplateGuid;"PRTemplate"=$DefaultPRTemplateGuid;"CRTemplate"=$DefaultCRTemplateGuid})
}
else {
    $defaultIRTemplate = Get-SCSMObjectTemplate -Id $DefaultIRTemplateGuid @scsmMGMTParams
    $defaultSRTemplate = Get-SCSMObjectTemplate -Id $DefaultSRTemplateGuid @scsmMGMTParams
    $defaultPRTemplate = Get-SCSMObjectTemplate -Id $DefaultPRTemplateGuid @scsmMGMTParams
    $defaultCRTemplate = Get-SCSMObjectTemplate -Id $DefaultCRTemplateGuid @scsmMGMTParams
}
#endregion

#region #### SCSM Classes ####
$wiClass = get-scsmclass -name "System.WorkItem$" @scsmMGMTParams
$irClass = get-scsmclass -name "System.WorkItem.Incident$" @scsmMGMTParams
$srClass = get-scsmclass -name "System.WorkItem.ServiceRequest$" @scsmMGMTParams
$prClass = get-scsmclass -name "System.WorkItem.Problem$" @scsmMGMTParams
$crClass = get-scsmclass -name "System.Workitem.ChangeRequest$" @scsmMGMTParams
$rrClass = get-scsmclass -name "System.Workitem.ReleaseRecord$" @scsmMGMTParams
$maClass = get-scsmclass -name "System.WorkItem.Activity.ManualActivity$" @scsmMGMTParams
$raClass = get-scsmclass -name "System.WorkItem.Activity.ReviewActivity$" @scsmMGMTParams
$paClass = get-scsmclass -name "System.WorkItem.Activity.ParallelActivity$" @scsmMGMTParams
$saClass = get-scsmclass -name "System.WorkItem.Activity.SequentialActivity$" @scsmMGMTParams
$daClass = get-scsmclass -name "System.WorkItem.Activity.DependentActivity$" @scsmMGMTParams

$raHasReviewerRelClass = Get-SCSMRelationshipClass -name "System.ReviewActivityHasReviewer$" @scsmMGMTParams
$raReviewerIsUserRelClass = Get-SCSMRelationshipClass -name "System.ReviewerIsUser$" @scsmMGMTParams
$raVotedByUserRelClass = Get-SCSMRelationshipClass -name "System.ReviewerVotedByUser$" @scsmMGMTParams

#$userClass = get-scsmclass -name "System.User$" @scsmMGMTParams
$domainUserClass = get-scsmclass -name "System.Domain.User$" @scsmMGMTParams
$notificationClass = get-scsmclass -name "System.Notification.Endpoint$" @scsmMGMTParams

$irLowImpact = Get-SCSMEnumeration -name "System.WorkItem.TroubleTicket.ImpactEnum.Low$" @scsmMGMTParams
$irLowUrgency = Get-SCSMEnumeration -name "System.WorkItem.TroubleTicket.UrgencyEnum.Low$" @scsmMGMTParams
$irActiveStatus = Get-SCSMEnumeration -name "IncidentStatusEnum.Active$" @scsmMGMTParams

$affectedUserRelClass = get-scsmrelationshipclass -name "System.WorkItemAffectedUser$" @scsmMGMTParams
$assignedToUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItemAssignedToUser$" @scsmMGMTParams
$createdByUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItemCreatedByUser$" @scsmMGMTParams
$workResolvedByUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItem.TroubleTicketResolvedByUser$" @scsmMGMTParams
$wiAboutCIRelClass = Get-SCSMRelationshipClass -name "System.WorkItemAboutConfigItem$" @scsmMGMTParams
$wiRelatesToCIRelClass = Get-SCSMRelationshipClass -name "System.WorkItemRelatesToConfigItem$" @scsmMGMTParams
$wiRelatesToWIRelClass = Get-SCSMRelationshipClass -name "System.WorkItemRelatesToWorkItem$" @scsmMGMTParams
$wiContainsActivityRelClass = Get-SCSMRelationshipClass -name "System.WorkItemContainsActivity$" @scsmMGMTParams
#$sysUserHasPrefRelClass = Get-SCSMRelationshipClass -name "System.UserHasPreference$" @scsmMGMTParams

$fileAttachmentClass = Get-SCSMClass -Name "System.FileAttachment$" @scsmMGMTParams
$wiHasFileAttachRelClass = Get-SCSMRelationshipClass -name "System.WorkItemHasFileAttachment$" @scsmMGMTParams
$ciHasFileAttachRelClass = Get-SCSMRelationshipClass -name "System.ConfigItemHasFileAttachment$" @scsmMGMTParams
$fileAddedByUserRelClass = Get-SCSMRelationshipClass -name "System.FileAttachmentAddedByUser$" @scsmMGMTParams
$managementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup $scsmMGMTServer

$irTypeProjection = Get-SCSMTypeProjection -name "system.workitem.incident.projectiontype$" @scsmMGMTParams
$srTypeProjection = Get-SCSMTypeProjection -name "system.workitem.servicerequestprojection$" @scsmMGMTParams
$prTypeProjection = Get-SCSMTypeProjection -name "system.workitem.problem.projectiontype$" @scsmMGMTParams
$crTypeProjection = Get-SCSMTypeProjection -Name "system.workitem.changerequestprojection$" @scsmMGMTParams

$userHasPrefProjection = Get-SCSMTypeProjection -name "System.User.Preferences.Projection$" @scsmMGMTParams

# Retrieve Class Extensions on IR/SR/CR/MA if defined
if ($maSupportGroupEnumGUID)
{
    $maSupportGroupPropertyName = ($maClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$maSupportGroupEnumGUID*")}).Name
}
if ($crSupportGroupEnumGUID)
{
    $crSupportGroupPropertyName = ($crClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$crSupportGroupEnumGUID*")}).Name
}
if ($prSupportGroupEnumGUID)
{
    $prSupportGroupPropertyName = ($prClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$prSupportGroupEnumGUID*")}).Name
}
#azure cognitive services
if ($acsSentimentScoreIRClassExtensionName)
{
    $acsSentimentScoreIRClassExtensionName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$acsSentimentScoreIRClassExtensionName*")}).Name
}
if ($acsSentimentScoreSRClassExtensionName)
{
    $acsSentimentScoreSRClassExtensionName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$acsSentimentScoreSRClassExtensionName*")}).Name
}
#azure machine learning
#azure machine learning, work item type confidence % (IR/SR)
if ($amlWITypeIncidentScoreClassExtensionName)
{
    $amlWITypeIncidentScoreClassExtensionName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$amlWITypeIncidentScoreClassExtensionName*")}).Name
}
if ($amlWITypeServiceRequestScoreClassExtensionName)
{
    $amlWITypeServiceRequestScoreClassExtensionName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$amlWITypeServiceRequestScoreClassExtensionName*")}).Name
}
#azure machine learning, work item type prediction enum (IR/SR)
if ($amlWITypeIncidentStringClassExtensionName)
{
    $amlWITypeIncidentStringClassExtensionName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "String") -and ($_.Id -like "*$amlWITypeIncidentStringClassExtensionName*")}).Name
}
if ($amlWITypeServiceRequestStringClassExtensionName)
{
    $amlWITypeServiceRequestStringClassExtensionName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "String") -and ($_.Id -like "*$amlWITypeServiceRequestStringClassExtensionName*")}).Name
}
#azure machine learning, classification confidence % (Classification/Area)
if ($amlIncidentClassificationScoreClassExtensionName)
{
    $amlIncidentClassificationScoreClassExtensionName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$amlIncidentClassificationScoreClassExtensionName*")}).Name
}
if ($amlServiceRequestAreaScoreClassExtensionName)
{
    $amlServiceRequestAreaScoreClassExtensionName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$amlServiceRequestAreaScoreClassExtensionName*")}).Name
}
#azure machine learning, classification prediction enum (Classification/Area)
if ($amlIncidentClassificationEnumPredictionExtName)
{
    $amlIncidentClassificationEnumPredictionExtName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$amlIncidentClassificationEnumPredictionExtName*")}).Name
}
if ($amlServiceRequestAreaEnumPredictionExtName)
{
    $amlServiceRequestAreaEnumPredictionExtName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$amlServiceRequestAreaEnumPredictionExtName*")}).Name
}
#azure machine learning, support group confidence % (Tier Queue/Support Group)
if ($amlIncidentTierQueueScoreClassExtensionName)
{
    $amlIncidentTierQueueScoreClassExtensionName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$amlIncidentTierQueueScoreClassExtensionName*")}).Name
}
if ($amlServiceRequestSupportGroupScoreClassExtensionName)
{
    $amlServiceRequestSupportGroupScoreClassExtensionName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Decimal") -and ($_.Id -like "*$amlServiceRequestSupportGroupScoreClassExtensionName*")}).Name
}
#azure machine learning, support group prediction enum (Tier Queue/Support Group)
if ($amlIncidentTierQueueEnumPredictionExtName)
{
    $amlIncidentTierQueueEnumPredictionExtName = ($irClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$amlIncidentTierQueueEnumPredictionExtName*")}).Name
}
if ($amlServiceRequestSupportGroupEnumPredictionExtName)
{
    $amlServiceRequestSupportGroupEnumPredictionExtName = ($srClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.Id -like "*$amlServiceRequestSupportGroupEnumPredictionExtName*")}).Name
}
#endregion

#reply Regex
$replyRegex = '([R][E][:])|([A][W][:])|([S][V][:])|([A][n][t][w][:])|([V][S][:])|([R][E][F][:])|([R][I][F][:])|([S][V][:])|([B][L][S][:])|([A][t][b][\.][:])|([R][E][S][:])|([O][d][p][:])|([Y][N][T][:])|([A][T][B][:])'


#region #### Exchange Connector Functions ####
function New-WorkItem
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #Message from Exchange. Typically called from the core processing loop that has simlified the original Exchange object into a distilled PSCustomObject
        [parameter(Mandatory=$true)]
        [PSCustomObject]$message,
        #The type of Work Item to create by it's two letter shorthand
        [parameter(Mandatory=$true)]
        [ValidateSet("IR", "SR", "CR", "PR", "RR")]
        [string]$wiType,
        #Should the work item created by this function be returned back to the pipeline
        [parameter(Mandatory=$false)]
        [bool]$returnWIBool
    )

    if ($PSCmdlet.ShouldProcess("$message","New $wiType Work Item"))
    {
        $from = $message.From
        $to = $message.To
        $cced = $message.CC
        $title = $message.subject
        $description = $message.body

        if ($loggingLevel -ge 4)
        {
            $logMessage = "Creating $wiType
            From: $from
            CC Users: $($($cced.address) -join ',')
            Title: $title"
            New-SMEXCOEvent -Source "New-WorkItem" -EventId 0 -LogMessage $logMessage -Severity "Information"
        }

        #removes PII if RedactPiiFromMessage is enabled
        if ($redactPiiFromMessage -eq $true)
        {
            $description = remove-PII -body $description
        }

        #if the subject is longer than 200 characters take only the first 200.
        if ($title.length -ge "200")
        {
            $title = $title.substring(0,200)
        }

        #if the message is longer than 4000 characters take only the first 4000.
        if ($description.length -ge "4000")
        {
            $description = $description.substring(0,4000)
        }

        #find Affected User from the From Address
        $relatedUsers = @()
        $affectedUser = Get-SCSMUserByEmailAddress -EmailAddress "$from"
        if (($affectedUser)) {<#no change required#>}
        elseif ((!$affectedUser) -and ($createUsersNotInCMDB -eq $true)) {$affectedUser = New-CMDBUser "$from"}
        else {$affectedUser = New-CMDBUser -UserEmail $from -NoCommit}

        #find Related Users (To)
        if ($to.count -gt 0)
        {
            if ($to.count -eq 1)
            {
                $relatedUser = Get-SCSMUserByEmailAddress -EmailAddress "$($to.address)"
                if ($relatedUser)
                {
                    $relatedUsers += $relatedUser
                }
                else
                {
                    if ($createUsersNotInCMDB -eq $true)
                    {
                        $relatedUser = New-CMDBUser -UserEmail $to.address
                        $relatedUsers += $relatedUser
                    }
                }
            }
            else
            {
                $x = 0
                while ($x -lt $to.count)
                {
                    $ToSMTP = $to[$x]
                    $relatedUser = Get-SCSMUserByEmailAddress -EmailAddress "$($ToSMTP.address)"
                    if ($relatedUser)
                    {
                        $relatedUsers += $relatedUser
                    }
                    else
                    {
                        if ($createUsersNotInCMDB -eq $true)
                        {
                            $relatedUser = New-CMDBUser -UserEmail $ToSMTP.address
                            $relatedUsers += $relatedUser
                        }
                    }
                    $x++
                }
            }
        }

        #find Related Users (Cc)
        if ($cced.count -gt 0)
        {
            if ($cced.count -eq 1)
            {
                $relatedUser = Get-SCSMUserByEmailAddress -EmailAddress "$($cced.address)"
                if ($relatedUser)
                {
                    $relatedUsers += $relatedUser
                }
                else
                {
                    if ($createUsersNotInCMDB -eq $true)
                    {
                        $relatedUser = New-CMDBUser -UserEmail $cced.address
                        $relatedUsers += $relatedUser
                    }
                }
            }
            else
            {
                $x = 0
                while ($x -lt $cced.count)
                {
                    $ccSMTP = $cced[$x]
                    $relatedUser = Get-SCSMUserByEmailAddress -EmailAddress "$($ccSMTP.address)"
                    if ($relatedUser)
                    {
                        $relatedUsers += $relatedUser
                    }
                    else
                    {
                        if ($createUsersNotInCMDB -eq $true)
                        {
                            $relatedUser = New-CMDBUser -UserEmail $ccSMTP.address
                            $relatedUsers += $relatedUser
                        }
                    }
                    $x++
                }
            }
        }

        if ($loggingLevel -ge 4)
        {
            $logMessage = "User Relationships for $title
            Affected User: $($affectedUser.DisplayName)
            Related Users: $($($relatedUsers.DisplayName) -join ',')"
            New-SMEXCOEvent -Source "New-WorkItem" -EventId 1 -LogMessage $logMessage -Severity "Information"
        }

        if (($smexcoSettingsMP.UseMailboxRedirection -eq $true) -and ($smexcoSettingsMPMailboxes.Count -ge 1))
        {
            if ($loggingLevel -ge 4){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 0 -Severity "Information" -LogMessage "Mailbox redirection is being used. Attempting to identify Template to use."}
            $TemplatesForThisMessage = Get-TemplatesByMailbox -message $message
        }

        # Use the global default work item type or, if mailbox redirection is used, use the default work item type for the
        # specific mailbox that the current message was sent to. If Azure Cognitive Services is enabled
        # run the message through it to determine the Default Work Item type. Otherwise, use default if there is no match.
        if ($enableAzureCognitiveServicesForNewWI -eq $true)
        {
            $sentimentScore = Get-AzureEmailSentiment -messageToEvaluate $message.body

            #if the sentiment is greater than or equal to what is defined, create a Service Request.
            if ($sentimentScore -ge [int]$minPercentToCreateServiceRequest)
            {
                $wiType = "sr"
            }
            else #sentiment is lower than defined value, create an Incident.
            {
                $wiType = "ir"
            }
        }
        elseif ($enableKeywordMatchForNewWI -eq $true -and $(Test-KeywordsFoundInMessage -message $message) -eq $true) {
            #Keyword override is true and keyword(s) found in message
            $wiType = $workItemOverrideType
        }
        elseif ($UseMailboxRedirection -eq $true) {
            $wiType = if ($TemplatesForThisMessage) {$TemplatesForThisMessage["DefaultWiType"]} else {$defaultNewWorkItem}
        }
        elseif ($enableAzureMachineLearning -eq $true){
            $amlProbability = Get-AMLWorkItemProbability -EmailSubject $title -EmailBody $description
            $wiType = if ($amlProbability.WorkItemTypeConfidence -ge $amlWorkItemTypeMinPercentConfidence) {$amlProbability.WorkItemType} else {$defaultNewWorkItem}
        }
        elseif ($wiType){
            $wiType = $wiType
        }
        else {
            $wiType = $defaultNewWorkItem
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeCreateAnyWorkItem }

        #create the Work Item based on the globally defined Work Item type and Template
        switch ($wiType)
        {
            "ir" {
                        if ($UseMailboxRedirection -eq $true -And $TemplatesForThisMessage.Count -gt 0) {
                            $IRTemplate = Get-ScsmObjectTemplate -Id $($TemplatesForThisMessage["IRTemplate"]) @scsmMGMTParams
                        }
                        else {
                            $IRTemplate = $defaultIRTemplate
                        }
                        $newWorkItem = New-SCSMObject -Class $irClass -PropertyHashtable @{"ID" = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Incident")["Prefix"] + "{0}"; "Status" = $irActiveStatus; "Title" = $title; "Description" = $description; "Classification" = $null; "Impact" = $irLowImpact; "Urgency" = $irLowUrgency; "Source" = "IncidentSourceEnum.Email$"} -PassThru @scsmMGMTParams
                        $irProjection = Get-SCSMObjectProjection -ProjectionName $irTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                        if($message.Attachments){$message.Attachments | Foreach-Object {Add-FileToSCSMObject -attachment $_ -smobject $newWorkItem}}
                        if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $newWorkItem}
                        Set-SCSMTemplate -Projection $irProjection -Template $IRTemplate
                        #Set-SCSMObjectTemplate -Projection $irProjection -Template $IRTemplate @scsmMGMTParams
                        Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description; "Displayname" = "$($newWorkItem.Id) - $title"} @scsmMGMTParams
                        if ($affectedUser)
                        {
                            try {New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "New-WorkItem" -EventId 3 -LogMessage "The Created By User for $($newWorkItem.Name) could not be set." -Severity "Warning"}}
                            try {New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "New-WorkItem" -EventId 4 -LogMessage "The Affected User for $($newWorkItem.Name) could not be set." -Severity "Warning"}}
                        }
                        if ($relatedUsers)
                        {
                            foreach ($relatedUser in $relatedUsers)
                            {
                                New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk @scsmMGMTParams
                            }
                        }

                        #Translate the Description and enter as an Action Log entry if Azure Translate is used
                        #Since this function kicks off before Dynamic Analyst assignment, we'll prevent unnecesary notifications
                        #AdhocAdam Advanced Action Log Notifier won't engage on Comments Entered By -like "Azure Translate"
                        if ($enableAzureTranslateForNewWI -eq $true)
                        {
                            $descriptionSentence = $description.IndexOf(".")
                            $sampleDescription = $description.substring(0, $descriptionSentence)
                            if ($sampleDescription.Length -gt 1)
                            {
                                $detectedLanguage = Get-AzureEmailLanguage -TextToEvaluate $sampleDescription
                            }
                            else
                            {
                                $detectedLanguage = Get-AzureEmailLanguage -TextToEvaluate $description
                            }

                            if (($detectedLanguage.isTranslationSupported -eq $true) -and ($detectedLanguage.language -ne $defaultAzureTranslateLanguage))
                            {
                                $translatedDescription = Get-AzureEmailTranslation -TextToTranslate $description -SourceLanguage "$($detectedLanguage.language)" -TargetLanguage "$defaultAzureTranslateLanguage"
                                Add-ActionLogEntry -WIObject $newWorkItem -Action "EndUserComment" -Comment $translatedDescription -EnteredBy "Azure Translate/$($affectedUser.DisplayName)"
                            }

                            #if the detected language scores identical to the top alternative, use source language that isn't the default target language
                            $primaryAlternativeLang = $detectedLanguage.alternatives | Select-Object -first 1
                            if (($detectedLanguage.score -eq $primaryAlternativeLang.score) -and ($detectedLanguage.isTranslationSupported -eq $true) -and ($primaryAlternativeLang.isTranslationSupported -eq $true))
                            {
                                if (($detectedLanguage.language -eq $defaultAzureTranslateLanguage) -and ($primaryAlternativeLang.language -ne $defaultAzureTranslateLanguage))
                                {
                                    #translate with alternative
                                    $translatedDescription = Get-AzureEmailTranslation -TextToTranslate $description -SourceLanguage "$($primaryAlternativeLang.language)" -TargetLanguage "$defaultAzureTranslateLanguage"
                                    Add-ActionLogEntry -WIObject $newWorkItem -Action "EndUserComment" -Comment $translatedDescription -EnteredBy "Azure Translate/$($affectedUser.DisplayName)"
                                }
                            }
                        }

                        #Set Urgency/Impact from ACS Sentiment Analysis. If it was previously defined use it, otherwise make the ACS call
                        if (($enableAzureCognitiveServicesForNewWI -eq $true) -and ($enableAzureCognitiveServicesPriorityScoring -eq $true))
                        {
                            $priorityEnumArray = Get-ACSWorkItemPriority -score $sentimentScore -wiClass "System.WorkItem.Incident"
                            try {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Impact" = $priorityEnumArray[0]; "Urgency" = $priorityEnumArray[1]} @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventID 2 -Severity "Warning" -LogMessage $_.Exception}}
                        }
                        elseif (($enableAzureCognitiveServicesForNewWI -eq $false) -and ($enableAzureCognitiveServicesPriorityScoring -eq $true))
                        {
                            $sentimentScore = Get-AzureEmailSentiment -messageToEvaluate $newWorkItem.description
                            $priorityEnumArray = Get-ACSWorkItemPriority -score $sentimentScore -wiClass "System.WorkItem.Incident"
                            try {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Impact" = $priorityEnumArray[0]; "Urgency" = $priorityEnumArray[1]} @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventID 2 -Severity "Warning" -LogMessage $_.Exception}}
                        }

                        #write the sentiment score into the custom Work Item extension
                        if ($acsSentimentScoreIRClassExtensionName)
                        {
                            try {Set-SCSMObject -SMObject $newWorkItem -Property $acsSentimentScoreIRClassExtensionName -value $sentimentScore @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventID 3 -Severity "Warning" -LogMessage $_.Exception}}
                        }

                        #update the Support Group and Classification if Azure Machine Learning is being used
                        if ($enableAzureMachineLearning -eq $true)
                        {
                            #write confidence scores and enum predictions into Work Item
                            if ($amlWITypeIncidentStringClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlWITypeIncidentStringClassExtensionName" = $amlProbability.WorkItemType} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 2 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlWITypeIncidentScoreClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlWITypeIncidentScoreClassExtensionName" = $amlProbability.WorkItemTypeConfidence} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 3 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlIncidentClassificationScoreClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlIncidentClassificationScoreClassExtensionName" = $amlProbability.WorkItemClassificationConfidence} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 4 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlIncidentTierQueueScoreClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlIncidentTierQueueScoreClassExtensionName" = $amlProbability.WorkItemSupportGroupConfidence} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 5 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlIncidentTierQueueEnumPredictionExtName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlIncidentTierQueueEnumPredictionExtName" = $amlProbability.WorkItemSupportGroup} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 6 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlIncidentClassificationEnumPredictionExtName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlIncidentClassificationEnumPredictionExtName" = $amlProbability.WorkItemClassification} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 7 -Severity "Error" -LogMessage $_.Exception}}}

                            #when scores exceed thresholds, further define Work Item
                            if ($amlProbability.WorkItemSupportGroupConfidence -ge $amlWorkItemSupportGroupMinPercentConfidence)
                            {
                                try {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"TierQueue" = $amlProbability.WorkItemSupportGroup} @scsmMGMTParams}
                                catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 8 -Severity "Error" -LogMessage $_.Exception}}
                            }
                            if ($amlProbability.WorkItemClassificationConfidence -ge $amlWorkItemClassificationMinPercentConfidence)
                            {
                                try {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Classification" = $amlProbability.WorkItemClassification} @scsmMGMTParams}
                                catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 9 -Severity "Error" -LogMessage $_.Exception}}
                            }
                            if($amlProbability.AffectedConfigItemConfidence -ge $amlImpactedConfigItemMinPercentConfidence)
                            {
                                if($amlProbability.AffectedConfigItem.IndexOf(",") -gt 1)
                                {
                                    $amlProbability.AffectedConfigItem.Split(",") | ForEach-Object{try{New-SCSMRelationshipObject -Relationship $wiAboutCIRelClass -Source $newWorkItem -Target $_ -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 10 -Severity "Error" -LogMessage $_.Exception}}}
                                }
                                else
                                {
                                    try {New-SCSMRelationshipObject -Relationship $wiAboutCIRelClass -Source $newWorkItem -Target $amlProbability.AffectedConfigItem -Bulk @scsmMGMTParams}
                                    catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 10 -Severity "Error" -LogMessage $_.Exception}}
                                }
                            }
                        }

                        #Assign to an Analyst based on the Support Group that was set via Azure Machine Learning or just from the Template
                        if (($DynamicWorkItemAssignment) -and ($enableAzureMachineLearning -eq $true))
                        {
                            Set-AssignedToPerSupportGroup -SupportGroupID $amlProbability.WorkItemSupportGroup -WorkItem $newWorkItem
                        }
                        elseif ($DynamicWorkItemAssignment)
                        {
                            $templateSupportGroupID = $IRTemplate | select-object -expandproperty propertycollection | where-object{($_.path -like "*TierQueue*")} | select-object -ExpandProperty mixedvalue
                            if ($templateSupportGroupID)
                            {
                                Set-AssignedToPerSupportGroup -SupportGroupID $templateSupportGroupID -WorkItem $newWorkItem
                            }
                            else
                            {
                                if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Set-AssignedToPerSupportGroup" -EventID 3 -Severity "Warning" -LogMessage "The Assigned To User could not be set on $($newWorkItem.Name). Using Template: $($IRTemplate.DisplayName). This Template either does not have a Support Group defined that corresponds to a mapped Cireson Portal Group Mapping OR the Template being used/was copied from an OOB SCSM Template"}
                            }
                        }

                        #### Determine auto-response logic for Knowledge Base and/or Request Offering Search. Verify User exists in SCSM (IsNew = $false) vs. created in memory for this run (IsNew = $true) ####
                        if ($affectedUser.IsNew -eq $false) {$ciresonSuggestionURLs = Get-CiresonSuggestionURL -SuggestKA:$searchCiresonHTMLKB -AzureKA:$enableAzureCognitiveServicesForKA -SuggestRO:$searchAvailableCiresonPortalOfferings -AzureRO:$enableAzureCognitiveServicesForRO -WorkItem $newWorkItem -AffectedUser $affectedUser}
                        if ($null -ne $ciresonSuggestionURLs)
                        {
                            if ($ciresonSuggestionURLs[0] -and $ciresonSuggestionURLs[1])
                            {
                                Send-CiresonSuggestionEmail -KnowledgeBaseURLs $ciresonSuggestionURLs[0] -RequestOfferingURLs $ciresonSuggestionURLs[1] -Workitem $newWorkItem -AffectedUserEmailAddress $from
                            }
                            elseif ($ciresonSuggestionURLs[1])
                            {
                                Send-CiresonSuggestionEmail -RequestOfferingURLs $ciresonSuggestionURLs[1] -Workitem $newWorkItem -AffectedUserEmailAddress $from
                            }
                            elseif ($ciresonSuggestionURLs[0])
                            {
                                Send-CiresonSuggestionEmail -KnowledgeBaseURLs $ciresonSuggestionURLs[0] -Workitem $newWorkItem -AffectedUserEmailAddress $from
                            }
                        }

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterCreateIR }

                    }
            "sr" {
                        if ($UseMailboxRedirection -eq $true -and $TemplatesForThisMessage.Count -gt 0) {
                            $SRTemplate = Get-ScsmObjectTemplate -Id $($TemplatesForThisMessage["SRTemplate"]) @scsmMGMTParams
                        }
                        else {
                            $SRTemplate = $defaultSRTemplate
                        }
                        $newWorkItem = new-scsmobject -class $srClass -propertyhashtable @{"ID" = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.ServiceRequest")["Prefix"] + "{0}"; "Title" = $title; "Description" = $description; "Status" = "ServiceRequestStatusEnum.New$"} -PassThru @scsmMGMTParams
                        $srProjection = Get-SCSMObjectProjection -ProjectionName $srTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                        if($message.Attachments){$message.Attachments | Foreach-Object {Add-FileToSCSMObject -attachment $_ -smobject $newWorkItem}}
                        if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $newWorkItem}
                        Set-SCSMTemplate -Projection $srProjection -Template $SRTemplate
                        #Set-SCSMObjectTemplate -projection $srProjection -Template $SRTemplate @scsmMGMTParams
                        Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description; "Displayname" = "$($newWorkItem.Id) : $title"} @scsmMGMTParams
                        if ($affectedUser)
                        {
                            try {New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "New-WorkItem" -EventId 3 -LogMessage "The Created By User for $($newWorkItem.Name) could not be set." -Severity "Warning"}}
                            try {New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "New-WorkItem" -EventId 4 -LogMessage "The Affected User for $($newWorkItem.Name) could not be set." -Severity "Warning"}}
                        }
                        if ($relatedUsers)
                        {
                            foreach ($relatedUser in $relatedUsers)
                            {
                                New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk @scsmMGMTParams
                            }
                        }

                        #Translate the Description and enter as an Action Log entry if Azure Translate is used
                        #Since this function kicks off before Dynamic Analyst assignment, we'll prevent unnecesary notifications
                        #AdhocAdam Advanced Action Log Notifier won't engage on Comments Entered By -like "Azure Translate"
                        if ($enableAzureTranslateForNewWI -eq $true)
                        {
                            $descriptionSentence = $description.IndexOf(".")
                            $sampleDescription = $description.substring(0, $descriptionSentence)
                            if ($sampleDescription.Length -gt 1)
                            {
                                $detectedLanguage = Get-AzureEmailLanguage -TextToEvaluate $sampleDescription
                            }
                            else
                            {
                                $detectedLanguage = Get-AzureEmailLanguage -TextToEvaluate $description
                            }

                            if (($detectedLanguage.isTranslationSupported -eq $true) -and ($detectedLanguage.language -ne $defaultAzureTranslateLanguage))
                            {
                                $translatedDescription = Get-AzureEmailTranslation -TextToTranslate $description -SourceLanguage "$($detectedLanguage.language)" -TargetLanguage "$defaultAzureTranslateLanguage"
                                Add-ActionLogEntry -WIObject $newWorkItem -Action "EndUserComment" -Comment $translatedDescription -EnteredBy "Azure Translate/$($affectedUser.DisplayName)"
                            }

                            #if the detected language scores identical to the top alternative, use source language that isn't the default target language
                            $primaryAlternativeLang = $detectedLanguage.alternatives | Select-Object -first 1
                            if (($detectedLanguage.score -eq $primaryAlternativeLang.score) -and ($detectedLanguage.isTranslationSupported -eq $true) -and ($primaryAlternativeLang.isTranslationSupported -eq $true))
                            {
                                if (($detectedLanguage.language -eq $defaultAzureTranslateLanguage) -and ($primaryAlternativeLang.language -ne $defaultAzureTranslateLanguage))
                                {
                                    #translate with alternative
                                    $translatedDescription = Get-AzureEmailTranslation -TextToTranslate $description -SourceLanguage "$($primaryAlternativeLang.language)" -TargetLanguage "$defaultAzureTranslateLanguage"
                                    Add-ActionLogEntry -WIObject $newWorkItem -Action "EndUserComment" -Comment $translatedDescription -EnteredBy "Azure Translate/$($affectedUser.DisplayName)"
                                }
                            }
                        }

                        #Set Urgency/Priority from ACS Sentiment Analysis. If it was previously defined use it, otherwise make the ACS call
                        if (($enableAzureCognitiveServicesForNewWI -eq $true) -and ($enableAzureCognitiveServicesPriorityScoring -eq $true))
                        {
                            $priorityEnumArray = Get-ACSWorkItemPriority -score $sentimentScore -wiClass "System.WorkItem.ServiceRequest"
                            try {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Urgency" = $priorityEnumArray[0]; "Priority" = $priorityEnumArray[1]} @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventID 2 -Severity "Warning" -LogMessage $_.Exception}}
                        }
                        elseif (($enableAzureCognitiveServicesForNewWI -eq $false) -and ($enableAzureCognitiveServicesPriorityScoring -eq $true))
                        {
                            $sentimentScore = Get-AzureEmailSentiment -messageToEvaluate $newWorkItem.description
                            $priorityEnumArray = Get-ACSWorkItemPriority -score $sentimentScore -wiClass "System.WorkItem.ServiceRequest"
                            try {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Urgency" = $priorityEnumArray[0]; "Priority" = $priorityEnumArray[1]} @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventID 2 -Severity "Warning" -LogMessage $_.Exception}}
                        }

                        #write the sentiment score into the custom Work Item extension
                        if ($acsSentimentScoreSRClassExtensionName)
                        {
                            try {Set-SCSMObject -SMObject $newWorkItem -Property $acsSentimentScoreSRClassExtensionName -value $sentimentScore @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventID 3 -Severity "Warning" -LogMessage $_.Exception}}
                        }

                        #update the Support Group and Classification if Azure Machine Learning is being used
                        if ($enableAzureMachineLearning -eq $true)
                        {
                            #write confidence scores into Work Item
                            if ($amlWITypeServiceRequestStringClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlWITypeServiceRequestStringClassExtensionName" = $amlProbability.WorkItemType} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 2 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlWITypeServiceRequestScoreClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlWITypeServiceRequestScoreClassExtensionName" = $amlProbability.WorkItemTypeConfidence} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 3 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlServiceRequestAreaScoreClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlServiceRequestAreaScoreClassExtensionName" = $amlProbability.WorkItemClassificationConfidence} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 4 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlServiceRequestSupportGroupScoreClassExtensionName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlServiceRequestSupportGroupScoreClassExtensionName" = $amlProbability.WorkItemSupportGroupConfidence} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 5 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlServiceRequestSupportGroupEnumPredictionExtName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlServiceRequestSupportGroupEnumPredictionExtName" = $amlProbability.WorkItemSupportGroup} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 6 -Severity "Error" -LogMessage $_.Exception}}}
                            if ($amlServiceRequestAreaEnumPredictionExtName) {try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"$amlServiceRequestAreaEnumPredictionExtName" = $amlProbability.WorkItemClassification} @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 7 -Severity "Error" -LogMessage $_.Exception}}}


                            #when scores exceed thresholds, further define Work Item
                            if ($amlProbability.WorkItemSupportGroupConfidence -ge $amlWorkItemSupportGroupMinPercentConfidence)
                            {
                                try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"SupportGroup" = $amlProbability.WorkItemSupportGroup} @scsmMGMTParams}
                                catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 8 -Severity "Error" -LogMessage $_.Exception}}
                            }
                            if ($amlProbability.WorkItemClassificationConfidence -ge $amlWorkItemClassificationMinPercentConfidence)
                            {
                                try{Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Classification" = $amlProbability.WorkItemClassification} @scsmMGMTParams}
                                catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 9 -Severity "Error" -LogMessage $_.Exception}}
                            }
                            if($amlProbability.AffectedConfigItemConfidence -ge $amlImpactedConfigItemMinPercentConfidence)
                            {
                                if($amlProbability.AffectedConfigItem.IndexOf(",") -gt 1)
                                {
                                    $amlProbability.AffectedConfigItem.Split(",") | ForEach-Object{try{New-SCSMRelationshipObject -Relationship $wiAboutCIRelClass -Source $newWorkItem -Target $_ -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 10 -Severity "Error" -LogMessage $_.Exception}}}
                                }
                                else
                                {
                                    try{New-SCSMRelationshipObject -Relationship $wiAboutCIRelClass -Source $newWorkItem -Target $amlProbability.AffectedConfigItem -Bulk @scsmMGMTParams}
                                    catch {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 10 -Severity "Error" -LogMessage $_.Exception}}
                                }
                            }
                        }

                        #Assign to an Analyst based on the Support Group that was set via Azure Machine Learning or just from the Template
                        if (($DynamicWorkItemAssignment) -and ($enableAzureMachineLearning -eq $true))
                        {
                            Set-AssignedToPerSupportGroup -SupportGroupID $amlProbability.WorkItemSupportGroup -WorkItem $newWorkItem
                        }
                        elseif ($DynamicWorkItemAssignment)
                        {
                            $templateSupportGroupID = $SRTemplate | select-object -expandproperty propertycollection | where-object{($_.path -like "*SupportGroup*")} | select-object -ExpandProperty mixedvalue
                            if ($templateSupportGroupID)
                            {
                                Set-AssignedToPerSupportGroup -SupportGroupID $templateSupportGroupID -WorkItem $newWorkItem
                            }
                            else
                            {
                                if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Set-AssignedToPerSupportGroup" -EventID 3 -Severity "Warning" -LogMessage "The Assigned To User could not be set on $($newWorkItem.Name). Using Template: $($SRTemplate.DisplayName). This Template either does not have a Support Group defined that corresponds to a mapped Cireson Portal Group Mapping OR the Template being used/was copied from an OOB SCSM Template"}
                            }
                        }

                        #### Determine auto-response logic for Knowledge Base and/or Request Offering Search. Verify User exists in SCSM (IsNew = $false) vs. created in memory for this run (IsNew = $true) ####
                        if ($affectedUser.IsNew -eq $false) {$ciresonSuggestionURLs = Get-CiresonSuggestionURL -SuggestKA:$searchCiresonHTMLKB -AzureKA:$enableAzureCognitiveServicesForKA -SuggestRO:$searchAvailableCiresonPortalOfferings -AzureRO:$enableAzureCognitiveServicesForRO -WorkItem $newWorkItem -AffectedUser $affectedUser}
                        if ($null -ne $ciresonSuggestionURLs)
                        {
                            if ($ciresonSuggestionURLs[0] -and $ciresonSuggestionURLs[1])
                            {
                                Send-CiresonSuggestionEmail -KnowledgeBaseURLs $ciresonSuggestionURLs[0] -RequestOfferingURLs $ciresonSuggestionURLs[1] -Workitem $newWorkItem -AffectedUserEmailAddress $from
                            }
                            elseif ($ciresonSuggestionURLs[0])
                            {
                                Send-CiresonSuggestionEmail -RequestOfferingURLs $ciresonSuggestionURLs[1] -Workitem $newWorkItem -AffectedUserEmailAddress $from
                            }
                            elseif ($ciresonSuggestionURLs[1])
                            {
                                Send-CiresonSuggestionEmail -KnowledgeBaseURLs $ciresonSuggestionURLs[0] -Workitem $newWorkItem -AffectedUserEmailAddress $from
                            }
                        }

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterCreateSR }
                    }
            "pr" {
                        if ($UseMailboxRedirection -eq $true -and $TemplatesForThisMessage.Count -gt 0) {
                            $PRTemplate = Get-ScsmObjectTemplate -Id $($TemplatesForThisMessage["PRTemplate"]) @scsmMGMTParams
                        }
                        else {
                            $PRTemplate = $defaultPRTemplate
                        }
                        $newWorkItem = New-SCSMObject -class $prClass -propertyhashtable @{"ID" = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Problem")["Prefix"] + "{0}"; "Title" = $title; "Description" = $description; "Status" = "ProblemStatusEnum.Active$"; "Impact" = $irLowImpact; "Urgency" = $irLowUrgency} -PassThru @scsmMGMTParams
                        $prProjection = Get-SCSMObjectProjection -ProjectionName $prTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                        if($message.Attachments){$message.Attachments | Foreach-Object {Add-FileToSCSMObject -attachment $_ -smobject $newWorkItem}}
                        if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $newWorkItem}
                        Set-SCSMTemplate -Projection $prProjection -Template $PRTemplate
                        #Set-SCSMObjectTemplate -Projection $prProjection -Template $defaultPRTemplate @scsmMGMTParams
                        Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description; "Displayname" = "$($newWorkItem.Id): $title"} @scsmMGMTParams
                        #no Affected User to set on a Problem, set Created By using the Affected User object if it exists
                        if ($affectedUser)
                        {
                            try {New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "New-WorkItem" -EventId 3 -LogMessage "The Created By User for $($newWorkItem.Name) could not be set." -Severity "Warning"}}
                        }
                        if ($relatedUsers)
                        {
                            foreach ($relatedUser in $relatedUsers)
                            {
                                New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk @scsmMGMTParams
                            }
                        }

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterCreatePR }
                    }
            "cr" {
                        if ($UseMailboxRedirection -eq $true -and $TemplatesForThisMessage.Count -gt 0) {
                            $CRTemplate = Get-ScsmObjectTemplate -Id $($TemplatesForThisMessage["CRTemplate"]) @scsmMGMTParams
                        }
                        else {
                            $CRTemplate = $defaultCRTemplate
                        }
                        $newWorkItem = new-scsmobject -class $crClass -propertyhashtable @{"ID" = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.ChangeRequest")["Prefix"] + "{0}"; "Title" = $title; "Description" = $description; "Status" = "ChangeStatusEnum.New$"} -PassThru @scsmMGMTParams
                        $crProjection = Get-SCSMObjectProjection -ProjectionName $crTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                        #Set-SCSMObjectTemplate -Projection $crProjection -Template $defaultCRTemplate @scsmMGMTParams
                        Set-SCSMTemplate -Projection $crProjection -Template $CRTemplate
                        Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description; "Displayname" = "$($newWorkItem.Id): $title"} @scsmMGMTParams
                        #The Affected User relationship exists on Change Requests, but does not exist on the CR Form out of box.
                        #Cireson SCSM Portal customers may wish to set the Sender as the Affected User so that it follows Incident/Service Request style functionality in that the Sender/User
                        #in question can see the CR in the "My Requests" section of their SCSM portal. This can be acheived by uncommenting the New-SCSMRelationshipObject seen below
                        #Set Created By using the Affected User object if it exists
                        if ($affectedUser)
                        {
                            try {New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "New-WorkItem" -EventId 3 -LogMessage "The Created By User for $($newWorkItem.Name) could not be set." -Severity "Warning"}}
                            #New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
                        }
                        if ($relatedUsers)
                        {
                            foreach ($relatedUser in $relatedUsers)
                            {
                                New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk @scsmMGMTParams
                            }
                        }

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterCreateCR }
                    }
        }

        #if verbose logging, show details about the new work item
        if ($loggingLevel -ge 4)
        {
            $logMessage = "Created $wiType
            ID: $($newWorkItem.Name)
            Title: $($newWorkItem.Title)
            Affected User: $($affectedUser.DisplayName)
            Related Users: $($($relatedUsers.DisplayName) -join ',')"
            New-SMEXCOEvent -Source "New-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterCreateAnyWorkItem }

        if ($returnWIBool -eq $true)
        {
            return $newWorkItem
        }
    }
}

function Update-WorkItem
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #Message from Exchange. Typically called from the core processing loop that has simlified the original Exchange object into a distilled PSCustomObject
        [parameter(Mandatory=$true)]
        [PSCustomObject]$message,
        #The type of Work Item to create by it's two letter shorthand
        [parameter(Mandatory=$true)]
        [ValidateSet("IR", "SR", "CR", "PR", "RR", "RA", "MA", "PA", "SA", "DA")]
        [string]$wiType,
        #The Work Item that should be updated
        [parameter(Mandatory=$true)]
        $workitem
    )

    if ($PSCmdlet.ShouldProcess("$message","Update $workitem"))
    {
        if ($loggingLevel -ge 4)
        {
            $logMessage = "Updating $($workItem.Name)
            From: $($message.From)
            Title: $($message.Subject)"
            New-SMEXCOEvent -Source "Update-WorkItem" -EventId 0 -LogMessage $logMessage -Severity "Information"
        }

        #removes PII if RedactPiiFromMessage is enable
        if ($redactPiiFromMessage -eq $true)
        {
            $message.body = remove-PII -body $message.body
        }

        #determine the comment to add and ensure it's less than 4000 characters
        if ($includeWholeEmail -eq $true)
        {
            $commentToAdd = $message.body
            if ($commentToAdd.length -ge "4000")
            {
                $commentToAdd = $commentToAdd.substring(0, 4000)
            }
        }
        else
        {
            try
            {
                $fromKeywordPosition = $message.Body.IndexOf("$fromKeyword" + ":")
                if (($null -eq $fromKeywordPosition) -or ($fromKeywordPosition -eq -1))
                {
                    $commentToAdd = $message.body
                    if ($commentToAdd.length -ge "4000")
                    {
                        $commentToAdd = $commentToAdd.substring(0, 4000)
                    }
                }
                else
                {
                    $commentToAdd = $message.Body.substring(0, $fromKeywordPosition)
                    if ($commentToAdd.length -ge "4000")
                    {
                        $commentToAdd = $commentToAdd.substring(0, 4000)
                    }
                }
            }
            catch
            {
                $commentToAdd = $null
            }
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeUpdateAnyWorkItem }

        #determine who left the comment
        $commentLeftBy = Get-SCSMUserByEmailAddress -EmailAddress "$($message.From)"
        if ($commentLeftBy) {<#no change required#>}
        elseif ((!$commentLeftBy) -and ($createUsersNotInCMDB -eq $true) ){$commentLeftBy = New-CMDBUser -UserEmail $message.From}
        else {$commentLeftBy = New-CMDBUser -UserEmail $message.From -NoCommit}

        #add any attachments
        if ($message.Attachments)
        {
            switch ($wiType)
            {
                "ma" {
                        #$workItem = Get-SCSMObject -class $maClass -filter "Name -eq '$($workItem.Name)'" @scsmMGMTParams
                        $parentWorkItem = Get-SCSMWorkItemParent -WorkItemGUID $workItem.Get_Id().Guid
                        $message.Attachments | Foreach-Object {Add-FileToSCSMObject -attachment $_ -smobject $parentWorkItem}
                    }
                "ra" {
                        #$workItem = Get-SCSMObject -class $raClass -filter "Name -eq '$($workItem.Name)'" @scsmMGMTParams
                        $parentWorkItem = Get-SCSMWorkItemParent -WorkItemGUID $workItem.Get_Id().Guid
                        $message.Attachments | Foreach-Object {Add-FileToSCSMObject -attachment $_ -smobject $parentWorkItem}
                    }
                default { $message.Attachments | Foreach-Object {Add-FileToSCSMObject -attachment $_ -smobject $workItem} }
        }
        }
        #show the user who will perform the update and the [action] they are taking. If there is no [action] it's just a comment
        if ($loggingLevel -ge 4)
        {
            if ($commentToAdd -match '(?<=\[).*?(?=\])')
            {
                $logMessage = "Action for $($workItem.Name) invoked by:
                SCSM User: $($commentLeftBy.DisplayName)
                Action: $($commentToAdd -match '(?<=\[).*?(?=\])'|out-null;$matches[0])"
                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 1 -LogMessage $logMessage -Severity "Information"
            }
            else
            {
                $logMessage = "Leaving a Comment on $($workItem.Name) invoked by:
                SCSM User: $($commentLeftBy.DisplayName)
                Comment: $commentToAdd"
                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 1 -LogMessage $logMessage -Severity "Information"
            }
        }

        #update the work item with the comment and/or action
        switch ($wiType)
        {
            #### primary work item types ####
            "ir" {
                #$workItem = get-scsmobject -class $irClass -filter "Name -eq '$workItemID'" @scsmMGMTParams

                try {$existingWiStatusName = $workItem.Status.Name} catch {New-SMEXCOEvent -Source "New-WorkItem" -EventID 5 -Severity "Information" -LogMessage "$($workItem.Name) does not exist within SCSM."}
                if ($CreateNewWorkItemWhenClosed -eq $true -And $existingWiStatusName -eq "IncidentStatusEnum.Closed") {
                    $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" @scsmMGMTParams | foreach-object {Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $_ @scsmMGMTParams).sourceobject.id @scsmMGMTParams} | Sort-Object lastmodified -Descending | Select-Object -First 1 | where-object {$_.Status -ne "IncidentStatusEnum.Closed"}
                    if (($relatedWorkItemFromAttachmentSearch | get-unique).count -eq 1 -and $relatedWorkItemFromAttachmentSearch.Status.Name -ne "IncidentStatusEnum.Closed")
                    {
                        Update-WorkItem -message $message -wiType "ir" -workItem $relatedWorkItemFromAttachmentSearch
                    }
                    else
                    {
                        $newWi = New-WorkItem -message $message -wiType "ir" -returnWIBool $true

                        #copy essential info over from old to new
                        $NewDesc = "New ticket generated from reply to $($workItem.Name) (Closed). `n ---- `n $($newWi.Description) `n --- `n Original description: `n --- `n $($workItem.Description)"
                        $NewWiPropertiesFromOld = @{"Description"=$NewDesc;"TierQueue"=$($workItem.TierQueue);"Classification"=$($workItem.Classfification);"Impact"=$($workItem.Impact);"Urgency"=$($workItem.Urgency);}
                        Set-SCSMObject -SMObject $newWi -PropertyHashTable $newWiPropertiesFromOld @scsmMGMTParams

                        #relate old and new wi
                        New-SCSMRelationshipObject -Relationship $wiRelatesToWiRelClass -Source $newWi -Target $workItem -Bulk @scsmMGMTParams
                    }
                }
                else {
                    try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 7 -Severity "Warning" -LogMessage "The Affected User of $($workItem.Name) could not be found."}
                    if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 8 -Severity "Warning" -LogMessage "The Assigned User of $($workItem.Name) could not be found."}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    if ($assignedToSMTP.TargetAddress -eq $affectedUserSMTP.TargetAddress){$assignedToSMTP = $null}
                    #write to the Action log and take action on the Work Item if neccesary
                    switch ($message.From)
                    {
                        $affectedUserSMTP.TargetAddress {
                            if ($changeIncidentStatusOnReply -and (($workitem.Status.Name -ne "IncidentStatusEnum.Closed") -and ($workitem.Status.Name -ne "IncidentStatusEnum.Resolved"))) {try {Set-SCSMObject -SMObject $workItem -Property Status -Value "$changeIncidentStatusOnReplyAffectedUser" @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 6 -LogMessage "Attempting to change $($workItem.Name) to a Status of $($changeIncidentStatusOnReplyAffectedUser.DisplayName) due to Affected User Reply could not be performed. $($_.Exception)" -Severity "Warning"}}}
                            switch -Regex ($commentToAdd) {
                                "\[$acknowledgedKeyword]" {if ($null -eq $workItem.FirstResponseDate){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterAcknowledge }}}
                                "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; try {New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 3 -LogMessage "$($newWorkItem.Name) could not be Resolved By $($affectedUser.DisplayName)." -Severity "Warning"}}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Resolved" -IsPrivate $false; if ($defaultIncidentResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property ResolutionCategory -Value $defaultIncidentResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Closed" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "\[$takeKeyword]" {
                                    if ($takeRequiresGroupMembership -eq $false) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.TierQueue.Id)) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                    }
                                }
                                "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Undo-WorkItemResolution -WorkItem $workItem; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}}
                                "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "\[$irRegex[0-9]+\]")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem -message $message -wiType "ir" -returnWIBool $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk @scsmMGMTParams}catch{New-SMEXCOEvent -Source "Update-WorkItem" -EventID 9 -Severity "Warning" -LogMessage "$($workItem.Name) could not be related to $($returnedWorkItem.Name)"}; if ($ceScripts) { Invoke-AfterReactivate }}}
                                "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false}
                            }
                        }
                        $assignedToSMTP.TargetAddress {
                            if ($changeIncidentStatusOnReply -and (($workitem.Status.Name -ne "IncidentStatusEnum.Closed") -and ($workitem.Status.Name -ne "IncidentStatusEnum.Resolved"))) {try {Set-SCSMObject -SMObject $workItem -Property Status -Value "$changeIncidentStatusOnReplyAssignedTo" @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 6 -LogMessage "Attempting to change $($workItem.Name) to a Status of $($changeIncidentStatusOnReplyAssignedTo.DisplayName) due to Assigned User Reply could not be performed. $($_.Exception)" -Severity "Warning"}}}
                            switch -Regex ($commentToAdd) {
                                "\[$acknowledgedKeyword]" {if ($null -eq $workItem.FirstResponseDate){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterAcknowledge }}}
                                "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; try {New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 3 -LogMessage "$($newWorkItem.Name) could not be Resolved By $($affectedUser.DisplayName)." -Severity "Warning"}}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Resolved" -IsPrivate $false; if ($defaultIncidentResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property ResolutionCategory -Value $defaultIncidentResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Closed" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "\[$takeKeyword]" {
                                    if ($takeRequiresGroupMembership -eq $false) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.TierQueue.Id)) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                    }
                                }
                                "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Undo-WorkItemResolution -WorkItem $workItem; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}}
                                "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "\[$irRegex[0-9]+\]")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem -message $message -wiType "ir" -returnWIBool $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk @scsmMGMTParams}catch{New-SMEXCOEvent -Source "Update-WorkItem" -EventID 9 -Severity "Warning" -LogMessage "$($workItem.Name) could not be related to $($returnedWorkItem.Name)"}; if ($ceScripts) { Invoke-AfterReactivate }}}
                                "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                                "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $true}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                            }
                        }
                        default {
                            if ($changeIncidentStatusOnReply -and (($workitem.Status.Name -ne "IncidentStatusEnum.Closed") -and ($workitem.Status.Name -ne "IncidentStatusEnum.Resolved"))) {try {Set-SCSMObject -SMObject $workItem -Property Status -Value "$changeIncidentStatusOnReplyRelatedUser" @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 6 -LogMessage "Attempting to change $($workItem.Name) to a Status of $($changeIncidentStatusOnReplyRelatedUser.DisplayName) due to Related User Reply could not be performed. $($_.Exception)" -Severity "Warning"}}}
                            switch -Regex ($commentToAdd) {
                                "\[$acknowledgedKeyword]" {if ($null -eq $workItem.FirstResponseDate){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterAcknowledge }}}
                                "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; try {New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 3 -LogMessage "$($newWorkItem.Name) could not be Resolved By $($affectedUser.DisplayName)." -Severity "Warning"}}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Resolved" -IsPrivate $false; if ($defaultIncidentResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property ResolutionCategory -Value $defaultIncidentResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Closed" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "\[$takeKeyword]" {
                                    if ($takeRequiresGroupMembership -eq $false) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.TierQueue.Id)) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                    }
                                }
                                "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Undo-WorkItemResolution -WorkItem $workItem; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}}
                                "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "\[$irRegex[0-9]+\]")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem -message $message -wiType "ir" -returnWIBool $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk @scsmMGMTParams}catch{New-SMEXCOEvent -Source "Update-WorkItem" -EventID 9 -Severity "Warning" -LogMessage "$($workItem.Name) could not be related to $($returnedWorkItem.Name)"}; if ($ceScripts) { Invoke-AfterReactivate }}}
                                "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                                "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $true}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "$ExternalPartyCommentTypeIR" -IsPrivate $ExternalPartyCommentPrivacyIR}
                            }
                        }
                    }
                    #relate the user to the work item
                    try {New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 5 -LogMessage "$($commentLeftBy.DisplayName) could not be related to $($newWorkItem.Name)" -Severity "Warning"}}
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $workItem}
                }

                # Custom Event Handler
                if ($ceScripts) { Invoke-AfterUpdateIR }
            }
            "sr" {
                #$workItem = get-scsmobject -class $srClass -filter "Name -eq '$workItemID'" @scsmMGMTParams

                try {$existingWiStatusName = $workItem.Status.Name} catch {New-SMEXCOEvent -Source "New-WorkItem" -EventID 5 -Severity "Information" -LogMessage "$($workItem.Name) does not exist within SCSM."}
                if ($CreateNewWorkItemWhenClosed -eq $true -And $existingWiStatusName -eq "ServiceRequestStatusEnum.Closed") {
                    $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" @scsmMGMTParams | foreach-object {Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $_ @scsmMGMTParams).sourceobject.id @scsmMGMTParams} | Sort-Object lastmodified -Descending | Select-Object -First 1 | where-object {$_.Status -ne "ServiceRequestStatusEnum.Closed"}
                    if (($relatedWorkItemFromAttachmentSearch | get-unique).count -eq 1 -and $relatedWorkItemFromAttachmentSearch.Status.Name -ne "ServiceRequestStatusEnum.Closed")
                    {
                        Update-WorkItem -message $message -wiType "sr" -workItem $relatedWorkItemFromAttachmentSearch
                    }
                    else
                    {
                        $newWi = New-WorkItem -message $message -wiType "sr" -returnWIBool $true

                        #copy essential info over from old to new
                        $NewDesc = "New ticket generated from reply to $($workItem.Name) (Closed). `n ---- `n $($newWi.Description) `n --- `n Original description: `n --- `n $($workItem.Description)"
                        $NewWiPropertiesFromOld = @{"Description"=$NewDesc;"SupportGroup"=$($workItem.SupportGroup);"ServiceRequestCategory"=$($workItem.ServiceRequestCategory);"Priority"=$($workItem.Priority);"Urgency"=$($workItem.Urgency)}
                        Set-SCSMObject -SMObject $newWi -PropertyHashTable $newWiPropertiesFromOld @scsmMGMTParams

                        #relate old and new wi
                        New-SCSMRelationshipObject -Relationship $wiRelatesToWiRelClass -Source $newWi -Target $workItem -Bulk @scsmMGMTParams
                    }
                }
                else {
                    try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 7 -Severity "Warning" -LogMessage "The Affected User of $($workItem.Name) could not be found."}
                    if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 8 -Severity "Warning" -LogMessage "The Assigned User of $($workItem.Name) could not be found."}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    if ($assignedToSMTP.TargetAddress -eq $affectedUserSMTP.TargetAddress){$assignedToSMTP = $null}
                    switch ($message.From)
                    {
                        $affectedUserSMTP.TargetAddress {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$acknowledgedKeyword]" {if ($null -eq $workItem.FirstResponseDate){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false}}
                                "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                "\[$takeKeyword]" {
                                    if ($takeRequiresGroupMembership -eq $false) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.SupportGroup.Id)) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                    }
                                }
                                "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"CompletedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Completed$"; "Notes" = "$commentToAdd"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($defaultServiceRequestImplementationCategory) {Set-SCSMObject -SMObject $workItem -Property ImplementationResults -Value $defaultServiceRequestImplementationCategory}; if ($ceScripts) { Invoke-AfterCompleted }}
                                "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Canceled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                default {if($commentToAdd -match "#$privateCommentKeyword"){$isPrivateBool = $true}else{$isPrivateBool = $false};Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $isPrivateBool}
                            }
                        }
                        $assignedToSMTP.TargetAddress {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$acknowledgedKeyword]" {if ($null -eq $workItem.FirstResponseDate){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}}
                                "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                "\[$takeKeyword]" {
                                    if ($takeRequiresGroupMembership -eq $false) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.SupportGroup.Id)) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                    }
                                }
                                "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"CompletedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Completed$"; "Notes" = "$commentToAdd"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($defaultServiceRequestImplementationCategory) {Set-SCSMObject -SMObject $workItem -Property ImplementationResults -Value $defaultServiceRequestImplementationCategory}; if ($ceScripts) { Invoke-AfterCompleted }}
                                "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Canceled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $true}
                                "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                            }
                        }
                        default {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$acknowledgedKeyword]" {if ($null -eq $workItem.FirstResponseDate){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}}
                                "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                "\[$takeKeyword]" {
                                    if ($takeRequiresGroupMembership -eq $false) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.SupportGroup.Id)) {
                                        try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                        if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                    }
                                }
                                "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"CompletedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Completed$"; "Notes" = "$commentToAdd"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($defaultServiceRequestImplementationCategory) {Set-SCSMObject -SMObject $workItem -Property ImplementationResults -Value $defaultServiceRequestImplementationCategory}; if ($ceScripts) { Invoke-AfterCompleted }}
                                "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Canceled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $true}
                                "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "$ExternalPartyCommentTypeSR" -IsPrivate $ExternalPartyCommentPrivacySR}
                            }
                        }
                    }
                    #relate the user to the work item
                    try {New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 5 -LogMessage "$($commentLeftBy.DisplayName) could not be related to $($newWorkItem.Name)" -Severity "Warning"}}
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $workItem}
                }
                # Custom Event Handler
                if ($ceScripts) { Invoke-AfterUpdateSR }
            }
            "pr" {
                        #$workItem = get-scsmobject -class $prClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                        try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 8 -Severity "Warning" -LogMessage "The Assigned User of $($workItem.Name) could not be found."}
                        if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                        #write to the Action log
                        switch ($message.From)
                        {
                            $assignedToSMTP.TargetAddress {
                                switch -Regex ($commentToAdd)
                                {
                                    "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; try {New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 3 -LogMessage "$($newWorkItem.Name) could not be Resolved By $($affectedUser.DisplayName)." -Severity "Warning"}}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Resolved" -IsPrivate $false; if ($defaultProblemResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property Resolution -Value $defaultProblemResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                    "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                    "\[$takeKeyword]" {
                                        if ($takeRequiresGroupMembership -eq $false) {
                                            try {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk;} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false;
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            if ($ceScripts){ Invoke-AfterTake }
                                        }
                                        elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$prSupportGroupPropertyName.Id)) {
                                            try {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk;} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false;
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            if ($ceScripts){ Invoke-AfterTake }
                                        }
                                        else {
                                            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                        }
                                    }
                                    "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved") {Undo-WorkItemResolution -WorkItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}
                                    "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                                    default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                                }
                            }
                            default {
                                switch -Regex ($commentToAdd)
                                {
                                    "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; try {New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 3 -LogMessage "$($newWorkItem.Name) could not be Resolved By $($affectedUser.DisplayName)." -Severity "Warning"}}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Resolved" -IsPrivate $false; if ($defaultProblemResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property Resolution -Value $defaultProblemResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                    "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                    "\[$takeKeyword]" {
                                        if ($takeRequiresGroupMembership -eq $false) {
                                            try {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk;} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false;
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            if ($ceScripts){ Invoke-AfterTake }
                                        }
                                        elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$prSupportGroupPropertyName.Id)) {
                                            try {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk;} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false;
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            if ($ceScripts){ Invoke-AfterTake }
                                        }
                                        else {
                                            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                        }
                                    }
                                    "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved") {Undo-WorkItemResolution -WorkItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}
                                    "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                                    default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                                }
                            }
                        }
                        #relate the user to the work item
                        try {New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 5 -LogMessage "$($commentLeftBy.DisplayName) could not be related to $($newWorkItem.Name)" -Severity "Warning"}}
                        #add any new attachments
                        if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $workItem}

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterUpdatePR }
                    }
            "cr" {
                        #$workItem = get-scsmobject -class $crClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                        try{$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 8 -Severity "Warning" -LogMessage "The Assigned User of $($workItem.Name) could not be found."}
                        if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                        #write to the Action log and take action on the Work Item if neccesary
                        switch ($message.From)
                        {
                            $assignedToSMTP.TargetAddress {
                                switch -Regex ($commentToAdd)
                                {
                                    "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                    "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                    "\[$takeKeyword]" {
                                        if ($takeRequiresGroupMembership -eq $false) {
                                            try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterTake }
                                        }
                                        elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$crSupportGroupPropertyName.Id)) {
                                            try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterTake }
                                        }
                                        else {
                                            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                        }
                                    }
                                    {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "EndUserComment" -IsPrivate $false}
                                    "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $true}
                                    "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                                }
                            }
                            default {
                                switch -Regex ($commentToAdd)
                                {
                                    "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                    "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                    "\[$takeKeyword]" {
                                        if ($takeRequiresGroupMembership -eq $false) {
                                            try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterTake }
                                        }
                                        elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$crSupportGroupPropertyName.Id)) {
                                            try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                            if ($null -eq $workItem.FirstAssignedDate) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterTake }
                                        }
                                        else {
                                            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                        }
                                    }
                                    {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -EmailAddress $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                    "\[$addWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($addWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Add-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    "\[$removeWatchlistKeyword]" {if (($enableCiresonIntegration) -and ($removeWatchlistKeyword.Length -gt 1)) {$cpu = Get-CiresonPortalUser -username $commentLeftBy.UserName -domain $commentLeftBy.Domain; if ($cpu) {Remove-CiresonWatchListUser -userguid $cpu.Id -workitemguid $workItem.__InternalId} }}
                                    "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $true}
                                    default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                                }
                            }
                        }
                        #relate the user to the work item
                        try {New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 5 -LogMessage "$($commentLeftBy.DisplayName) could not be related to $($newWorkItem.Name)" -Severity "Warning"}}
                        #add any new attachments
                        if ($attachEmailToWorkItem -eq $true){Add-EmailToSCSMObject -message $message -smobject $workItem}

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterUpdateCR }
                    }

            #### activities ####
            "ra" {
                        #$workItem = get-scsmobject -class $raClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                        $isSenderValidReviewer = $false
                        $reviewers = Get-SCSMRelatedObject -SMObject $workItem -Relationship $raHasReviewerRelClass @scsmMGMTParams
                        foreach ($reviewer in $reviewers)
                        {
                            $reviewingUser = Get-SCSMRelatedObject -SMObject $reviewer -Relationship $raReviewerIsUserRelClass @scsmMGMTParams
                            if ($reviewingUser)
                            {
                                $reviewingUser = Get-SCSMObject -Id $reviewingUser.Id @scsmMGMTParams
                                $reviewingUserName = $reviewingUser.UserName #it is necessary to store this in its own variable for the AD filters to work correctly
                                $reviewingUserSMTP = Get-SCSMRelatedObject -SMObject $reviewingUser @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress

                                if ($commentToAdd.length -gt 256) { $decisionComment = $commentToAdd.substring(0,253)+"..." } else { $decisionComment = $commentToAdd }

                                #Reviewer is a User
                                if ([bool] (Get-ADUser @adParams -filter {SamAccountName -eq $reviewingUserName}))
                                {
                                    #approved
                                    if (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$approvedKeyword]"))
                                    {
                                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk @scsmMGMTParams
                                            if ($loggingLevel -ge 4)
                                            {
                                                $logMessage = "Voting on $($workItem.Name)
                                                SCSM User: $($commentLeftBy.DisplayName)
                                                Vote: $($commentToAdd -match '(?<=\[).*?(?=\])'|out-null;$matches[0])"
                                                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
                                            }
                                            $isSenderValidReviewer = $true
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterApproved }
                                    }
                                    #rejected
                                    elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                                    {
                                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk @scsmMGMTParams
                                            if ($loggingLevel -ge 4)
                                            {
                                                $logMessage = "Voting on $($workItem.Name)
                                                SCSM User: $($commentLeftBy.DisplayName)
                                                Vote: $($commentToAdd -match '(?<=\[).*?(?=\])'|out-null;$matches[0])"
                                                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
                                            }
                                            $isSenderValidReviewer = $true
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterRejected }
                                    }
                                    #no keyword, add a comment to parent work item
                                    elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and (($commentToAdd -notmatch "\[$approvedKeyword]") -or ($commentToAdd -notmatch "\[$rejectedKeyword]")))
                                    {
                                        $parentWorkItem = Get-SCSMWorkItemParent -WorkItemGUID $workItem.Get_Id().Guid
                                        switch ($parentWorkItem.Classname)
                                        {
                                            "System.WorkItem.ChangeRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                            "System.WorkItem.ServiceRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                            "System.WorkItem.Incident" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                        }
                                        if ($loggingLevel -ge 4)
                                        {
                                            $logMessage = "No vote to process for $($workItem.Name). Adding to Parent Work Item $($parentWorkItem.Name)
                                            SCSM User: $($commentLeftBy.DisplayName)
                                            Comment: $commentToAdd"
                                            New-SMEXCOEvent -Source "Update-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
                                        }
                                        $isSenderValidReviewer = $true
                                    }
                                }
                                else {
                                    # Identify the user
                                    $votedOnBehalfOfUser = Get-SCSMUserByEmailAddress -EmailAddress $message.From

                                    #Reviewer is in a Group and "Voting on Behalf of Groups" is enabled
                                    if ($voteOnBehalfOfGroups -eq $true -and $votedOnBehalfOfUser.UserName)
                                    {
                                        #determine which domain to query, in case of multiple domains and trusts
                                        $AdRoot = (Get-AdDomain @adParams -Identity $reviewingUser.Domain).DNSRoot

                                        $isReviewerGroupMember = Get-ADGroupMember -Server $AdRoot -Identity $reviewingUser.UserName -recursive @adParams | Where-Object { $_.objectClass -eq "user" -and $_.name -eq $votedOnBehalfOfUser.UserName }

                                        #approved on behalf of
                                        if (($isReviewerGroupMember) -and ($commentToAdd -match "\[$approvedKeyword]"))
                                        {
                                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $votedOnBehalfOfUser -Bulk @scsmMGMTParams
                                            if ($loggingLevel -ge 4)
                                            {
                                                $logMessage = "Voting on $($workItem.Name)
                                                SCSM User: $($commentLeftBy.DisplayName)
                                                On Behalf of: $($reviewingUser.UserName)
                                                Vote: $($commentToAdd -match '(?<=\[).*?(?=\])'|out-null;$matches[0])"
                                                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
                                            }
                                            $isSenderValidReviewer = $true
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterApprovedOnBehalf }

                                        }
                                        #rejected on behalf of
                                        elseif (($isReviewerGroupMember) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                                        {
                                            Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                            New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $votedOnBehalfOfUser -Bulk @scsmMGMTParams
                                            if ($loggingLevel -ge 4)
                                            {
                                                $logMessage = "Voting on $($workItem.Name)
                                                SCSM User: $($commentLeftBy.DisplayName)
                                                On Behalf of: $($reviewingUser.UserName)
                                                Vote: $($commentToAdd -match '(?<=\[).*?(?=\])'|out-null;$matches[0])"
                                                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
                                            }
                                            $isSenderValidReviewer = $true
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterRejectedOnBehalf }
                                        }
                                        #no keyword, add a comment to parent work item
                                        elseif (($isReviewerGroupMember) -and (($commentToAdd -notmatch "\[$approvedKeyword]") -or ($commentToAdd -notmatch "\[$rejectedKeyword]")))
                                        {
                                            $parentWorkItem = Get-SCSMWorkItemParent -WorkItemGUID $workItem.Get_Id().Guid
                                            switch ($parentWorkItem.Classname)
                                            {
                                                "System.WorkItem.ChangeRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $votedOnBehalfOfUser -Action "EndUserComment" -IsPrivate $false}
                                                "System.WorkItem.ServiceRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $votedOnBehalfOfUser -Action "EndUserComment" -IsPrivate $false}
                                                "System.WorkItem.Incident" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $votedOnBehalfOfUser -Action "EndUserComment" -IsPrivate $false}
                                            }
                                            if ($loggingLevel -ge 4)
                                            {
                                                $logMessage = "No vote to process on behalf of for $($workItem.Name)
                                                SCSM User: $($commentLeftBy.DisplayName)
                                                Vote: $commentToAdd"
                                                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 2 -LogMessage $logMessage -Severity "Information"
                                            }
                                            $isSenderValidReviewer = $true
                                        }
                                        else {
                                            if ($loggingLevel -ge 3)
                                            {
                                                #user is not a member of the group. The keyword may or may not exist.
                                                $logMessage = "AD User: $($votedOnBehalfOfUser.UserName) could not vote on behalf of AD Group:$($reviewingUser.UserName).
                                                They are either not a member of the AD Group or their Comment did not contain a valid keyword. Their comment was:
                                                $commentToAdd"
                                                New-SMEXCOEvent -Source "Update-WorkItem" -EventId 11 -Severity "Error" -LogMessage $logMessage
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if ($loggingLevel -ge 3)
                                        {
                                            #Vote on Behalf of AD groups is either disabled, or the user couldn't be found in SCSM
                                            $logMessage = "Voting On Behalf of AD Groups is currently set to: $($voteOnBehalfOfGroups.ToString())
                                            SCSM User: User Display/User Name: $($votedOnBehalfOfUser.DisplayName) / $($votedOnBehalfOfUser.Username)
                                            Vote: $commentToAdd"
                                            New-SMEXCOEvent -Source "Update-WorkItem" -EventId 12 -Severity "Error" -LogMessage $logMessage
                                        }
                                    }
                                }
                            }
                        }

                        #user wasn't a reviewer, log an event
                        if (($isSenderValidReviewer -eq $false) -and ($loggingLevel -ge 3))
                        {
                            $logMessage = "$($message.From)/$($commentLeftBy.DisplayName) could not be matched to a corresponding Reviewer on $($workItem.Name). Either they are not a Reviewer or their User object in SCSM does not have a valid and related SMTP Notification Channel"
                            New-SMEXCOEvent -Source "Update-WorkItem" -EventId 13 -LogMessage $logMessage -Severity "Error"
                        }

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterUpdateRA }
                    }
            "ma" {
                        #$workItem = get-scsmobject -class $maClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                        try {$activityImplementer = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {New-SMEXCOEvent -Source "Update-WorkItem" -EventID 8 -Severity "Warning" -LogMessage "The Activity Implementer of $($workItem.Name) could not be found."}
                        if ($activityImplementer){$activityImplementerSMTP = Get-SCSMRelatedObject -SMObject $activityImplementer @scsmMGMTParams | Where-Object{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                        switch ($message.From)
                        {
                            $activityImplementerSMTP.TargetAddress {
                                switch -Regex ($commentToAdd)
                                {
                                    "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Completed$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams; if ($ceScripts) { Invoke-AfterCompleted }}
                                    "\[$skipKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Skipped$"; "Skip" = $true; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams; if ($ceScripts) { Invoke-AfterSkipped }}
                                    default {
                                        $parentWorkItem = Get-SCSMWorkItemParent -WorkItemGUID $workItem.Get_Id().Guid
                                        switch ($parentWorkItem.Classname)
                                        {
                                            "System.WorkItem.ChangeRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                            "System.WorkItem.ServiceRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                            "System.WorkItem.Incident" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                        }
                                    }
                                }
                            }
                            default {
                                switch -Regex ($commentToAdd)
                                {
                                    "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Completed$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams; if ($ceScripts) { Invoke-AfterCompleted }}
                                    "\[$skipKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Skipped$"; "Skip" = $true; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams; if ($ceScripts) { Invoke-AfterSkipped }}
                                    "\[$takeKeyword]" {
                                        if ($takeRequiresGroupMembership -eq $false) {
                                            try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterTake }
                                        }
                                        elseif (($takeRequiresGroupMembership -eq $true) -and (Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$maSupportGroupPropertyName.Id)) {
                                            try {New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk} catch {if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 4 -LogMessage "$($newWorkItem.Name) could not be taken by $($commentLeftBy.DisplayName)." -Severity "Warning"}}
                                            # Custom Event Handler
                                            if ($ceScripts) { Invoke-AfterTake }
                                        }
                                        else {
                                            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Update-WorkItem" -EventId 10 -Severity "Error" -LogMessage "Self-Assignment for $($workItem.Name) via the [take] keyword(s) failed for $($commentLeftBy.DisplayName)."}
                                        }
                                    }
                                    default {
                                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Notes" = "$($workItem.Notes)$($commentLeftBy.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams
                                    }
                                }
                            }
                        }

                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterUpdateMA }
                    }
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterUpdateAnyWorkItem }
    }
}

function Add-EmailToSCSMObject
{
    param (
        #the Exchange message to add to the SCSM Work Item
        $message,
        #SCSM Work Item that will have an Exchange message added to it
        $smobject
    )

    #determine if the incoming object is a WorkItem or ConfigItem
    if ($smobject.ClassName -like "*WorkItem*")
    {
        $itemType = "WorkItem"

        # Get attachment limits and attachment count in ticket, if configured to
        if ($checkAttachmentSettings -eq $true) {
            $workItemSettings = Get-SCSMWorkItemSetting -WorkItemClass $smobject.ClassName

            # Get count of attachments already in ticket
            try {$existingAttachmentsCount = (Get-ScsmRelatedObject @scsmMGMTParams -SMObject $smobject -Relationship $wiHasFileAttachRelClass).Count} catch {$existingAttachmentsCount = 0}
        }
    }
    else {$itemType = "ConfigItem"}

    try
    {
        $messageMime = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
        $MemoryStream = New-Object System.IO.MemoryStream($messageMime.MimeContent.Content,0,$messageMime.MimeContent.Content.Length)

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeAttachEmail }

        # if #checkAttachmentSettings -eq $true, test whether the email size (IN KB!) exceeds the limit and if the number of existing attachments is under the limit
        $workItemAttachmentCriteria = if ($itemType -eq "WorkItem"){$checkAttachmentSettings -eq $false -or (($MemoryStream.Length / 1024) -le $($workItemSettings["MaxAttachmentSize"]) -and $existingAttachmentsCount -le $($workItemSettings["MaxAttachments"]))}
        if (($itemType -eq "WorkItem" -and $workItemAttachmentCriteria) -or ($itemType -eq "ConfigItem"))
        {
            #Create the attachment object itself and set its properties for SCSM
            $emailAttachment = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
            $emailAttachment.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
            $emailAttachment.Item($fileAttachmentClass, "DisplayName").Value = "message.eml"
            $emailAttachment.Item($fileAttachmentClass, "Description").Value = "ExchangeConversationID:$($message.ConversationID);"
            $emailAttachment.Item($fileAttachmentClass, "Extension").Value = ".eml"
            $emailAttachment.Item($fileAttachmentClass, "Size").Value = $MemoryStream.Length
            $emailAttachment.Item($fileAttachmentClass, "AddedDate").Value = [DateTime]::Now.ToUniversalTime()
            $emailAttachment.Item($fileAttachmentClass, "Content").Value = $MemoryStream

            #Add the attachment to the work item and commit the changes
            $WorkItemProjection = Get-SCSMObjectProjection "System.$itemType.Projection" -Filter "id -eq '$($smobject.Id)'" @scsmMGMTParams
            switch ($itemType)
            {
                "WorkItem" {$WorkItemProjection.__base.Add($emailAttachment, $wiHasFileAttachRelClass.Target)}
                "ConfigItem" {$WorkItemProjection.__base.Add($emailAttachment, $ciHasFileAttachRelClass.Target)}
            }
            $WorkItemProjection.__base.Commit()

            #create the Attached By relationship if possible
            $attachedByUser = Get-SCSMUserByEmailAddress -EmailAddress "$($message.from)"
            if ($attachedByUser)
            {
                New-SCSMRelationshipObject -Source $emailAttachment -Relationship $fileAddedByUserRelClass -Target $attachedByUser @scsmMGMTParams -Bulk
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterAttachEmail }
        }
        else
        {
            if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Add-EmailToSCSMObject" -EventID 0 -Severity "Warning" -LogMessage "Email from $($message.From) on $($smobject.DisplayName) was not attached. Current Attachment Count: $existingAttachmentsCount/$($workItemSettings["MaxAttachments"]). File Size/Allowed Size: $($MemoryStream.Length/1024)/$($workItemSettings["MaxAttachmentSize"])"}
        }
    }
    catch
    {
        if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Add-EmailToSCSMObject" -EventID 1 -Severity "Warning" -LogMessage "Email from $($message.From) on $($smobject.DisplayName) could not be attached. $($_.Exception)"}
    }
}

#inspired and modified from Stefan Roth here - https://stefanroth.net/2015/03/28/scsm-passing-attachments-via-web-service-e-g-sma-web-service/
function Add-FileToSCSMObject
{
    param (
        #the attachment on the message to add to the SCSM Work Item
        $attachment,
        #SCSM Work Item that will have an attachment added to it
        $smobject
    )

    #determine if the incoming object is a WorkItem or ConfigItem
    if ($smobject.ClassName -like "*WorkItem*")
    {
        $itemType = "WorkItem"

        # Get attachment limits and attachment count in ticket, if configured to
        if ($checkAttachmentSettings -eq $true) {
            $workItemSettings = Get-SCSMWorkItemSetting -WorkItemClass $smobject.ClassName
            $attachMaxSize = $workItemSettings["MaxAttachmentSize"]
            $attachMaxCount = $workItemSettings["MaxAttachments"]

            # Get count of attachments already in ticket
            $existingAttachments = Get-ScsmRelatedObject @scsmMGMTParams -SMObject $smobject -Relationship $wiHasFileAttachRelClass
            # Only use at before the loop
            try {$existingAttachmentsCount = $existingAttachments.Count } catch { $existingAttachmentsCount = 0 }
        }
    }
    else {$itemType = "ConfigItem"}

    # Custom Event Handler
    if ($ceScripts) { Invoke-BeforeAttachFiles }

        #file attachment logging
        if ($loggingLevel -ge 4)
        {
            $logMessage = "Attaching File
            From: $($message.From)
            Title: $($message.Subject)"
            New-SMEXCOEvent -Source "Add-FileToSCSMObject" -EventId 1 -LogMessage $logMessage -Severity "Information"
        }

        try
        {
            #determine if a File Attachment
            if ($attachment.GetType().Name -eq "FileAttachment")
            {
                $attachment.Load()
                if ($attachment.GetType().BaseType.Name -like "Mime*")
                {
                    #This is a signed/encrypted attachment
                    $signedAttachArray = $attachment.ContentObject.Stream.ToArray()
                    $base64attachment = [System.Convert]::ToBase64String($signedAttachArray)
                    $AttachmentContent = [convert]::FromBase64String($base64attachment)

                    #Create a new MemoryStream object out of the attachment data
                    $MemoryStream = New-Object System.IO.MemoryStream($signedAttachArray,0,$signedAttachArray.Length)
                }
                else
                {
                    #this is a regular File Attachment
                    $base64attachment = [System.Convert]::ToBase64String($attachment.Content)
                    $AttachmentContent = [convert]::FromBase64String($base64attachment)

                    #Create a new MemoryStream object out of the attachment data
                    $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)
                }
            }
            #determine if an Item Attachment
            elseif ($attachment.GetType().Name -eq "ItemAttachment")
            {
                if ($attachment.GetType().BaseType.Name -like "Mime*")
                {
                    #This is a signed/encrypted attachment
                    $signedAttachArray = $attachment.ContentObject.Stream.ToArray()
                    $base64attachment = [System.Convert]::ToBase64String($signedAttachArray)
                    $AttachmentContent = [convert]::FromBase64String($base64attachment)

                    #Create a new MemoryStream object out of the attachment data
                    $MemoryStream = New-Object System.IO.MemoryStream($signedAttachArray,0,$signedAttachArray.Length)
                }
                else
                {
                    $attachment.Load($mimeContentSchema)
                    $base64attachment = [System.Convert]::ToBase64String($attachment.Item.MimeContent.Content)
                    $AttachmentContent = [convert]::FromBase64String($base64attachment)

                    #Create a new MemoryStream object out of the attachment data
                    $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)

                    #update the $attachment variable's name
                    if ($attachment.Item.GetType().Name -eq "EmailMessage")
                    {
                        $attachment | Add-Member -NotePropertyName DisplayName -NotePropertyValue ($attachment.Name + ".eml")
                        $attachment | Add-Member -NotePropertyName Extension -NotePropertyValue ".eml"
                    }
                }
            }
            #must be part of a digitally signed/encrypted message, determine Mime Object type
            else
            {
                #generic mime attachment
                if ($attachment.GetType().Name -eq "MimePart")
                {
                    #Create a new MemoryStream object out of the attachment data
                    $MemoryStream = New-Object System.IO.MemoryStream #($signedAttachArray,0,$signedAttachArray.Length)
                    $attachment.Content.DecodeTo($MemoryStream)

                    #update the $attachment variable's name
                    $attachment | Add-Member -NotePropertyName Name -NotePropertyValue $attachment.FileName
                }
                #exchange object that is a mime type
                if ($attachment.GetType().Name -eq "MessagePart")
                {
                    #This is an attached email
                    $MemoryStream = New-Object System.IO.MemoryStream
                    $attachment.WriteTo($MemoryStream)

                    #$attachment.Load($mimeContentSchema)
                    $base64attachment = [System.Convert]::ToBase64String($MemoryStream.ToArray())
                    $AttachmentContent = [convert]::FromBase64String($base64attachment)

                    #Create a new MemoryStream object out of the attachment data
                    $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)

                    #update the $attachment variable's name
                    $attachment | Add-Member -NotePropertyName Name -NotePropertyValue "message.eml"
                    $attachment | Add-Member -NotePropertyName Extension -NotePropertyValue ".eml"
                }
            }

            #create the File Attachment object for SCSM
            $workItemAttachmentCriteria = if ($itemType -eq "WorkItem"){$MemoryStream.Length -gt $minFileSizeInKB+"kb" -and ($checkAttachmentSettings -eq $false -or ($existingAttachmentsCount -lt $attachMaxCount -And $MemoryStream.Length -le "$attachMaxSize"+"mb"))}
            if (($itemType -eq "WorkItem" -and $workItemAttachmentCriteria) -or ($itemType -eq "ConfigItem"))
            {
                #Create the attachment object itself and set its properties for SCSM
                $NewFile = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
                $NewFile.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
                if (!($attachment.DisplayName))
                {
                    $NewFile.Item($fileAttachmentClass, "DisplayName").Value = $attachment.Name.Split([IO.Path]::GetInvalidFileNameChars()) -join ""
                }
                else
                {
                    $NewFile.Item($fileAttachmentClass, "DisplayName").Value = $attachment.DisplayName.Split([IO.Path]::GetInvalidFileNameChars()) -join ""
                }
                #optional, use Azure Cognitive Services Vision, OCR, or Speech to set the Description property on the file
                try
                {
                    #The $Attachment does not have an Extension property, attempt to generate one for the SCSM File Object from the $Attachment.Name
                    if (!($attachment.Extension))
                    {
                        $fileExtensionArrayPosition = $attachment.Name.Split(".").Length - 1
                        $NewFile.Item($fileAttachmentClass, "Extension").Value = "." + $attachment.Name.Split(".")[$fileExtensionArrayPosition]
                    }
                    #The $Attachment has a Extension defined, set the SCSM File Attachment object's Extension property based on the Attachment's Extension
                    else
                    {
                        if ($attachment.Extension.StartsWith("."))
                        {
                            $NewFile.Item($fileAttachmentClass, "Extension").Value = $attachment.Extension
                        }
                        else
                        {
                            $NewFile.Item($fileAttachmentClass, "Extension").Value = "." + $attachment.Extension
                        }

                    }

                    #See if the File Extension is a known image type and if Azure Vision is being used
                    if (((".png", ".jpg", ".jpeg", ".bmp", ".gif") -contains $NewFile.Item($fileAttachmentClass, "Extension").Value) -and ($enableAzureVision))
                    {
                        $azureVisionResult = Get-AzureEmailImageAnalysis -imageToEvalute $AttachmentContent
                        $azureVisionTags = $azureVisionResult.tags.name | select-object -first 5
                        $azureVisionTags = $azureVisionTags -join ','
                        #if one of the Tags is "text" then attempt to extract text from the image through OCR as long as the confidence is greater than 90
                        if ($azureVisionTags.contains("text"))
                        {
                            $AzureOCRConfidence = [math]::round((($azureVisionTags.tags | select-object name, confidence | Where-Object{$_.name -eq "text"} | select-object confidence -ExpandProperty confidence) * 100), 2)
                            if ($AzureOCRConfidence -ge 90)
                            {
                                $azureImageText = Get-AzureEmailImageText -imageToEvalute $AttachmentContent
                                $ocrResult = $azureImageText.regions.Lines.words.text -join " "
                                if ($ocrResult.length -ge 256){$ocrResult = $ocrResult.Substring(0,255)}
                                #set the Description on the File Attachment with the Tags + OCR result
                                $NewFile.Item($fileAttachmentClass, "Description").Value = "Tags:$($azureVisionTags);Desc:$ocrResult"
                            }
                            else
                            {
                                #The OCR confidence wasn't high enough to test
                                $NewFile.Item($fileAttachmentClass, "Description").Value = "Tags:$($azureVisionTags)"
                            }
                        }
                        else
                        {
                            #The returned Azure Vision Tags don't contain the word "text"
                            $NewFile.Item($fileAttachmentClass, "Description").Value = "Tags:$($azureVisionTags)"
                        }
                    }
                    #See if the File Extension is a known audio type and if Azure Speech is being used
                    if (((".wav", ".ogg") -contains $NewFile.Item($fileAttachmentClass, "Extension").Value) -and ($enableAzureSpeech))
                    {
                        $NewFile.Item($fileAttachmentClass, "Description").Value = (Get-AzureSpeechEmailAudioText -audioFileToEvaluate $attachmentContent).DisplayText
                    }
                }
                catch
                {
                    #file doesn't have a parseable extension or the call to Azure Vision/Speech failed
                    New-SMEXCOEvent -Source "Add-FileToSCSMObject" -EventID 2 -Severity "Warning" -LogMessage $_.Exception
                }
                $NewFile.Item($fileAttachmentClass, "Size").Value = $MemoryStream.Length
                $NewFile.Item($fileAttachmentClass, "AddedDate").Value = [DateTime]::Now.ToUniversalTime()
                $NewFile.Item($fileAttachmentClass, "Content").Value = $MemoryStream

                #Add the attachment to the work/config item and commit the changes
                $WorkItemProjection = Get-SCSMObjectProjection "System.$itemType.Projection" -Filter "id -eq '$($smObject.Id)'" @scsmMGMTParams
                switch ($itemType)
                {
                    "WorkItem" {$WorkItemProjection.__base.Add($NewFile, $wiHasFileAttachRelClass.Target)}
                    "ConfigItem" {$WorkItemProjection.__base.Add($NewFile, $ciHasFileAttachRelClass.Target)}
                }
                $WorkItemProjection.__base.Commit()

                #create the Attached By relationship if possible
                $attachedByUser = Get-SCSMUserByEmailAddress -EmailAddress "$($message.from)"
                if ($attachedByUser)
                {
                    New-SCSMRelationshipObject -Source $NewFile -Relationship $fileAddedByUserRelClass -Target $attachedByUser @scsmMGMTParams -Bulk
                }
            }
        }
        catch
        {
            #file could not be added
            if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Add-FileToSCSMObject" -EventID 0 -Severity "Warning" -LogMessage "A File Attachment from $($message.From) could not be added to $($smObject.Name). $($_.Exception)"}
        }

    # Custom Event Handler
    if ($ceScripts) { Invoke-AfterAttachFiles }
}

function Get-WorkItem
{
    param (
        #The SCSM Work Item ID identified in the Exchange Message Subject
        [parameter(Mandatory=$true)]
        [string]$workItemID,
        #The SCSM Work Item Class object
        [parameter(Mandatory=$true)]
        [Microsoft.EnterpriseManagement.Configuration.ManagementPackClass]$workItemClass
    )

    #removes [] surrounding a Work Item ID if neccesary
    if ($workitemID.StartsWith("[") -and $workitemID.EndsWith("]"))
    {
        $workitemID = $workitemID.TrimStart("[").TrimEnd("]")
    }

    #get the work item
    $wi = get-scsmobject -Class $workItemClass -Filter "Name -eq '$workItemID'" @scsmMGMTParams
    return $wi
}

function Get-SCSMUserByEmailAddress
{
    param (
        #The Email Address to look up an SCSM User by
        [parameter(Mandatory=$true)]
        [string]$EmailAddress
    )

    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq $EmailAddress" @scsmMGMTParams | sort-object lastmodified -Descending | select-object -first 1
    if ($userSMTPNotification)
    {
        $user = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification @scsmMGMTParams).sourceObject.id @scsmMGMTParams
        if ($loggingLevel -ge 4){New-SMEXCOEvent -Source "Get-SCSMUserByEmailAddress" -EventId 0 -LogMessage "Address: $EmailAddress was matched to SCSM User: $($user.Domain)\$($user.UserName)" -Severity "Information"}
        return $user
    }
    else
    {
        if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Get-SCSMUserByEmailAddress" -EventId 1 -LogMessage "Address: $EmailAddress could not be matched to a user in SCSM" -Severity "Warning"}
        return $null
    }
}

# Nested group membership check inspired by Piotr Lewandowski https://gallery.technet.microsoft.com/scriptcenter/Get-nested-group-15f725f2#content
function Get-TierMembership {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'UserSamAccountName', Justification = 'False positive')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'TierId', Justification = 'False positive')]
    param (
        #The SCSM User who will be determined to be part of the given Tier
        [parameter(Mandatory=$true)]
        [string]$UserSamAccountName,
        #The Tier (Support Group) to validate if the UserSamAccountName is a part of
        [parameter(Mandatory=$true)]
        $TierId
    )

    try
    {
        #define classes
        $mapCls = Get-SCSMClass @scsmMGMTParams -Name "Cireson.SupportGroupMapping$"

        #pull the group based on support tier mapping
        $mapping = $mapCls | Get-SCSMObject @scsmMGMTParams | Where-Object { $_.SupportGroupId.Guid -eq $TierId.Guid }
        $groupId = $mapping.AdGroupId

        #get the AD group object name
        $grpInScsm = (Get-SCSMObject @scsmMGMTParams -Id $groupId)
        $grpSamAccountName = $grpInScsm.UserName

        #determine which domain to query, in case of multiple domains and trusts
        $AdRoot = (Get-ADDomain @adParams -Identity $grpInScsm.Domain).DNSRoot

        if ($grpSamAccountName) {
            # Get the group membership
            [array]$members = Get-ADGroupMember @adParams -Server $AdRoot -Identity $grpSamAccountName -Recursive | Where-Object {$_.objectClass -eq "user"}

            # check if the members of the AD group that underpins this support group contains the user
            $groupContainsMember = $members | Where-Object {$_.SamAccountName -eq $UserSamAccountName}
            if ($groupContainsMember) {
                $isMember = $true
            }
            else {
                $isMember = $false
            }
            return $isMember
        }
    }
    catch
    {
        if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-TierMembership" -EventID 0 -Severity "Warning" -LogMessage $_.Exception}
    }
}

function Get-TierMember
{
    param (
        #The Tier (Support Group) to retrieve the members for
        [parameter(Mandatory=$true)]
        $TierEnumId
    )

    $supportTierMembers = $null

    try
    {
        #logging
        if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-TierMember" -EventID 0 -Severity "Information" -LogMessage "Get AD Group Associated with enum: $TierEnumId"}

        #define classes
        $mapCls = Get-ScsmClass @scsmMGMTParams -Name "Cireson.SupportGroupMapping$"

        #pull the group based on support tier mapping
        $mapping = $mapCls | Get-ScsmObject @scsmMGMTParams | Where-Object { $_.SupportGroupId.Guid -eq $TierEnumId }
        $groupId = $mapping.AdGroupId
        if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-TierMember" -EventID 1 -Severity "Information" -LogMessage "Get SCSM object/Group for: $groupId"}

        #get the AD group object name
        $grpInScsm = (Get-ScsmObject @scsmMGMTParams -Id $groupId)
        $grpSamAccountName = $grpInScsm.UserName
        if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-TierMember" -EventID 2 -Severity "Information" -LogMessage "AD Group Name: $grpSamAccountName"}

        #determine which domain to query, in case of multiple domains and trusts
        $AdRoot = (Get-AdDomain @adParams -Identity $grpInScsm.Domain).DNSRoot

        if ($grpSamAccountName)
        {
            # Get the group membership
            [array]$supportTierMembers = Get-ADGroupMember @adParams -Server $AdRoot -Identity $grpSamAccountName -Recursive | foreach-object {Get-SCSMObject -Class $domainUserClass -filter "Username -eq '$($_.samaccountname)'"}
            if ($loggingLevel -ge 4) {
                $supportTierMembersLogString = $supportTierMembers.Name -join ","
                New-SMEXCOEvent -Source "Get-TierMember" -EventID 3 -Severity "Information" -LogMessage "AD Group Members: $supportTierMembersLogString"
            }
        }
    }
    catch
    {
        if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Get-TierMember" -EventID 4 -Severity "Warning" -LogMessage $_.Exception}
    }
    return $supportTierMembers
}

function Get-AssignedToWorkItemVolume
{
    param (
        #The SCSM User to retrieve to see how many open Work Items are assigned to them
        [parameter(Mandatory=$true)]
        $SCSMUser
    )

    #initialize the counter, get the user's assigned Work Items that aren't in some form of "Done"
    $assignedCount = 0
    $assignedWorkItemRelationships = Get-SCSMRelationshipObject -TargetRelationship $assignedToUserRelClass -TargetObject $SCSMUser @scsmMGMTParams
    $assignedWorkItemRelationships = $assignedWorkItemRelationships | select-object SourceObject -ExpandProperty SourceObject | select-object -ExpandProperty values | Where-Object{($_.type.name -eq "Status") -and (($_.value -notlike "*Resolve*") -and ($_.value -notlike "*Close*") -and ($_.value -notlike "*Complete*") -and ($_.value -notlike "*Skip*") -and ($_.value -notlike "*Cancel*"))}
    $assignedWorkItemRelationships | foreach-object {$assignedCount++}

    #build Assigned To Volume object
    $assignedToVolume = [PSCustomObject] @{
        SCSMUser      = $SCSMUser
        AssignedCount = $assignedCount
    }
    if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-AssignedToWorkItemVolume" -EventID 0 -Severity "Information" -LogMessage "$($assignedToVolume.SCSMUser.DisplayName) : $($assignedToVolume.AssignedCount)"}
    return $assignedToVolume
}

function Set-AssignedToPerSupportGroup
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The GUID of the Enum for the Support Group/Tier
        [parameter(Mandatory=$true)]
        $SupportGroupID,
        #The Work Item that whose Support Group/Tier should be updated
        [parameter(Mandatory=$true)]
        $WorkItem
    )

    if ($PSCmdlet.ShouldProcess("$($workItem.Name)",'Set Support Group/Tier'))
    {
        #log how Dynamic Analyst Assignment will be executed
        if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Set-AssignedToPerSupportGroup" -EventID 0 -Severity "Information" -LogMessage "Using Dynamic Analyst Assignment: $DynamicWorkItemAssignment"}

        #get the template's support group members
        $supportGroupMembers = Get-TierMember -TierEnumID $SupportGroupID

        #based on how Dynamic Work Item assignment was configured, set the Assigned To User
        switch ($DynamicWorkItemAssignment)
        {
            "volume" {$userToAssign = $supportGroupMembers | foreach-object {Get-AssignedToWorkItemVolume -SCSMUser $_} | Sort-Object AssignedCount | Select-Object SCSMUser -ExpandProperty SCSMUser -first 1 }
            "OOOvolume" {$userToAssign = $supportGroupMembers | Where-Object {$_.OutOfOffice -ne $true} | foreach-object {Get-AssignedToWorkItemVolume -SCSMUser $_} | Sort-Object AssignedCount | Select-Object SCSMUser -ExpandProperty SCSMUser -first 1}
            "random" {$userToAssign = $supportGroupMembers | Get-Random}
            "OOOrandom" {$userToAssign = $supportGroupMembers | Where-Object {$_.OutOfOffice -ne $true} | Get-Random}
            default {if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Set-AssignedToPerSupportGroup" -EventID 1 -Severity "Error" -LogMessage "$DynamicWorkItemAssignment is not supported. No user will be assigned to $($WorkItem.Name). Please use the Settings UI to properly set this value."}}
        }

        #assign the work item to the selected user and set the first assigned date
        if ($userToAssign)
        {
            New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $WorkItem -Target $userToAssign -Bulk @scsmMGMTParams
            Set-SCSMObject -SMObject $WorkItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams
            if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Set-AssignedToPerSupportGroup" -EventID 2 -Severity "Information" -LogMessage "Assigned $($WorkItem.Name) to $($userToAssign.Name)"}
        }
    }
}

#courtesy of Leigh Kilday. Modified.
function Get-SCSMWorkItemParent
{
    [CmdletBinding()]
    PARAM (
        #The GUID of the Work Item
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
                if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-SCSMWorkItemParent" -EventID 0 -Severity "Information" -LogMessage "[PROCESS] Retrieving WI with GUID"}
                $ActivityObject = Get-SCSMObject -Id $WorkItemGUID @scsmMGMTParams
            }

            #Retrieve Parent
            if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-SCSMWorkItemParent" -EventID 1 -Severity "Information" -LogMessage "[PROCESS] Activity: $($ActivityObject.Name)"}
            if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-SCSMWorkItemParent" -EventID 2 -Severity "Information" -LogMessage "[PROCESS] Retrieving WI Parent"}
            $ParentRelatedObject = Get-SCSMRelationshipObject -ByTarget $ActivityObject @scsmMGMTParams | Where-Object{$_.RelationshipID -eq $wiContainsActivityRelClass.id.Guid}
            $ParentObject = $ParentRelatedObject.SourceObject

            if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-SCSMWorkItemParent" -EventID 3 -Severity "Information" -LogMessage "[PROCESS] Activity: $($ActivityObject.Name) - Parent: $($ParentObject.Name)"}

            If ($ParentObject.ClassName -eq 'System.WorkItem.ServiceRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ChangeRequest' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.ReleaseRecord' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Incident' `
            -or $ParentObject.ClassName -eq 'System.WorkItem.Problem')
            {
                if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-SCSMWorkItemParent" -EventID 4 -Severity "Information" -LogMessage "[PROCESS] This is the top level parent"}

                #return parent object Work Item
                Return $ParentObject
            }
            Else
            {
                if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-SCSMWorkItemParent" -EventID 5 -Severity "Information" -LogMessage "[PROCESS] Not the top level parent. Running against this object"}
                Get-SCSMWorkItemParent -WorkItemGUID $ParentObject.Id.GUID @scsmMGMTParams
            }
        }
        CATCH
        {
            Write-Error -Message $Error[0].Exception.Message
        }
    }
}

#inspired and modified from Travis Wright here - https://techcommunity.microsoft.com/t5/system-center-blog/creating-membership-and-hosting-objects-relationships-using-new/ba-p/347537
function New-CMDBUser
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The Email Address to use to the create the System.Domain.User object in SCSM
        [Parameter(Mandatory=$true)]
        [string]$userEmail,
        #If present, the user is only created in memory and never saved into SCSM
        [Parameter(Mandatory=$false)]
        [switch]$NoCommit
    )
    #The ID for external users appears to be a GUID, but it can't be identified by get-scsmobject
    #The ID for internal domain users takes the form of domain_username_SMTP
    #It's unclear how this ID should be generated. Opted to take the form of an internal domain for the ID
    #By using the internal domain style (_SMTP) this means New/Update Work Item tasks will understand how to find these new external users going forward

    if ($PSCmdlet.ShouldProcess("$userEmail",'New CMDB User'))
    {
        $username = $userEmail.Split("@")[0]
        $domainAndTLD = $userEmail.Split("@")[1]
        $domain = $domainAndTLD.Split(".")[0]
        $newID = $domain + "_" + $username + "_SMTP"

        #create the new user
        $newUser = New-SCSMObject -Class $domainUserClass -PropertyHashtable @{"domain" = "$domainAndTLD"; "username" = "$username"; "displayname" = "$userEmail"} @scsmMGMTParams -PassThru -NoCommit:$NoCommit
        if (($loggingLevel -ge 1) -and ($newUser)){New-SMEXCOEvent -Source "New-CMDBUser" -EventId 0 -LogMessage "New User created in SCSM. Username: $username" -Severity "Information"}

        #create the user notification projection
        $userNoticeProjection = @{__CLASS = "$($domainUserClass.Name)";
                                    __SEED = $newUser;
                                    Notification = @{__CLASS = "$($notificationClass)";
                                                        __OBJECT = @{"ID" = $newID; "TargetAddress" = "$userEmail"; "DisplayName" = "E-mail address"; "ChannelName" = "SMTP"}
                                                    }
                                    }

        #create the user's email notification channel
        $userHasNotification = New-SCSMObjectProjection -Type "$($userHasPrefProjection.Name)" -Projection $userNoticeProjection -PassThru -NoCommit:$NoCommit @scsmMGMTParams
        if (($loggingLevel -ge 1) -and ($userHasNotification)){New-SMEXCOEvent -Source "New-CMDBUser" -EventId 1 -LogMessage "New User: $username successfully related to their new Notification: $userEmail" -Severity "Information"}

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterUserCreatedInCMDB }

        return $newUser
    }
}

#combined previous 4 individual comment functions featured from versions 1 to 1.4.3 into single function and introduced more Action Log functionality
#inspired and modified from Travis Wright here - https://techcommunity.microsoft.com/t5/system-center-blog/creating-membership-and-hosting-objects-relationships-using-new/ba-p/347537
#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
#inspired and modified from Xapity here - http://www.xapity.com/single-post/2016/11/27/PowerShell-for-SCSM-Updating-the-Action-Log
function Add-ActionLogEntry {
    param (
        #The SCSM Work Item
        [parameter(Mandatory=$true, Position=0)]
        $WIObject,
        #The Action to invoke against the Action Log
        [parameter(Mandatory=$true, Position=1)]
        [ValidateSet("Assign","AnalystComment","Closed","Escalated","EmailSent","EndUserComment","FileAttached","FileDeleted","Reactivate","Resolved","TemplateApplied")]
        [string] $Action,
        #The string to write on the associate Action
        [parameter(Mandatory=$false, Position=2)]
        [string] $Comment,
        #The string of the user who performed the Action
        [parameter(Mandatory=$true, Position=3)]
        [string] $EnteredBy,
        #Whether or not the Comment is Private. This is only applicable when using the EndUserComment or AnalystComment Action
        [parameter(Mandatory=$false, Position=4)]
        [Nullable[boolean]] $IsPrivate = $false
    )

    #Choose the Action Log Entry to be created. Depending on the Action Log being used, the $propDescriptionComment Property could be either Comment or Description.
    switch ($Action)
    {
        Assign {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordAssigned"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        AnalystComment {$CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"; $propDescriptionComment = "Comment"}
        Closed {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordClosed"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        Escalated {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordEscalated"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        EmailSent {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.EmailSent"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        EndUserComment {$CommentClass = "System.WorkItem.TroubleTicket.UserCommentLog"; $propDescriptionComment = "Comment"}
        FileAttached {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.FileAttached"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        FileDeleted {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.FileDeleted"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        Reactivate {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordReopened"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        Resolved {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.RecordResolved"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
        TemplateApplied {$CommentClass = "System.WorkItem.TroubleTicket.ActionLog"; $ActionType = "System.WorkItem.ActionLogEnum.TemplateApplied"; $ActionEnum = get-scsmenumeration $ActionType @scsmMGMTParams; $propDescriptionComment = "Description"}
    }

    #Alias on Type Projection for Service Requests and Problem and are singular, whereas Incident and Change Request are plural. Update $CommentClassName
    if (($WIObject.ClassName -eq "System.WorkItem.Problem") -or ($WIObject.ClassName -eq "System.WorkItem.ServiceRequest")) {$CommentClassName = "ActionLog"} else {$CommentClassName = "ActionLogs"}

    #Analyst and End User Comments Classes have different Names based on the Work Item class
    if ($Action -eq "AnalystComment")
    {
        switch ($WIObject.ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = "AnalystComments"}
            "System.WorkItem.ServiceRequest" {$CommentClassName = "AnalystCommentLog"}
            "System.WorkItem.Problem" {$CommentClassName = "Comment"}
            "System.WorkItem.ChangeRequest" {$CommentClassName = "AnalystComments"}
        }
    }
    if ($Action -eq "EndUserComment")
    {
        switch ($WIObject.ClassName)
        {
            "System.WorkItem.Incident" {$CommentClassName = "UserComments"}
            "System.WorkItem.ServiceRequest" {$CommentClassName = "EndUserCommentLog"}
            "System.WorkItem.Problem" {$CommentClass = "System.WorkItem.TroubleTicket.AnalystCommentLog"; $CommentClassName = "Comment"}
            "System.WorkItem.ChangeRequest" {$CommentClassName = "UserComments"}
        }
    }

    # Generate a new GUID for the entry
    $NewGUID = ([guid]::NewGuid()).ToString()

    # Create the object projection with properties
    $Projection = @{__CLASS = "$($WIObject.ClassName)";
                    __SEED = $WIObject;
                    $CommentClassName = @{__CLASS = $CommentClass;
                                        __OBJECT = @{Id = $NewGUID;
                                            DisplayName = $NewGUID;
                                            ActionType = $ActionType;
                                            $propDescriptionComment = $Comment;
                                            Title = "$($ActionEnum.DisplayName)";
                                            EnteredBy  = $EnteredBy;
                                            EnteredDate = (Get-Date).ToUniversalTime();
                                            IsPrivate = $IsPrivate;
                                        }
                    }
    }

    #create the projection based on the work item class
    switch ($WIObject.ClassName)
    {
        "System.WorkItem.Incident" {New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection$" -Projection $Projection @scsmMGMTParams}
        "System.WorkItem.ServiceRequest" {New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection$" -Projection $Projection @scsmMGMTParams}
        "System.WorkItem.Problem" {New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType$" -Projection $Projection @scsmMGMTParams}
        "System.WorkItem.ChangeRequest" {New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel$" -Projection $Projection @scsmMGMTParams}
    }
}

#if using windows authentication, retrieve a Cireson Web API token
function Get-CiresonPortalAPIToken
{
    #determine if the connector is running as a workflow
    if ($scsmLFXConfigMP.GetRules() | Where-Object {($_.Name -eq "SMLets.Exchange.Connector.15d8b765a2f8b63ead14472f9b3c12f0")} | Select-Object Enabled -ExpandProperty Enabled)
    {
        $ciresonPortalCredentials = @{"username" = "$ciresonPortalRunAsUsername"; "password" = "$ciresonPortalRunAsPassword"; "languagecode" = "ENU" } | ConvertTo-Json
    }
    else
    {
        $ciresonPortalCredentials = @{"username" = "$ciresonPortalUsername"; "password" = "$ciresonPortalPassword"; "languagecode" = "ENU" } | ConvertTo-Json
    }

    #make the call to the portal to retrieve a token
    $ciresonTokenURL = $ciresonPortalServer+"api/V3/Authorization/GetToken"
    $ciresonAPIToken = Invoke-RestMethod -uri $ciresonTokenURL -Method post -Body $ciresonPortalCredentials
    $ciresonAPIToken = "Token" + " " + $ciresonAPIToken

    return $ciresonAPIToken
}

#retrieve a user from SCSM through the Cireson Web Portal API
function Get-CiresonPortalUser
{
    param (
        #The User to lookup by Username
        [parameter(Mandatory=$true, Position=0)]
        $username,
        #The User's Domain
        [parameter(Mandatory=$true, Position=1)]
        $domain
    )

    $isAuthUserAPIurl = "api/V3/User/IsUserAuthorized?userName=$username&domain=$domain"
    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $ciresonPortalUserObject = Invoke-RestMethod -Uri ($ciresonPortalServer+$isAuthUserAPIurl) -Method post -UseDefaultCredentials
    }
    else
    {
        $ciresonPortalUserObject = Invoke-RestMethod -Uri ($ciresonPortalServer+$isAuthUserAPIurl) -Method post -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }
    return $ciresonPortalUserObject
}

#retrieve a group from SCSM through the Cireson Web Portal API
function Get-CiresonPortalGroup
{
    param (
        #The Group to retrieve by their Email Address
        [parameter(Mandatory=$true, Position=0)]
        $groupEmail
    )

    $adGroup = Get-ADGroup @adParams -Filter "Mail -eq $groupEmail"

    if($ciresonPortalWindowsAuth)
    {
        #wanted to use a get groups style request, but "api/V3/User/GetConsoleGroups" feels costly instead of a search
        $cwpGroupResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+"api/V3/User/GetUserList?userFilter=$($adGroup.Name)&filterByAnalyst=false&groupsOnly=true&maxNumberOfResults=25") -UseDefaultCredentials
    }
    else
    {
        $cwpGroupResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+"api/V3/User/GetUserList?userFilter=$($adGroup.Name)&filterByAnalyst=false&groupsOnly=true&maxNumberOfResults=25") -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }
    $ciresonPortalGroup = $cwpGroupResponse | select-object @{Name='AccessGroupId'; Expression={$_.Id}}, name | Where-Object{$_.name -eq $($adGroup.Name)}
    return $ciresonPortalGroup
}

#retrieve all the announcements on the portal
function Get-CiresonPortalAnnouncement
{
    param (
        #The Language Code to use to search Announcements
        [parameter(Mandatory=$true, Position=0)]
        $languageCode
    )

    $allAnnouncementsURL = "api/V3/Announcement/GetAllAnnouncements?languageCode=$($languageCode)"
    if($ciresonPortalWindowsAuth)
    {
        $allCiresonPortalAnnouncements = Invoke-RestMethod -uri ($ciresonPortalServer+$allAnnouncementsURL) -UseDefaultCredentials
    }
    else
    {
        $allCiresonPortalAnnouncements = Invoke-RestMethod -uri ($ciresonPortalServer+$allAnnouncementsURL) -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }
    return $allCiresonPortalAnnouncements
}

#search for available Request Offerings based on content from a New Work Item and notify the Affected user via the Cireson Portal API
function Search-AvailableCiresonPortalOffering
{
    param (
        #The words to use to search for relevant Request Offerings in the Service Catalogue
        [parameter(Mandatory=$true, Position=0)]
        $searchQuery,
        #The user whose Service Catalog Offering should be used to search for available Request Offerings
        [parameter(Mandatory=$true, Position=1)]
        $ciresonPortalUser
    )

    $serviceCatalogAPIurl = "api/V3/ServiceCatalog/GetServiceCatalog?userId=$($ciresonPortalUser.id)&isScoped=$($ciresonPortalUser.Security.IsServiceCatalogScoped)"
    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $serviceCatalogResults = Invoke-RestMethod -Uri ($ciresonPortalServer+$serviceCatalogAPIurl) -Method get -UseDefaultCredentials
    }
    else
    {
        $serviceCatalogResults = Invoke-RestMethod -Uri ($ciresonPortalServer+$serviceCatalogAPIurl) -Method get -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }

    #### If the user has access to some Request Offerings, find which RO Titles/Description contain words from their original message ####
    if ($serviceCatalogResults)
    {
        #prepare the results by generating a URL array for the email
        $matchingRequestURLs = @()
        foreach ($serviceCatalogResult in $serviceCatalogResults)
        {
            $wordsMatched = ($searchQuery.Trim().Split() | Where-Object{($serviceCatalogResult.RequestOfferingTitle -match "\b$_\b") -or ($serviceCatalogResult.RequestOfferingDescription -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToRO)
            {
                $ciresonPortalRequestURL = "`"" + $ciresonPortalServer + "SC/ServiceCatalog/RequestOffering/" + $serviceCatalogResult.RequestOfferingId + "," + $serviceCatalogResult.ServiceOfferingId + "`""
                $RequestOfferingURL = "<a href=$ciresonPortalRequestURL/>$($serviceCatalogResult.RequestOfferingTitle)</a><br />"
                $requestOfferingSuggestion = [PSCustomObject] @{
                    RequestOfferingURL = $RequestOfferingURL
                    WordsMatched       = $wordsMatched
                }
                $matchingRequestURLs += $requestOfferingSuggestion
            }
        }
        $matchingRequestURLs = ($matchingRequestURLs | sort-object WordsMatched -Descending).RequestOfferingURL
        return $matchingRequestURLs
    }
}

#search the Cireson KB based on content from a New Work Item and notify the Affected User
function Search-CiresonKnowledgeBase
{
    param (
        #The words to use to search for relevant Knowledge Articles
        [parameter(Mandatory=$true, Position=0)]
        $searchQuery
    )

    $kbAPIurl = "api/V3/Article/FullTextSearch?&searchValue=$searchQuery"
    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $kbResults = Invoke-RestMethod -Uri ($ciresonPortalServer+$kbAPIurl) -UseDefaultCredentials
    }
    else
    {
        $kbResults = Invoke-RestMethod -Uri ($ciresonPortalServer+$kbAPIurl) -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }

    $kbResults =  $kbResults | Where-Object{$_.endusercontent -ne ""} | select-object articleid, title

    if ($kbResults)
    {
        $matchingKBURLs = @()
        foreach ($kbResult in $kbResults)
        {
            $wordsMatched = ($searchQuery.Trim().Split() | Where-Object{($kbResult.title -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToKA)
            {
                $KnowledgeArticleURL = "<a href=$ciresonPortalServer" + "KnowledgeBase/View/$($kbResult.articleid)#/>$($kbResult.title)</a><br />"
                $knowledgeSuggestion = [PSCustomObject] @{
                    KnowledgeArticleURL = $KnowledgeArticleURL
                    WordsMatched        = $wordsMatched
                }
                $matchingKBURLs += $knowledgeSuggestion
            }
        }
        $matchingKBURLs = ($matchingKBURLs | sort-object WordsMatched -Descending).KnowledgeArticleURL
        return $matchingKBURLs
    }
}

#retrieve Cireson Knowledge Base articles and/or Request Offerings, optionally leverage Azure Cognitive Services if enabled. Return results as an HTML formatted array of URLs
function Get-CiresonSuggestionURL
{
    [CmdletBinding()]
    [OutputType([Object[]])]
    Param
    (
        #If present, your respective Cireson Knowledge Base will be searched using text from the Work item
        [Parameter()]
        [switch]$SuggestKA,
        #If present, your respective Cireson Service Catalogue will be searched using text from the Work item
        [Parameter()]
        [switch]$SuggestRO,
        #If present, use Azure Cognitive Services to extract keywords from the Work Item Text, and use those keywords to search your respective Cireson Knowledge Base
        [Parameter()]
        [switch]$AzureKA,
        #If present, use Azure Cognitive Services to extract keywords from the Work Item Text, and use those keywords to search your respective Cireson Service Catalogue
        [Parameter()]
        [switch]$AzureRO,
        #The SCSM Work Item
        [Parameter()]
        [object]$WorkItem,
        #The Affected User of the SCSM Work Item
        [Parameter()]
        [object]$AffectedUser
    )

    #if Suggestions are both false, just exit this function call
    if (($SuggestKA -eq $false) -and ($SuggestRO -eq $false))
    {
        return $null
    }

    #check string length
    $wiTitle = if ($WorkItem.title.length -ge 1) {$WorkItem.title.trim()} else {$null}
    $wiDesc = if ($WorkItem.description.length -ge 1) {$WorkItem.description.trim()} else {$null}
    if (($null -eq $wiTitle) -and ($null -eq $wiDesc))
    {
        return $null
    }
    else
    {
        $searchQuery = ("$($wiTitle) $($wiDesc)").Trim()
    }

    #retrieve the cireson portal user
    $portalUser = Get-CiresonPortalUser -username $AffectedUser.UserName -domain $AffectedUser.Domain

    #Define the initial keyword hashtable to use against the Cireson Web API
    $searchQueriesHash = @{"AzureRO" = "$searchQuery"; "AzureKA" = "$searchQuery"}

    #if at least 1 ACS feature is being used, retrieve the keywords from ACS
    if ($AzureKA -or $AzureRO)
    {
        $acsKeywordsToSet = (Get-AzureEmailKeyword -messageToEvaluate "$searchQuery")
    }

    #update the hashtable to set the ACS Keywords on the relevant feature(s)
    foreach ($paramName in 'AzureKA', 'AzureRO')
    {
        if ($PSBoundParameters[$paramName])
        {
            $change = $searchQueriesHash.GetEnumerator() | where-object {$_.Name -eq $paramName}
            $change | foreach-object {$searchQueriesHash[$_.Key]="$acsKeywordsToSet"}
        }
    }

    #determine which Suggestion features will be used
    $isSuggestionFeatureUsed =
        foreach ($paramName in 'SuggestKA', 'SuggestRO')
        {
            if ($PSBoundParameters[$paramName]) {$paramName}
        }

    #call the Suggestion functions passing the search query (work item description/keywords) per the enabled features
    switch ($isSuggestionFeatureUsed)
    {
        "SuggestKA" {$kbURLs = Search-CiresonKnowledgeBase -searchQuery $($searchQueriesHash["AzureKA"]); if ($null -eq $kbURLs) {$kbURLs = ""}}
        "SuggestRO" {$requestURLs = Search-AvailableCiresonPortalOffering -searchQuery $($searchQueriesHash["AzureRO"]) -ciresonPortalUser $portalUser; if ($null -eq $requestURLs) {$requestURLs = ""}}
    }

    return $kbURLs, $requestURLs
}

#take suggestion URL arrays returned from Get-CiresonSuggestionURL, load custom HTML templates, and send results back out to the Affected User about their Work Item
function Send-CiresonSuggestionEmail
{
    [CmdletBinding()]
    Param
    (
        #The URL(s) that corrspond to Knowledge Base entries in your respective Cireson Knowledge Base
        [Parameter()]
        [array]$KnowledgeBaseURLs,
        #The URL(s) that corrspond to Request Offering entries in your respective Cireson Service Catalogue
        [Parameter()]
        [array]$RequestOfferingURLs,
        #The Work Item that might be Resolved/Completed via the suggested Knowledge Base or Request Offering URLs
        [Parameter()]
        [object]$WorkItem,
        #The Email Address of the User to send suggestions to
        [Parameter()]
        [string]$AffectedUserEmailAddress
    )

    switch ($WorkItem.ClassName)
    {
        "System.WorkItem.Incident" {$resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $WorkItem.id + "]" + "&body=This%20can%20be%20[$resolvedKeyword]" + "`">resolved</a>"}
        "System.WorkItem.ServiceRequest" {$resolveMailTo= "<a href=`"mailto:$workflowEmailAddress" + "?subject=" + "[" + $WorkItem.id + "]" + "&body=This%20can%20be%20[$cancelledKeyword]" + "`">cancel</a>"}
    }

    #determine which template to use
    if ($KnowledgeBaseURLs -and $RequestOfferingURLs) {$emailTemplate = get-content ("$htmlSuggestionTemplatePath" + "suggestKARO.html") -raw}
    if ($KnowledgeBaseURLs -and !$RequestOfferingURLs) {$emailTemplate = get-content ("$htmlSuggestionTemplatePath" + "suggestKA.html") -raw}
    if (!$KnowledgeBaseURLs -and $RequestOfferingURLs) {$emailTemplate = get-content ("$htmlSuggestionTemplatePath" + "suggestRO.html") -raw}

    #replace tokens in the template with URLs
    $emailTemplate = try {$emailTemplate.Replace("{0}", $KnowledgeBaseURLs)} catch {New-SMEXCOEvent -Source "Send-CiresonSuggestionEmail" -EventID 1 -Severity "Error" -LogMessage "The suggestion email's Knowledge Articles could not be populated. $($_.Exception).)"}
    $emailTemplate = try {$emailTemplate.Replace("{1}", $RequestOfferingURLs)} catch {New-SMEXCOEvent -Source "Send-CiresonSuggestionEmail" -EventID 2 -Severity "Error" -LogMessage "The suggestion email's Request Offerings could not be populated. $($_.Exception).)"}
    $emailTemplate = try {$emailTemplate.Replace("{2}", $resolveMailTo)} catch {New-SMEXCOEvent -Source "Send-CiresonSuggestionEmail" -EventID 3 -Severity "Error" -LogMessage "The suggestion email's Resolve/Completed text could not be populated. $($_.Exception).)"}

    #send the email to the affected user
    Send-EmailFromWorkflowAccount -subject "[$($WorkItem.id)] - $($WorkItem.title)" -body $emailTemplate -bodyType "HTML" -toRecipients $AffectedUserEmailAddress

    #if enabled, as part of the Suggested KA or RO process set the First Response Date on the Work Item
    if ($enableSetFirstResponseDateOnSuggestions) {Set-SCSMObject -SMObject $WorkItem -Property FirstResponseDate -Value (get-date).ToUniversalTime() @scsmMGMTParams}
}

function Add-CiresonWatchListUser
{
    param (
        #The GUID of the SCSM user
        [parameter(Mandatory=$true, Position=0)]
        $userguid,
        #The GUID of the SCSM Work item
        [parameter(Mandatory=$true, Position=1)]
        $workitemguid
    )

    $addToWatchlistAPIurl = "api/V3/WorkItem/AddToWatchlist?workitemId=$workitemguid&userId=$userguid"
    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $addToWatchlistResult = Invoke-RestMethod -Uri ($ciresonPortalServer+$addToWatchlistAPIurl) -UseDefaultCredentials -Method post
    }
    else
    {
        $addToWatchlistResult = Invoke-RestMethod -Uri ($ciresonPortalServer+$addToWatchlistAPIurl) -Method post -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }
    return $addToWatchlistResult
}
function Remove-CiresonWatchListUser
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The GUID of the SCSM user
        $userguid,
        #The GUID of the SCSM Work item
        $workitemguid
    )

    if($PSCmdlet.ShouldProcess($workitemguid, "Remove user $userguid from watchlist"))
    {
        $removeFromWatchlistAPIurl = "api/V3/WorkItem/DeleteFromWatchlist?workitemId=$workitemguid&userId=$userguid"
        if ($ciresonPortalWindowsAuth -eq $true)
        {
            $removeFromWatchlistResult = Invoke-RestMethod -Uri ($ciresonPortalServer+$removeFromWatchlistAPIurl) -UseDefaultCredentials -Method DELETE
        }
        else
        {
            $removeFromWatchlistResult = Invoke-RestMethod -Uri ($ciresonPortalServer+$removeFromWatchlistAPIurl) -Method DELETE -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
        }
        return $removeFromWatchlistResult
    }
}

#send an email from the SCSM Workflow Account
function Send-EmailFromWorkflowAccount
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The Subject of the Email to send
        [parameter(Mandatory=$true)]
        $subject,
        #The Body of the Email to send
        [parameter(Mandatory=$true)]
        $body,
        #The Body Type to use (Text or HTML) of the Email to send
        [parameter(Mandatory=$true)]
        [ValidateSet("Text", "HTML")]
        $bodyType,
        #Who the email will be sent to
        [parameter(Mandatory=$true)]
        $toRecipients
    )

    $emailToSendOut = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchangeService
    $emailToSendOut.Subject = $subject
    $emailToSendOut.Body = New-Object Microsoft.Exchange.WebServices.Data.MessageBody
    $emailToSendOut.Body.Text = $body
    $emailToSendOut.Body.BodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::$bodyType
    $emailToSendOut.ToRecipients.Add($toRecipients)
    $emailToSendOut.Send()

    if ($loggingLevel -ge 4)
    {
        $logMessage = "Email sent from Workflow account
        Subject: $subject
        Body: $body"
        New-SMEXCOEvent -Source "Send-EmailFromWorkflowAccount" -EventId 0 -LogMessage $logMessage -Severity "Information"
    }
}

function Set-WorkItemScheduledTime
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The calendar ppointment from Exchange distilled down to a PSCustomObject
        $calAppt,
        #The Work Item whose Scheduled Start/End Dates should be updated
        $workItem
    )

    if ($PSCmdlet.ShouldProcess("$workItem","Set Scheduled Start/End Time"))
    {
        try
        {
            switch ($calAppt.ItemClass)
            {
                "IPM.Schedule.Meeting.Request" {Set-SCSMObject -SMObject $workItem -propertyhashtable @{"ScheduledStartDate" = $calAppt.StartTime.ToUniversalTime(); "ScheduledEndDate" = $calAppt.EndTime.ToUniversalTime()} @scsmMGMTParams}
                "IPM.Schedule.Meeting.Canceled" {Set-SCSMObject -SMObject $workItem -propertyhashtable @{"ScheduledStartDate" = $null; "ScheduledEndDate" = $null} @scsmMGMTParams}
            }
        }
        catch
        {
            if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Set-WorkItemScheduledTime" -EventId 1 -LogMessage "Meeting $($calAppt.ItemClass.Split(".")[3]) for $($workItem.Name). Scheduled Start/End Times could not be updated." -Severity "Warning"}
        }

        if ($loggingLevel -ge 1)
        {
            New-SMEXCOEvent -Source "Set-WorkItemScheduledTime" -EventId 0 -LogMessage "Meeting $($calAppt.ItemClass.Split(".")[3]) for $($workItem.Name). Scheduled Start/End Times have been updated." -Severity "Information"
        }
    }
}

function Confirm-WorkItem
{
    param (
        #The Exchange message whose Conversation ID will be searched in SCSM
        [Parameter(Mandatory=$true)]
        $message,
        #Return the Work Item that is created from this function
        [Parameter(Mandatory=$false)]
        $returnWorkItem
    )

    #If emails are being attached to New Work Items, filter on the File Attachment Description that equals the Exchange Conversation ID as defined in the Add-EmailToSCSMObject function
    if ($attachEmailToWorkItem -eq $true)
    {
        $emailAttachmentSearchObject = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" @scsmMGMTParams | select-object -first 1
        if ($emailAttachmentSearchObject)
        {
            if ($loggingLevel -ge 4){New-SMEXCOEvent -Source "Confirm-WorkItem" -EventId 1 -LogMessage "File Attachment: $($emailAttachmentSearchObject.DisplayName) found" -Severity "Information"}
            $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $emailAttachmentSearchObject @scsmMGMTParams).sourceobject.id @scsmMGMTParams
            if ($relatedWorkItemFromAttachmentSearch)
            {
                if ($loggingLevel -ge 4){New-SMEXCOEvent -Source "Confirm-WorkItem" -EventId 2 -LogMessage "File Attachment: $($emailAttachmentSearchObject.DisplayName) has related Work Item $($relatedWorkItemFromAttachmentSearch.Name)" -Severity "Information"}
                switch ($relatedWorkItemFromAttachmentSearch.ClassName)
                {
                    "System.WorkItem.Incident" {Update-WorkItem -message $message -wiType "ir" -workItem $relatedWorkItemFromAttachmentSearch; if($returnWorkItem){return $relatedWorkItemFromAttachmentSearch}}
                    "System.WorkItem.ServiceRequest" {Update-WorkItem -message $message -wiType "sr" -workItem $relatedWorkItemFromAttachmentSearch; if($returnWorkItem){return $relatedWorkItemFromAttachmentSearch}}
                    "System.WorkItem.ChangeRequest" {Update-WorkItem -message $message -wiType "cr" -workItem $relatedWorkItemFromAttachmentSearch; if($returnWorkItem){return $relatedWorkItemFromAttachmentSearch}}
                }
            }
            else
            {
                if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Confirm-WorkItem" -EventId 3 -LogMessage "A File Attachment was found to merge this email with a known Work Item. But the Work Item could not be found." -Severity "Warning"}
                #the File Attachment (email) was found, but the related Work Item was not, Create a New Work Item
                New-WorkItem -message $message -wiType $defaultNewWorkItem
            }
        }
        else
        {
            if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Confirm-WorkItem" -EventId 4 -LogMessage "A File Attachment was not found to merge this email with a known Work Item" -Severity "Warning"}
            #the File Attachment (email) was not found, Create a New Work Item
            New-WorkItem -message $message -wiType $defaultNewWorkItem
        }
    }
    else
    {
        #will never engage as Confirm-WorkItem currently only works when attaching emails to work items
    }
}

function Test-EmailPattern
{
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'MessageClass', Justification = 'False positive')]
    param (
        #The Exchange message class, e.g. IPM.Note, IPM.Schedule.Meeting.Request, etc.
        [Parameter(Mandatory=$true)]
        $MessageClass,
        #The message from Exchange being evaluated
        [Parameter(Mandatory=$true)]
        $Email
    )

    #log that we've entered Custom Rules function
    if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 0 -Severity "Information" -LogMessage "No out of box Work Item match found. Attempting to reconcile against custom defined matching patterns."}

    #initialize variables, null the switch string/content block
    $searchBody = $false
    $searchSubject = $false
    $subjectSwitchBlockContent = [String]::Empty
    $bodySwitchBlockContent = [String]::Empty
    $smexcoSettingsCustomRulesMessageClass = $smexcoSettingsCustomRules | Where-Object {($_.CustomRuleMessageClassEnum.DisplayName -eq "$MessageClass")}
    $smexcoSettingsCustomRulesKnownMessageClass = $smexcoSettingsCustomRulesMessageClass | Where-Object {($_.CustomRuleItemType -eq "IR") -or ($_.CustomRuleItemType -eq "SR") -or ($_.CustomRuleItemType -eq "CR") -or ($_.CustomRuleItemType -eq "PR") -or ($_.CustomRuleItemType -eq "RR")}
    $smexcoSettingsCustomRulesUnknownMessageClass = $smexcoSettingsCustomRulesMessageClass | Where-Object {($_.CustomRuleItemType -ne "IR") -and ($_.CustomRuleItemType -ne "SR") -and ($_.CustomRuleItemType -ne "CR") -and ($_.CustomRuleItemType -ne "PR") -and ($_.CustomRuleItemType -ne "RR")}
    $workItemMatch = $null
    $customItemMatch = $null

    #evaluate Work Items first
    if ($smexcoSettingsCustomRulesKnownMessageClass.Count -ge 1)
    {
        #log that we're about to evaluate out of box Work Item conditions - IR, SR, CR, PR, RR
        if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 1 -Severity "Information" -LogMessage "Evaluating Work Item Patterns."}

        #{0} = IR, SR, PR, CR, RR
        #{1} = the Class Extension that was defined in the Settings MP to write the matched value to
        $wiActionStatement = '$result = get-scsmobject -class ${0}Class -filter "{1} -eq $($matches[0])"; if ($result) {$isUpdate = $true; update-workitem -message $email -wiType "{0}" -workItem $result; return $isUpdate} else {return $matches[0]}'

        #loop through and load custom patterns defined within the Settings MP to construct the dynamic switch statements
        foreach ($customWIPattern in $smexcoSettingsCustomRulesKnownMessageClass)
        {
            #build a switch statement just for the Email Subject
            if ($customWIPattern.CustomRuleMessagePart -eq "Subject")
            {
                $searchSubject = $true
                $subjectActionStatement = $wiActionStatement.Replace("{0}", $($customWIPattern.CustomRuleItemType))
                $subjectActionStatement = $subjectActionStatement.Replace("{1}", $($customWIPattern.CustomRuleRegexMatchProperty))
                $subjectSwitchBlockContent += '"' + $($customWIPattern.CustomRuleRegex) + '" { ' + $subjectActionStatement + " }`r`n"

                #log that we're evaluating Subject patterns
                if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 2 -Severity "Information" -LogMessage "Building Email Subject Switch Pattern: $subjectSwitchBlockContent"}
            }
            #build a switch statement just for the Email Body
            if ($customWIPattern.CustomRuleMessagePart -eq "Body")
            {
                $searchBody = $true
                $bodyActionStatement = $wiActionStatement.Replace("{0}", $($customWIPattern.CustomRuleItemType))
                $bodyActionStatement = $bodyActionStatement.Replace("{1}", $($customWIPattern.CustomRuleRegexMatchProperty))
                $bodySwitchBlockContent += '"' + $($customWIPattern.CustomRuleRegex) + '" { ' + $bodyActionStatement + " }`r`n"

                #log that we're evaluating Body patterns
                if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 3 -Severity "Information" -LogMessage "Building Email Body Switch Pattern: $bodySwitchBlockContent"}
            }
        }

        #### $switchResult + $wiActionStatment. $wiActionStatement either performed an update ($true) OR it returned the matched pattern. $switchResult drives further logic as its a BOOL or a STRING ####
        #attempt to turn the Subject switch string into an actual SWITCH statement, then execute it
        try
        {
            #convert the subject string
            if ($searchSubject)
            {
                #look for an Update to perform. If no update is found, return the matched pattern result
                $subjectSwitchBlock = [scriptblock]::Create($switchBlockTemplate + $subjectSwitchBlockContent + '}')
                Invoke-Command -ScriptBlock $subjectSwitchBlock -ArgumentList:@($email.subject) -OutVariable dynSubjSwitchResult
            }
            #$switchResult could either be $true, in which case an update was performed
            #$switchResult could be the matched pattern, in which case no update was performed
            #switchResult could be null, because no update was perform and no pattern was identified
            if ($dynSubjSwitchResult[0].Length -ge 1)
            {
                $switchResult = $dynSubjSwitchResult
                if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 4 -Severity "Information" -LogMessage "Email Subject Patten Search will either return true or the matched pattern: $($switchResult.ToString())."}
            }
        }
        catch
        {
            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 5 -Severity "Error" -LogMessage "Email Subject Switch Pattern failed. Examine Event ID 2."}
        }

        #attempt to turn the Body switch string into an actual SWITCH statement, then execute it
        try
        {
            #convert the body string
            if ($searchBody)
            {
                $bodySwitchBlock = [scriptblock]::Create($switchBlockTemplate + $bodySwitchBlockContent + '}')
                Invoke-Command -ScriptBlock $bodySwitchBlock -ArgumentList:@($email.body) -OutVariable dynBodySwitchResult
            }
            #$switchResult could either be $true, in which case an update was performed
            #$switchResult could be the matched pattern, in which case no update was performed
            #switchResult could be null, because no update was perform and no pattern was identified
            if ($dynBodySwitchResult[0].Length -ge 1)
            {
                $switchResult = $dynBodySwitchResult
                if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 6 -Severity "Information" -LogMessage "Email Body Patten Search will either return true or the matched pattern: $switchResult."}
            }
        }
        catch
        {
            if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 7 -Severity "Error" -LogMessage "Email Subject Switch Pattern failed. Examine Event ID 3."}
        }

        #switchResult was either $true OR it contains the custom pattern that was matched
        if (($switchResult -ne $true) -and ($null -ne $switchResult))
        {
            #Identify which pattern in the list of Custom Rule Types it belongs to create a New Work Item, then set the Custom Rule Class Extension property on it
            $customRuleMatchPattern = $smexcoSettingsCustomRules | Where-Object {$switchResult -match $_.CustomRuleRegex} | Select-Object -first 1

            #switchResult contains a known matched pattern
            if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 8 -Severity "Information" -LogMessage "The pattern matched a Custom Rule, but nothing was updated. Create a New Work Item of type $($customRuleMatchPattern.CustomRuleItemType). Write $switchResult into it's $($customRuleMatchPattern.CustomRuleRegexMatchProperty) property so subsequent updates can be performed against it."}

            $newWi = New-WorkItem -message $email -wiType $customRuleMatchPattern.CustomRuleItemType -returnWIBool $true
            Set-SCSMObject -SMObject $newWi -property $customRuleMatchPattern.CustomRuleRegexMatchProperty -value $switchResult[0]
        }

        #No Work Item was created or updated given the criteria
        if (($null -eq $switchResult)) {$workItemMatch = $false}
        else {$workItemMatch = $true}
    }
    #evaluate conditions that are not stock Work Items assuming there was no known Work Item action invoked
    if (($smexcoSettingsCustomRulesUnknownMessageClass.Count -ge 1) -and ($workItemMatch -eq $false))
    {
        #log that we're about to evaluate something that was not IR, SR, CR, PR, RR
        if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 9 -Severity "Information" -LogMessage "Evaluating Patterns for a custom defined Work/Config Item type."}

        #loop through and identify matched patterns before invoking Custom Events
        foreach ($customWIPattern in $smexcoSettingsCustomRulesUnknownMessageClass)
        {
            #match occured in the Subject and a configured custom pattern, invoke custom events
            if ($customWIPattern.CustomRuleMessagePart -eq "Subject")
            {
                #log that we're evaluating Subject patterns
                if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 2 -Severity "Information" -LogMessage "Evaluating Email Subject Switch Pattern: $($customWIPattern.CustomRuleRegex)"}

                $searchSubject = $true
                if (($ceScripts) -and ($Email.Subject -match $customWIPattern.CustomRuleRegex)) {Invoke-CustomRuleAction}
                else {$customItemMatch = $false; if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 10 -Severity "Warning" -LogMessage "Either Custom Events are not being used or the Subject did not match the $($customWIPattern.CustomRuleRegex) pattern."}}
            }
            #match occured in the Body and a configured custom pattern, invoke custom events
            if ($customWIPattern.CustomRuleMessagePart -eq "Body")
            {
                #log that we're evaluating Body patterns
                if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 3 -Severity "Information" -LogMessage "Evaluating Email Body Switch Pattern: $($customWIPattern.CustomRuleRegex)"}

                $searchBody = $true
                if (($ceScripts) -and ($Email.Body -match $customWIPattern.CustomRuleRegex)) {Invoke-CustomRuleAction}
                else {$customItemMatch = $false; if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 11 -Severity "Warning" -LogMessage "Either Custom Events are not being used or the Body did not match the $($customWIPattern.CustomRuleRegex) pattern."}}
            }
        }
    }
    #No stock Work Item was created/updated based on a Custom Pattern.
    #No custom Item was created/updated based on a Custom Pattern.
    #Create a Work Item e.g. "Default" of the above two switch statements
    if (($workItemMatch -eq $false) -and ($customItemMatch -eq $false))
    {
        if ($loggingLevel -ge 1) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 12 -Severity "Information" -LogMessage "No Work Item or Config Item was created or updated through Custom Rules and Custom Event's 'Invoke-CustomRuleAction' function. Create a Default Work Item."}
        New-WorkItem -message $email -wiType $defaultNewWorkItem -returnWIBool $false
    }
}

function Undo-WorkItemResolution
{
    param (
        #The SCSM Work Item (Incident or Problem) whose Resolution should be undone
        [Parameter(Mandatory=$true)]
        $WorkItem
    )

    switch ($WorkItem.ClassName)
    {
        "System.WorkItem.Incident" {$enumValue = "IncidentStatusenum.Active$"; $resName = "ResolutionCategory"}
        "System.WorkItem.Problem" {$enumValue = "ProblemStatusEnum.Active$"; $resName = "Resolution"}
    }
    Set-SCSMObject -SMObject $WorkItem -propertyhashtable @{"Status" = $enumValue; $resName = $null; "ResolutionDescription" = $null; "ResolvedDate" = $null} @scsmMGMTParams;
}

function Read-MIMEMessage
{
    param (
        #The Message to get the MIME content from
        [Parameter(Mandatory=$true)]
        $message
    )

    #Get the Mime Content of the message via MimeKit
    $messageWithMimeContent = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $mimeMessageMemoryStream = New-Object System.IO.MemoryStream($messageWithMimeContent.MimeContent.Content,0,$messageWithMimeContent.MimeContent.Content.Length)
    $parsedMimeMessage = New-Object MimeKit.MimeParser($mimeMessageMemoryStream)

    return $parsedMimeMessage
}

# Get-TemplatesByMailbox returns a hashtable with DefaultWiType, IRTemplate, SRTemplate, PRTemplate, and CRTemplate
function Get-TemplatesByMailbox
{
    param (
        #The Exchange message whose To should be evaluated TO field should be evaluated to determine which SCSM Template to apply
        [Parameter(Mandatory=$true)]
        $message
    )

    if ($loggingLevel -ge 1){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 1 -Severity "Information" -LogMessage "To: $($message.To)"}
    # There could be more than one addressee--loop through and match to our list
    foreach ($recipient in $message.To) {
        if ($loggingLevel -ge 1){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 2 -Severity "Information" -LogMessage "Recipient Address: $($recipient.Address)"}

        # Break on the first match
        if ($Mailboxes[$recipient.Address]) {
            $MailboxToUse = $recipient.Address
            break
        }
    }

    if ($MailboxToUse) {
        if ($loggingLevel -ge 1){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 3 -Severity "Information" -LogMessage "Redirection from known mailbox: $mailboxToUse.  Using custom templates."}
        return $Mailboxes[$MailboxToUse]
    }
    else {
        # check CC
        foreach ($recipient in $message.CC) {
            if ($recipient.Address) { $recipientAddress = $recipient.Address } else { $recipientAddress = $recipient }
            Write-Debug $recipientAddress

            # Break on the first match
            if ($Mailboxes[$recipientAddress]) {
                $MailboxToUse = $recipientAddress
                break
            }
        }

        if ($MailboxToUse) {
            if ($loggingLevel -ge 1){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 4 -Severity "Information" -LogMessage "Redirection from known mailbox: $mailboxToUse.  Found in CC field.  Using custom templates."}
            return $Mailboxes[$MailboxToUse]
        }
        else {
            # If not found in the To OR CC field, look in headers (BCC won't be readable)
            # Resent-From is the ideal field, but usually removed before the object is accessed.  Return-Path is a good second choice
            $HeaderSchema = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageHeaders)
            $msgWithHeaders = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.Id,$HeaderSchema)
            $ReturnPath = $msgWithHeaders.InternetMessageHeaders.Find("Return-Path").Value
            if ($Mailboxes[$ReturnPath]) {
                $MailboxToUse = $ReturnPath
            }

            if ($MailboxToUse) {
                if ($loggingLevel -ge 1){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 5 -Severity "Information" -LogMessage "Redirection from known mailbox: $mailboxToUse.  Found in Return-Path field.  Using custom templates."}
                return $Mailboxes[$MailboxToUse]
            }
            else {
                if ($loggingLevel -ge 2){New-SMEXCOEvent -Source "Get-TemplatesByMailbox" -EventID 6 -Severity "Warning" -LogMessage "No redirection from known mailbox.  Using Default templates"}
                return $Mailboxes[$workflowEmailAddress]
            }
        }
    }
}

function Get-SCSMWorkItemSetting {
    param (
        #A string of an SCSM Class Name
        [Parameter()]
        [string]$WorkItemClass
    )

    switch ($WorkItemClass) {
        "System.WorkItem.Incident" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.WorkItem.Incident.GeneralSetting$"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxAttachments
            $maxSize = $settings.MaxAttachmentSize
            $prefix = $settings.PrefixForId
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.ServiceRequest" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ServiceRequestSettings$"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.ServiceRequestPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.ChangeRequest" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ChangeSettings$"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.SystemWorkItemChangeRequestIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Problem" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ProblemSettings$"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.ProblemIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Release" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ReleaseSettings$"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.SystemWorkItemReleaseRecordIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Activity.ReviewActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.SystemWorkItemActivityReviewActivityIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Activity.ManualActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.SystemWorkItemActivityManualActivityIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Activity.ParallelActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.SystemWorkItemActivityParallelActivityIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Activity.SequentialActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.SystemWorkItemActivitySequentialActivityIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Activity.DependentActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.SystemWorkItemActivityDependentActivityIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "Microsoft.SystemCenter.Orchestrator.RunbookAutomationActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.MicrosoftSystemCenterOrchestratorRunbookAutomationActivityBaseIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Activity.SMARunbookActivity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.MicrosoftSystemCenterOrchestratorRunbookAutomationActivityBaseIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "Cireson.Powershell.Activity" {
            $ActivitySettingsObj = Get-SCSMObject -Class (Get-SCSMClass -Name "System.GlobalSetting.ActivitySettings$" @scsmMGMTParams) @scsmMGMTParams
            $prefix = $ActivitySettingsObj.SystemWorkItemActivityIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
    }

    return @{"MaxAttachments"=$maxAttach;"MaxAttachmentSize"=$maxSize;"Prefix"=$prefix;"PrefixRegex"=$prefixRegex}
}

# Test a message for the presence of certain key words
function Test-KeywordsFoundInMessage {
    param (
        #The Exchange message whose Subject + Body should be evaluated for defined keywords in the Settings UI
        [Parameter()]
        [string]$message
    )

    $found = $false
    #check the subject first
    $found = ($message.subject -match $workItemTypeOverrideKeywords)
    #if necessary, check the body
    if (-Not $found) {
        $found = ($message.body -match $workItemTypeOverrideKeywords)
    }

    if (($loggingLevel -ge 1) -and ($found)){New-SMEXCOEvent -Source "Test-KeywordsFoundInMessage" -EventId 0 -LogMessage "Override keywords found in email. Will create: $workItemOverrideType instead of: $defaultNewWorkItem" -Severity "Information"}
    return $found
}

#retrieve sender's ability to post announcement based on previously defined email addresses or an AD group
function Get-SCSMAuthorizedAnnouncer
{
    param (
        #The Email Address to verify permissions to post Announcements to SCSM
        [Parameter(Mandatory=$true)]
        [string]$EmailAddress
    )

    switch ($approvedMemberTypeForSCSMAnnouncer)
    {
        "users" {if ($approvedUsersForSCSMAnnouncements -match $EmailAddress)
                    {
                        return $true
                    }
                    else
                    {
                        return $false
                    }
                }
        "group" {$group = Get-ADGroup @adParams -Identity $approvedADGroupForSCSMAnnouncements
                    $adUser = Get-ADUser @adParams -Filter "EmailAddress -eq '$EmailAddress'"
                    if ($adUser)
                    {
                        if ((Get-ADUser @adParams -Identity $adUser -Properties memberof).memberof -eq $group.distinguishedname)
                        {
                            return $true
                        }
                        else
                        {
                            return $false
                        }
                    }
                }
    }
}

function Set-CoreSCSMAnnouncement
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The message to post in the annoucement
        $message,
        #The work item to cite in the announcement
        $workItem
    )

    if ($PSCmdlet.ShouldProcess("$workItem","Set Core SCSM Announcement $message"))
    {
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeSetCoreScsmAnnouncement }

        #if the message is an email, we need to add the end time property to the object
        #otherwise, it's a calendar appointment/meeting which already has these properties
        if ($message.ItemClass -eq "IPM.Note")
        {
            $message | Add-Member -type NoteProperty -name StartTime -Value $message.DateTimeReceived
            $message | Add-Member -type NoteProperty -name EndTime -Value $null
        }

        #Parse the message body for a Priority #keyword to correlate to a SCSM Priority
        #Rearrange the keyword from the title body before posting the announcement
        #If the end time is null, generate it
        $announcementTitle = $message.Subject -replace "\[$($workItem.Name)\]", ""
        $announcementTitle = $announcementTitle + " " + "[$($workItem.Name)]"
        if ($message.body -match [Regex]::Escape("#$lowAnnouncemnentPriorityKeyword"))
        {
            #low priority
            $scsmPriorityName = "Low"
            $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
            $announcementBody = $announcementBody -replace "\#$lowAnnouncemnentPriorityKeyword", ""
            if ($null -eq $message.EndTime) {$message.EndTime = $message.StartTime.AddHours($lowAnnouncemnentExpirationInHours)}
        }
        elseif ($message.body -match [Regex]::Escape("#$criticalAnnouncemnentPriorityKeyword"))
        {
            #high priority
            $scsmPriorityName = "Critical"
            $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
            $announcementBody = $announcementBody -replace "#$criticalAnnouncemnentPriorityKeyword", ""
            if ($null -eq $message.EndTime) {$message.EndTime = $message.StartTime.AddHours($criticalAnnouncemnentExpirationInHours)}
        }
        else
        {
            #normal priority
            $scsmPriorityName = "Medium"
            $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
            if ($null -eq $message.EndTime) {$message.EndTime = $message.StartTime.AddHours($normalAnnouncemnentExpirationInHours)}
        }

        $announcementClass = get-scsmclass -name "System.Announcement.Item$" @scsmMGMTParams
        $announcementPropertyHashtable = @{"Title" = $announcementTitle; "Body" = $announcementBody; "ExpirationDate" = $message.EndTime.ToUniversalTime(); "Priority" = $scsmPriorityName}

        #get any current announcement to update, otherwise create
        $currentSCSMAnnouncements = Get-SCSMObject -Class $announcementClass -Filter "Title -like '*$($workitem.Name)*'" @scsmMGMTParams

        if ($currentSCSMAnnouncements)
        {
            foreach ($currentSCSMAnnouncement in $currentSCSMAnnouncements)
            {
                Set-SCSMObject -SMObject $currentSCSMAnnouncement -PropertyHashtable $announcementPropertyHashtable @scsmMGMTParams
            }
        }
        else
        {
            #create the announcement in SCSM
            New-SCSMObject -Class $announcementClass -PropertyHashtable $announcementPropertyHashtable @scsmMGMTParams
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterSetCoreScsmAnnouncement }
    }
}

function Set-CiresonPortalAnnouncement
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The message to post in the annoucement
        $message,
        #The work item to cite in the announcement
        $workItem
    )

    if ($PSCmdlet.ShouldProcess("$workItem","Set Cireson Portal Announcement $message"))
    {
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeSetPortalAnnouncement }

        $updateAnnouncementURL = "api/v3/Announcement/UpdateAnnouncement"

        #if the message is an email, we need to add the start time and end time property to the object
        #otherwise, it's a calendar appointment/meeting which already has these properties
        if ($message.ItemClass -eq "IPM.Note")
        {
            $message | Add-Member -type NoteProperty -name StartTime -Value $message.DateTimeReceived
            $message | Add-Member -type NoteProperty -name EndTime -Value $null
        }

        #Parse the message body for a Priority #keyword to correlate to a Cireson Priority enum
        #Remove the keyword from the title body before posting the announcement
        #If the end time is null, generate it
        $announcementTitle = $message.Subject -replace "\[$($workItem.Name)\]", ""
        $announcementTitle = $announcementTitle + " " + "[$($workItem.Name)]"
        if ($message.body -match [Regex]::Escape("#$lowAnnouncemnentPriorityKeyword"))
        {
            #low priority
            $ciresonPortalPriorityEnum = "F860661B-D9D6-41CB-A501-467B4DD81A7B"
            $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
            $announcementBody = $announcementBody -replace "\#$lowAnnouncemnentPriorityKeyword", ""
            if ($null -eq $message.EndTime) {$message.EndTime = $message.StartTime.AddHours($lowAnnouncemnentExpirationInHours)}
        }
        elseif ($message.body -match [Regex]::Escape("#$criticalAnnouncemnentPriorityKeyword"))
        {
            #high priority
            $ciresonPortalPriorityEnum = "F10A51C2-C569-4E64-8237-2B117D63DDB8"
            $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
            $announcementBody = $announcementBody -replace "#$criticalAnnouncemnentPriorityKeyword", ""
            if ($null -eq $message.EndTime) {$message.EndTime = $message.StartTime.AddHours($criticalAnnouncemnentExpirationInHours)}
        }
        else
        {
            #normal priority
            $ciresonPortalPriorityEnum = "64096F7F-F8E0-491C-A7FE-94FEDDED4715"
            $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
            if ($null -eq $message.EndTime) {$message.EndTime = $message.StartTime.AddHours($normalAnnouncemnentExpirationInHours)}
        }

        #Extract the groups that the message was sent to
        #rename the GroupID property to "AccessGroupID" so as to compare the difference later
        $groupEmails = @()
        $groupEmails += $message.To | Where-Object{$_.MailboxType -ne "Mailbox"}
        $groupEmails += $message.Cc | Where-Object{$_.MailboxType -ne "Mailbox"}
        $portalGroups = @()
        foreach ($groupEmail in $groupEmails)
        {
            $portalGroups += Get-CiresonPortalGroup -groupEmail $groupEmail.Name
        }

        #Get the user that is posting the announcement (from) from the SCSM/Cireson Portal to determine their language code to post the announcement
        $announcerSCSMObject = Get-SCSMUserByEmailAddress -EmailAddress "$($message.from)"
        $ciresonPortalAnnouncer = Get-CiresonPortalUser -username $announcerSCSMObject.username -domain $announcerSCSMObject.domain

        #Get any announcements that already exist for the Work Item
        $allPortalAnnouncements = Get-CiresonPortalAnnouncement -languageCode $ciresonPortalAnnouncer.LanguageCode
        $allPortalAnnouncements = $allPortalAnnouncements | Where-Object{$_.title -match "\[" + $workitem.name + "\]"}

        #determine authentication to use (windows/forms)
        if ($allPortalAnnouncements)
        {
            #### there are announcements to create/update ####

            #### combine the announcement objects and group objects together and group by GroupAccessID, then find object groups that don't have an announcement id ####
            #Announcement array has an AccessGroupID that does not match a group from the message. create announcement for that group
            $groupsToCreateAnnouncements = ($portalGroups + $allPortalAnnouncements) | Group-Object -Property AccessGroupId | Where-Object{$_.Count -eq 1} | Select-Object -Expand Group | Where-Object{$null -eq $_.Id}

            #Announcement array has an AccessGroupID that contains a current group from the message.
            $groupsToUpdateAnnouncements = ($portalGroups + $allPortalAnnouncements) | Group-Object -Property AccessGroupId | Select-Object -Expand Group | Where-Object{$null -ne $_.Id}

            # create announcement for new group
            foreach ($groupsToCreateAnnouncement in $groupsToCreateAnnouncements)
            {
                #create the Portal Announcement Hashtable, convert to JSON object, then POST to create
                $announcement = @{"Id" = [guid]::NewGuid();
                                    "Title" = $announcementTitle;
                                    "Body" = $announcementBody;
                                    "Priority" = @{"Id"=$ciresonPortalPriorityEnum};
                                    "AccessGroupId" = @{"Id"=$($groupsToCreateAnnouncement.AccessGroupId)};
                                    "StartDate" = $message.StartTime.ToUniversalTime();
                                    "EndDate" = $message.EndTime.ToUniversalTime();
                                    "Locale" = $($ciresonPortalAnnouncer.LanguageCode)}
                $announcement = $announcement | ConvertTo-Json

                #post the announcement
                if ($ciresonPortalWindowsAuth -eq $true)
                {
                    Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials | Out-Null
                }
                else
                {
                    Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -Headers @{"Authorization"=Get-CiresonPortalAPIToken}  | Out-Null
                }
            }

            # update current announcement's title and body
            foreach ($groupsToUpdateAnnouncement in $groupsToUpdateAnnouncements)
            {
                #create the Portal Announcement Hashtable using current Announcement ID, convert to JSON object, then POST to update
                $announcement = @{"Id" = $groupsToUpdateAnnouncement.Id;
                                    "Title" = $announcementTitle;
                                    "Body" = $announcementBody;
                                    "Priority" = @{"Id"=$ciresonPortalPriorityEnum};
                                    "AccessGroupId" = @{"Id"=$($groupsToUpdateAnnouncement.AccessGroupId)};
                                    "StartDate" = $message.StartTime.ToUniversalTime();
                                    "EndDate" = $message.EndTime.ToUniversalTime();
                                    "Locale" = $($ciresonPortalAnnouncer.LanguageCode)}
                $announcement = $announcement | ConvertTo-Json

                #post the announcement
                if ($ciresonPortalWindowsAuth -eq $true)
                {
                    Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials | Out-Null
                }
                else
                {
                    Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -Headers @{"Authorization"=Get-CiresonPortalAPIToken} | Out-Null
                }
            }
        }
        else
        {
            #### there are announcements to create ####

            #Cireson Portal Announcements can only target a single group. Create an announcement for each group
            foreach ($portalGroup in $portalGroups)
            {
                #create the Portal Announcement Hashtable, convert to JSON object, then POST
                $announcement = @{"Id" = [guid]::NewGuid();
                                    "Title" = $announcementTitle;
                                    "Body" = $announcementBody;
                                    "Priority" = @{"Id"=$ciresonPortalPriorityEnum};
                                    "AccessGroupId" = @{"Id"=$($portalGroup.AccessGroupId)};
                                    "StartDate" = $message.StartTime.ToUniversalTime();
                                    "EndDate" = $message.EndTime.ToUniversalTime();
                                    "Locale" = $($ciresonPortalAnnouncer.LanguageCode)}
                $announcement = $announcement | ConvertTo-Json

                #post the announcement
                if ($ciresonPortalWindowsAuth)
                {
                    Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials | Out-Null
                }
                else
                {
                    Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -Headers @{"Authorization"=Get-CiresonPortalAPIToken} | Out-Null
                }
            }
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterSetPortalAnnouncement }
    }
}

#modified from Ritesh Modi - https://blogs.msdn.microsoft.com/riteshmodi/2017/03/24/text-analytics-key-phrase-cognitive-services-powershell/
function Get-AzureEmailSentiment
{
    param (
        #The text to evaluate for text sentiment
        [Parameter(Mandatory=$true)]
        $messageToEvaluate
    )

    #define cognitive services URLs
    $sentimentURI = "https://$azureRegion.api.cognitive.microsoft.$azureTLD/text/analytics/v2.0/sentiment"

    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "en"; "id" = "1"; "text" = "$messageToEvaluate" };
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final

    try
    {
        #invoke the Cognitive Services Sentiment API
        $sentimentResult = Invoke-RestMethod -Method Post -Uri $sentimentURI -Header @{ "Ocp-Apim-Subscription-Key" = $azureCogSvcTextAnalyticsAPIKey } -Body $messagePayload -ContentType "application/json"

        #API contacted, log an info event
        if ($loggingLevel -ge 4)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailSentiment" -EventID 0 -Severity "Information" -LogMessage "Azure Sentiment Score: $($sentimentResult.documents.score * 100)"
        }
    }
    catch
    {
        #the API could not be contacted, log an error
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailSentiment" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #return the percent score
    return ($sentimentResult.documents.score * 100)
}

#determine the language being used before converting it
function Get-AzureEmailLanguage
{
    param (
        #The text to evaluate to determine the language of
        [Parameter(Mandatory=$true)]
        $TextToEvaluate
    )

    #build the request
    $translationServiceURI = "https://api.cognitive.microsofttranslator.$azureTLD/detect?api-version=3.0"
    $RecoRequestHeader = @{
      'Ocp-Apim-Subscription-Key' = "$azureCogSvcTranslateAPIKey";
      'Content-Type' = "application/json; charset=utf-8"
    }

    #prepare the body of the request
    $TextToEvaluate = @{'Text' = $($TextToEvaluate)} | ConvertTo-Json
    $originalBytes = [Text.Encoding]::Default.GetBytes($TextToEvaluate)
    $TextToEvaluate = [Text.Encoding]::Utf8.GetString($originalBytes)

    try
    {
        #Send text to Azure for translation
        $RecoResponse = Invoke-RestMethod -Method POST -Uri $translationServiceURI -Headers $RecoRequestHeader -Body "[$($TextToEvaluate)]"
        if ($loggingLevel -ge 4) {
            $azureEmailIdentifiedLang = "Language Identified: $($RecoResponse.translations[0].text)"
            New-SMEXCOEvent -Source "Get-AzureEmailLanguage" -EventID 0 -Severity "Information" -LogMessage $azureEmailIdentifiedLang
        }
    }
    catch
    {
        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Get-AzureEmailLanguage" -EventID 1 -Severity "Error" -LogMessage $_.Exception}
    }

    #Return the language with the highest match score
    return $RecoResponse | sort-object score | select-object -first 1
}

#translate the language
function Get-AzureEmailTranslation
{
    param (
        #The text translate from one language to another
        [Parameter(Mandatory=$true)]
        $TextToTranslate,
        #The source language to translate from
        [Parameter(Mandatory=$true)]
        $SourceLanguage,
        #The target language to translate into
        [Parameter(Mandatory=$true)]
        $TargetLanguage
    )

    #build the request
    $translationServiceURI = "https://api.cognitive.microsofttranslator.$azureTLD/translate?api-version=3.0&from=$($SourceLanguage)&to=$($TargetLanguage)"
    $RecoRequestHeader = @{
      'Ocp-Apim-Subscription-Key' = "$azureCogSvcTranslateAPIKey";
      'Content-Type' = "application/json; charset=utf-8"
    }

    #prepare the body of the request
    $TextToTranslate = @{'Text' = $($TextToTranslate)} | ConvertTo-Json
    $originalBytes = [Text.Encoding]::Default.GetBytes($TextToTranslate)
    $TextToTranslate = [Text.Encoding]::Utf8.GetString($originalBytes)

    try
    {
        #Send text to Azure for translation
        $RecoResponse = Invoke-RestMethod -Method POST -Uri $translationServiceURI -Headers $RecoRequestHeader -Body "[$($TextToTranslate)]"

        #API contacted, log an info event
        if ($loggingLevel -ge 4)
        {
            $azureEmailTranslationLogBody = "Source Lang: $SourceLanguage
                Target Lang: $TargetLanguage
                Original Text: $TextToTranslate
                Translated Text: $RecoResponse.translations[0].text
            "
            New-SMEXCOEvent -Source "Get-AzureEmailTranslation" -EventID 0 -Severity "Information" -LogMessage $azureEmailTranslationLogBody
        }
    }
    catch
    {
        #API could not be contacted, log an error
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailTranslation" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #Return the converted text
    return $($RecoResponse.translations[0].text)
}

function Get-ACSWorkItemPriority
{
    param (
        #The confidence score from Azure
        [Parameter(Mandatory=$true)]
        $score,
        #The SCSM Work Item Class as a string
        [Parameter(Mandatory=$true)]
        $wiClass
    )

    if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventId 0 -Severity "Information" -LogMessage "Fetching Boundary for $wiClass with a score of $($score.ToString())"}
    $wiClass = $wiClass.Replace("System.WorkItem.Incident", "IR").Replace("System.WorkItem.ServiceRequest", "SR")
    [xml]$acsXMLBoundaries = $smexcoSettingsMP.ACSPriorityScoringBoundaries
    $priorityCalc = $acsXMLBoundaries.ACSPriorityBoundaries.ACSPriorityBoundary | foreach-object {if (($score -ge $_.Min) -and ($score -le $_.Max) -and ($wiClass -eq $_.WorkItemType)) {$_.IRImpactSRUrgencyEnum, $_.IRUrgencySRPriorityEnum}}
    $priorityCalc = $priorityCalc | foreach-object {Get-SCSMEnumeration -id $_ | select-object name -ExpandProperty name } | foreach-object {$_ + "$"}

    if ($loggingLevel -ge 4)
    {
        switch ($wiClass)
        {
            "IR" {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventId 1 -Severity "Information" -LogMessage "Incident Impact: $($priorityCalc[0]). Urgency: $($priorityCalc[1])"}
            "SR" {New-SMEXCOEvent -Source "Get-ACSWorkItemPriority" -EventId 1 -Severity "Information" -LogMessage "Service Request Urgency: $($priorityCalc[0]). Priority: $($priorityCalc[1])"}
        }
    }

    return $priorityCalc
}

function Get-AzureEmailKeyword
{
    param (
        #The text to evaluate for keyword extraction
        [Parameter(Mandatory=$true)]
        $messageToEvaluate
    )

    #define cognitive services URLs
    $keyPhraseURI = "https://$azureRegion.api.cognitive.microsoft.$azureTLD/text/analytics/v2.0/keyPhrases"

    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "en"; "id" = "1"; "text" = "$messageToEvaluate" };
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final

    try
    {
        #invoke the Text Analytics Keyword API
        $keywordResult = Invoke-RestMethod -Method Post -Uri $keyPhraseURI -Header @{ "Ocp-Apim-Subscription-Key" = $azureCogSvcTextAnalyticsAPIKey } -Body $messagePayload -ContentType "application/json"

        #API contacted, log an info event
        if ($loggingLevel -ge 4)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailKeyword" -EventID 0 -Severity "Information" -LogMessage "Keywords identified from Azure: $($keywordResult.documents.keyPhrases)"
        }
    }
    catch
    {
        #API could not be contacted, log an error
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailKeyword" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #return the keywords
    return $keywordResult.documents.keyPhrases
}
function Get-AzureEmailImageAnalysis
{
    param (
        #The file attachment (image) to analyze to summarize with tags (text)
        [Parameter(Mandatory=$true)]
        $imageToEvalute
    )

    #azure cognitive services, vision URL
    $imageAnalysisURI = "https://$azureVisionRegion.api.cognitive.microsoft.$azureTLD/vision/v3.0/analyze?visualFeatures=Tags"

    #adapted from C# per: https://docs.microsoft.com/en-us/azure/cognitive-services/computer-vision/quickstarts/csharp-print-text
    Add-Type -AssemblyName "System.Net.Http"
    $httpClient = New-Object -TypeName "System.Net.Http.Httpclient"
    $httpClient.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "$azureCogSvcVisionAPIKey")
    $content = New-Object "System.Net.Http.ByteArrayContent" -ArgumentList @(,$imageToEvalute)
    $content.Headers.ContentType = "application/octet-stream"

    try
    {
        $request = $httpClient.PostAsync($imageAnalysisURI,$content)
        $request.wait();
        if($request.IsCompleted) {$result = $request.Result.Content.ReadAsStringAsync().Result | ConvertFrom-Json}

        #API contacted, log an info event
        if ($loggingLevel -ge 4)
        {
            $logAzureVisionTags = $azureVisionResult.tags.name | select-object -first 5
            $emailImageTagsLogEventBody = "Identified Image Tags from Azure: $($logAzureVisionTags -join ',')"
            New-SMEXCOEvent -Source "Get-AzureEmailImageAnalysis" -EventID 0 -Severity "Information" -LogMessage $emailImageTagsLogEventBody
        }
    }
    catch
    {
        #the API could not be contacted, log an error
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailImageAnalysis" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #return the Vision API analysis
    return $result
}

function Get-AzureSpeechEmailAudioText
{
    param (
        #The wave/file attachment (audio) to perform speech to text against
        [Parameter(Mandatory=$true)]
        $waveFileToEvaluate
    )

    #build the request
    $SpeechServiceURI = "https://$azureSpeechRegion.stt.speech.microsoft.$azureTLD/speech/recognition/conversation/cognitiveservices/v1?language=en-us"
    $RecoRequestHeader = @{
      'Ocp-Apim-Subscription-Key' = "$azureCogSvcSpeechAPIKey";
      'Content-type' = "audio/wav; codecs=audio/pcm; samplerate=16000";
      'Transfer-Encoding' = 'chunked'
      'Except' = "100-continue"
      'Accept' = "application/json"
    }

    try
    {
        #Pass the audio byte array into the body and submit the request
        $RecoResponse = Invoke-RestMethod -Method POST -Uri $SpeechServiceURI -Headers $RecoRequestHeader -Body $waveFileToEvaluate

        #API contacted, log an info event
        if ($loggingLevel -ge 4)
        {
            $emailSpeechToTextLogEventBody = "Azure Speech to Text: $($RecoResponse.DisplayText)"
            New-SMEXCOEvent -Source "Get-AzureSpeechEmailAudioText" -EventID 0 -Severity "Information" -LogMessage $emailSpeechToTextLogEventBody
        }
    }
    catch
    {
        #the API could not be contacted, log an error
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AzureSpeechEmailAudioText" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #return the result
    return $RecoResponse
}
function Get-AzureEmailImageText
{
    param (
        #The image/file attachment (image) to perform OCR against
        [Parameter(Mandatory=$true)]
        $imageToEvalute
    )


    #azure cognitive services, vision URL
    $imageTextURI = "https://$azureVisionRegion.api.cognitive.microsoft.$azureTLD/vision/v3.0/ocr?detectOrientation=true"

    #adapted from C# per: https://docs.microsoft.com/en-us/azure/cognitive-services/computer-vision/quickstarts/csharp-print-text
    Add-Type -AssemblyName "System.Net.Http"
    $httpClient = New-Object -TypeName "System.Net.Http.Httpclient"
    $httpClient.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", "$azureCogSvcVisionAPIKey")
    $content = New-Object "System.Net.Http.ByteArrayContent" -ArgumentList @(,$imageToEvalute)
    $content.Headers.ContentType = "application/octet-stream"

    try
    {
        $request = $httpClient.PostAsync($imageTextURI,$content)
        $request.wait();
        if($request.IsCompleted) {$result = $request.Result.Content.ReadAsStringAsync().Result | ConvertFrom-Json}

        #API contacted, log an info event
        if ($loggingLevel -ge 4)
        {
            $emailImageTextLogEventBody = "Identified Image Text from Azure: $($result.regions.Lines.words.text -join " ")"
            New-SMEXCOEvent -Source "Get-AzureEmailImageText" -EventID 0 -Severity "Information" -LogMessage $emailImageTextLogEventBody
        }
    }
    catch
    {
        #the API could not be contacted, log an error
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AzureEmailImageText" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #return the Vision API analysis
    return $result
}
#endregion

function Get-AMLWorkItemProbability
{
    param (
        #The email subject to submit to Azure Machine Learning
        [Parameter(Mandatory=$true)]
        $EmailSubject,
        #The email body to submit to Azure Machine Learning
        [Parameter(Mandatory=$true)]
        $EmailBody
    )

    #create the header
    $headerTable = @{"Authorization" = "Bearer $amlAPIKey"; "Content-Type" = "application/json"}

    #create the JSON request
    $messagePayload = @"
    {
        "Inputs": {
            "Input1" : {
                "ColumnNames": ["Email_Subject", "Email_Description"],
                "Values": [
                    ["$EmailSubject", "$EmailBody"]
                ]
            }
        },
        "GlobalParameters": {}
    }
"@

    try
    {
        #invoke the Azure Machine Learning web service for predicting Work Item Type, Classification, and Support Group
        $probabilityResponse = Invoke-RestMethod -Uri $amlURL -Method Post -Header $headerTable -Body $messagePayload -ContentType "application/json"

        #return custom probability object
        $probabilityResults = $probabilityResponse.Results.output1.value.Values[0]
        $probabilityMatrix = [PSCustomObject]@{
            WorkItemType                     = $probabilityResults[0]
            WorkItemTypeConfidence           = (($probabilityResults[1] -as [decimal]) * 100)
            WorkItemClassification           = $probabilityResults[2]
            WorkItemClassificationConfidence = (($probabilityResults[3] -as [decimal]) * 100)
            WorkItemSupportGroup             = $probabilityResults[4]
            WorkItemSupportGroupConfidence   = (($probabilityResults[5] -as [decimal]) * 100)
            AffectedConfigItem               = $probabilityResults[6]
            AffectedConfigItemConfidence     = (($probabilityResults[7] -as [decimal]) * 100)
        }

        #logging is verbose, record the entire AML prediction
        if ($loggingLevel -ge 4)
        {
            $amlLogInfoLogEventBody = "AML predictions: $EmailSubject
                WorkItemType: $($probabilityMatrix.WorkItemType)
                WorkItemTypeConfidence: $($probabilityMatrix.WorkItemTypeConfidence)
                WorkItemClassification: $($probabilityMatrix.WorkItemClassification)
                WorkItemClassificationConfidence: $($probabilityMatrix.WorkItemClassificationConfidence)
                WorkItemSupportGroup: $($probabilityMatrix.WorkItemSupportGroup)
                WorkItemSupportGroupConfidence: $($probabilityMatrix.WorkItemSupportGroupConfidence)
                AffectedConfigItem: $($probabilityMatrix.AffectedConfigItem)
                AffectedConfigItemConfidence: $($probabilityMatrix.AffectedConfigItemConfidence)
            "
            New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 0 -Severity "Information" -LogMessage $amlLogInfoLogEventBody
        }
    }
    catch
    {
        #the Web Service couldn't be contacted
        if ($loggingLevel -ge 3)
        {
            New-SMEXCOEvent -Source "Get-AMLWorkItemProbability" -EventID 1 -Severity "Error" -LogMessage $_.Exception
        }
    }

    #return the percent score
    return ($probabilityMatrix)
}

#region #### Modified version of Set-SCSMTemplateWithActivities from Morton Meisler seen here http://blog.ctglobalservices.com/service-manager-scsm/mme/set-scsmtemplatewithactivities-powershell-script/
function Update-SCSMPropertyCollection
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateObject]$Object =$(throw "Please provide a valid template object"),
        $Alias
    )

    if($PSCmdlet.ShouldProcess("$($Object.DisplayName)"))
    {
        #Regex - Find class from template object property between ! and ']
        $pattern = '(?<=!)[^!]+?(?=''\])'
        if (($Object.Path -match $pattern) -and (($Matches[0].StartsWith("System.WorkItem.Activity")) -or ($Matches[0].StartsWith("Microsoft.SystemCenter.Orchestrator")) -or ($Matches[0].StartsWith("Cireson.Powershell.Activity"))))
        {
            #Set prefix from activity class
            $prefix = (Get-SCSMWorkItemSetting -WorkItemClass $Matches[0])["Prefix"]

            #Create template property object
            $propClass = [Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateProperty]
            $propObject = New-Object $propClass

            #Add new item to property object
            $propObject.Path = "`$Context/Property[Type='$alias!System.WorkItem']/Id$"
            $propObject.MixedValue = "$prefix{0}"

            #Add property to template
            $Object.PropertyCollection.Add($propObject)

            #recursively update activities in activities
            if ($Object.ObjectCollection.Count -ne 0)
            {
                foreach ($obj in $Object.ObjectCollection)
                {
                    Update-SCSMPropertyCollection -Object $obj -Alias $alias
                }
            }
        }
    }
}

function Set-SCSMTemplate
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param (
        [Microsoft.EnterpriseManagement.Common.EnterpriseManagementObjectProjection]$Projection =$(throw "Please provide a valid projection object"),
        [Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplate]$Template = $(throw 'Please provide an template object, ex. -template template')
    )
    if($PSCmdlet.ShouldProcess("$Projection","$Template"))
    {
        #Get alias from system.workitem.library managementpack to set id property
        $templateMP = $Template.GetManagementPack()
        $alias = $templateMP.References.GetAlias((Get-SCSMManagementPack "System.WorkItem.Library$" @scsmMGMTParams))

        #Update Activities in template
        foreach ($TemplateObject in $Template.ObjectCollection)
        {
            Update-SCSMPropertyCollection -Object $TemplateObject -Alias $alias
        }
        #Apply update template
        Set-SCSMObjectTemplate -Projection $Projection -Template $Template -ErrorAction Stop @scsmMGMTParams
    }
}
#endregion

function Remove-PII
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    param (
        #The body of the email
        $body
    )

    if($PSCmdlet.ShouldProcess("email body"))
    {
        <#Regexes
        Visa = 4[0-9]{15}
        Mastercard = (?:5[1-5][0-9]{2}|222[1-9]|22[3-9][0-9]|2[3-6][0-9]{2}|27[01][0-9]|2720)[0-9]{12}
        Discover = 6(?:011|5[0-9]{2})[0-9]{12}
        American Express = 3[47][0-9]{13}
        SSN/ITIN = \d{3}-?\d{2}-?\d{4}
        #>
        $piiRegexes = Get-Content $piiRegexPath
        ForEach ($regex in $piiRegexes)
        {
            switch -Regex ($body)
            {
                "$regex" {$body = $body -replace "$regex", "[redacted]"}
            }
        }
        return $body
    }
}

#region #### SCOM Request Functions ####
function Get-SCOMAuthorizedRequester
{
    param (
        #The email address to verify is an authorized requester of SCOM Distributed Application Health
        [Parameter(Mandatory=$true)]
        $EmailAddress
    )

    switch ($approvedMemberTypeForSCOM)
    {
        "users" {if ($approvedUsersForSCOM -match $EmailAddress)
                    {
                        return $true
                    }
                    else
                    {
                        return $false
                    }
                }
        "group" {$group = Get-ADGroup @adParams -Identity "$approvedADGroupForSCOM"
                    $adUser = Get-ADUser @adParams -Filter "EmailAddress -eq '$EmailAddress'"
                    if ($adUser)
                    {
                        if ((Get-ADUser @adParams -Identity $adUser -Properties memberof).memberof -eq $group.distinguishedname)
                        {
                            return $true
                        }
                        else
                        {
                            return $false
                        }
                    }
                }
    }
}

function Get-SCOMDistributedAppHealth
{
    param (
        #The email/text to use to search SCOM to look up Distributed Applications and their Health state
        [Parameter(Mandatory=$true)]
        $message
    )

    #determine if the sender is authorized to make SCOM Health requests
    $isAuthorized = Get-SCOMAuthorizedRequester -EmailAddress $message.From

    if (($isAuthorized -eq $true))
    {
        #find the distributed application to search for based on the [Distributed App Name] from the email body
        #"\[(.*?)\]" - will match something [Service Manager] or [Operations Manager Management Group]
        if ($message.body -match "\[(.*?)\]"){$appName = $Matches[0].Replace("[", "").Replace("]", "")}
        else {<#body not [formed] correctly#>}

        #get Distributed Applications that meet search criteria
        $distributedApps = Invoke-Command -ScriptBlock {(Get-SCOMClass -Name "System.Service" | Get-SCOMMonitoringObject | Where-Object {$_.displayname -like "*$using:appName*"}) | select-object Displayname, healthstate} -ComputerName $scomMGMTServer
        $healthySCOMApps = @()
        $unhealthySCOMApps = @()
        $notMonitoredSCOMApps = @()
        $unhealthySCOMAppsAlerts = @()
        $emailBody = @()

        #create, define, and load custom PS Object from SCOM DA Objects
        foreach ($distributedApp in $distributedApps)
        {
            switch ($distributedApp.HealthState.Value)
            {
                "Success" {$scomDAObject = [PSCustomObject] @{Name = $distributedApp.displayname; Status = "Healthy"}; $healthySCOMApps += $scomDAObject}
                "Error" {$scomDAObject = [PSCustomObject] @{Name = $distributedApp.displayname; Status = "Critical"}; $unhealthySCOMApps += $scomDAObject}
                "Uninitialized" {$scomDAObject = [PSCustomObject] @{Name = $distributedApp.displayname; Status = "Not Monitored"}; $notMonitoredSCOMApps += $scomDAObject}
            }
            $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
        }

        #if there are unhealthy apps/red agent states, get their Active alerts in SCOM
        if ($unhealthySCOMApps)
        {
            foreach ($unhealthySCOMApp in $unhealthySCOMApps)
            {
                $unhealthyAppName = $unhealthySCOMApp.Name
                $unhealthySCOMAppsAlerts += Invoke-Command -scriptblock {Get-SCOMClass -DisplayName "$using:unhealthyAppName" | Get-SCOMClassInstance | Foreach-Object {$_.GetRelatedMonitoringObjects('Recursive')} | Get-SCOMAlert -ResolutionState 0} -computername $scomMGMTServer
            }
        }

        $emailBody = $emailBody + "<br /><br />" + "NOTE: Responding to this message will trigger the creation of a New Work Item in Service Manager!"
        Send-EmailFromWorkflowAccount -subject "SCOM Health Status" -body $emailBody -bodyType "HTML" -toRecipients $message.From
    }
    else
    {
        return $false
    }
}
#endregion

#determine/enforce merge logic in the event this was omitted in configuration
if (($mergeReplies -eq $true) -or ($processCalendarAppointment -eq $true))
{
    $attachEmailToWorkItem = $true
}

#load the MimeKit assembly and decrypting context/certificate
if (($processDigitallySignedMessages -eq $true) -or ($processEncryptedMessages -eq $true))
{
    try
    {
        [void][System.Reflection.Assembly]::LoadFile($mimeKitDLLPath)
        [void][System.Reflection.Assembly]::LoadFile($($mimeKitDLLPath -ireplace [regex]::Escape("MimeKit.dll"), "BouncyCastle.Crypto.dll"))
        if ($certStore -eq "user")
        {
            $certStore = New-Object MimeKit.Cryptography.WindowsSecureMimeContext("CurrentUser")
        }
        else
        {
            $certStore = New-Object MimeKit.Cryptography.WindowsSecureMimeContext("LocalMachine")
        }
        if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Cryptography" -EventID 0 -Severity "Information" -LogMessage "Email certificate loaded."}
    }
    catch
    {
        #decrypting certificate or mimekit couldn't be loaded. Don't process signed/encrypted emails
        $processDigitallySignedMessages = $false
        $processEncryptedMessages = $false
        if ($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Cryptography" -EventID 1 -Severity "Error" -LogMessage $_.Exception}
    }
}

#load the Work Item regex patterns before the email processing loop
$irRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Incident")["PrefixRegex"]
$srRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.ServiceRequest")["PrefixRegex"]
$prRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Problem")["PrefixRegex"]
$crRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.ChangeRequest")["PrefixRegex"]
$maRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Activity.ManualActivity")["PrefixRegex"]
$paRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Activity.ParallelActivity")["PrefixRegex"]
$saRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Activity.SequentialActivity")["PrefixRegex"]
$daRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Activity.DependentActivity")["PrefixRegex"]
$raRegex = (Get-SCSMWorkItemSetting -WorkItemClass "System.Workitem.Activity.ReviewActivity")["PrefixRegex"]

#custom rules
$UseCustomRules = $smexcoSettingsMP.UseCustomRules

# Custom Event Handler
if ($ceScripts) { Invoke-BeforeConnect }
#define Exchange assembly and connect to EWS
[void] [Reflection.Assembly]::LoadFile("$exchangeEWSAPIPath")
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

#figure out if the workflow should be used
if ($scsmLFXConfigMP.GetRules() | Where-Object {($_.Name -eq "SMLets.Exchange.Connector.15d8b765a2f8b63ead14472f9b3c12f0")} | Select-Object Enabled -ExpandProperty Enabled)
{
    #the workflow exists and it is enabled, determine how to connect to Exchange
    if ($UseExchangeOnline)
    {
        #validate the Run As Account format to ensure it is an email address
        if (!(($ewsUsername + "@" + $ewsDomain) -match "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$"))
        {
            New-SMEXCOEvent -Source "General" -EventId 4 -LogMessage "The address/SCSM Run As Account used to sign into 365 is not a valid email address and is currently entered as $($ewsUsername + "@" + $ewsDomain). This will prevent a successful connection. To fix this, go to the Run As account in SCSM and for the username enter it as an email address like user@domain.tld" -Severity "Error"
        }
        #request an access token from Azure
        $ReqTokenBody = @{
            Grant_Type    = "Password"
            client_Id     = $AzureClientID
            Username      = $ewsUsername + "@" + $ewsDomain
            Password      = $ewspassword
            Scope         = $azureScopeURL
        }
        try{
            $response = Invoke-RestMethod -Uri $azureTokenURL -Method "POST" -Body $ReqTokenBody

            #instead of a username/password, use the OAuth access_token as the means to authenticate to Exchange
            $exchangeService.Url = [System.Uri]$ExchangeEndpoint
            $exchangeService.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]($response.Access_Token)

            if ($loggingLevel -ge 4){
                New-SMEXCOEvent -Source "General" -EventID 7 -LogMessage "Successfully retrieved an OAuth token from 365" -Severity "Information"
            }
        }
        catch{
            #couldn't retrieve the OAuth token
            if ($loggingLevel -ge 3)
            {
                New-SMEXCOEvent -Source "General" -EventId 8 -LogMessage "Could not retrieve OAuth token from 365: $($_.Exception)`nUsername: $($ReqTokenBody.Username)`nClient ID: $($ReqTokenBody.client_Id)" -Severity "Error"
            }
        }
    }
    else
    {
        #local exchange server
        $exchangeService.Credentials = New-Object Net.NetworkCredential($ewsusername, $ewspassword, $ewsdomain)
        if ($UseAutoDiscover -eq $true)
        {
            $exchangeService.AutodiscoverUrl($workflowEmailAddress)
        }
        else
        {
            $exchangeService.Url = [System.Uri]$ExchangeEndpoint
        }
    }
    #impersonation is being used with the workflow
    if ($smexcoSettingsMP.UseImpersonation)
    {
        $exchangeService.ImpersonatedUserId = [Microsoft.Exchange.WebServices.Data.ImpersonatedUserId]::new([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $workflowEmailAddress)
    }
}
else
{
    #the workflow either doesn't exist or it's not enabled, determine how to connect to Exchange
    if ($UseExchangeOnline)
    {
        #validate the Run As Account format to ensure it is an email address
        if (!(($username + "@" + $domain) -match "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$"))
        {
            New-SMEXCOEvent -Source "General" -EventId 4 -LogMessage "The address/SCSM Run As Account used to sign into 365 is not a valid email address and is currently entered as $($username + "@" + $password). This will prevent a successful connection. To fix this, go to the Run As account in SCSM and for the username enter it as an email address like user@domain.tld" -Severity "Error"
        }
        #request an access token from Azure
        $ReqTokenBody = @{
            Grant_Type    = "Password"
            client_Id     = $AzureClientID
            Username      = $username
            Password      = $password
            Scope         = $azureScopeURL
        }
        try{
            $response = Invoke-RestMethod -Uri $azureTokenURL -Method "POST" -Body $ReqTokenBody

            #instead of a username/password, use the OAuth access_token as the means to authenticate to Exchange
            $exchangeService.Url = [System.Uri]$ExchangeEndpoint
            $exchangeService.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]($response.Access_Token)

            if ($loggingLevel -ge 4){
                New-SMEXCOEvent -Source "General" -EventID 7 -LogMessage "Successfully retrieved an OAuth token from 365" -Severity "Information"
            }
        }
        catch{
            #couldn't retrieve the OAuth token
            if ($loggingLevel -ge 3)
            {
                New-SMEXCOEvent -Source "General" -EventId 8 -LogMessage "Could not retrieve OAuth token from 365: $($_.Exception)`nUsername: $($ReqTokenBody.Username)`nClient ID: $($ReqTokenBody.client_Id)" -Severity "Error"
            }
        }
    }
    else
    {
        #local exchange server
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
    }
}

#define search parameters and search on the defined classes
$inboxFolderName = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
#authenticate to Exchange
try
{
    $inboxFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchangeService,$inboxFolderName)
    #the authentication bind to Exchange service and Inbox folder worked, log an information event
    if ($loggingLevel -ge 4)
    {
        New-SMEXCOEvent -Source "General" -EventId 0 -LogMessage "Successfully connected to Exchange" -Severity "Information"
    }
}
catch
{
    #couldn't retrieve the Inbox, log an error and exit the connector
    if ($loggingLevel -ge 3)
    {
        New-SMEXCOEvent -Source "General" -EventId 1 -LogMessage $_.Exception -Severity "Error"
    }
    break
}
#define search parameters, search on the defined classes and get messages that are older than the current time
$itemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1000
$propertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$propertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text
$mimeContentSchema = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.ItemSchema]::MimeContent)
$dateTimeItem = [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
$now = get-date
$searchFilter = New-Object -TypeName Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo -ArgumentList $dateTimeItem,$now

#build the Where-Object scriptblock based on defined configuration
#by default the connector will ALWAYS process regular emails as seen in the $emailFilterString variable
$emailFilterString = '($_.ItemClass -eq "IPM.Note")'
$calendarFilterString = '($_.ItemClass -eq "IPM.Schedule.Meeting.Request") -or ($_.ItemClass -eq "IPM.Schedule.Meeting.Canceled")'
$digitallySignedFilterString = '($_.ItemClass -eq "IPM.Note.SMIME.MultipartSigned")'
$encryptedFilterString = '($_.ItemClass -eq "IPM.Note.SMIME")'
$unreadFilterString = '($_.isRead -eq $false)'
$inboxFilterString = @()
if ($processCalendarAppointment -eq $true)
{
    $inboxFilterString += $calendarFilterString
}
if ($processDigitallySignedMessages -eq $true)
{
    $inboxFilterString += $digitallySignedFilterString
}
if ($processEncryptedMessages -eq $true)
{
    $inboxFilterString += $encryptedFilterString
}
if ($UseCustomRules)
{
    #retrieve any custom rule patterns that are not the supported out of box enums
    $customMessageClasses = $smexcoSettingsCustomRules | Where-Object {$_.CustomRuleMessageClassEnum.Name -notlike "SMLets.Exchange.Connector.MessageClassEnum.*"}
    if ($customMessageClasses.count -eq 1)
    {
        $inboxFilterString += "(`$_.ItemClass -eq '$($smexcoSettingsExternalTicket.CustomRuleMessageClassEnum.DisplayName)')"
    }
    elseif ($customMessageClasses.count -ge 2)
    {
        foreach ($smexcoSettingsExternalTicket in $customMessageClasses)
        {
            $inboxFilterString += "(`$_.ItemClass -eq '$($smexcoSettingsExternalTicket.CustomRuleMessageClassEnum.DisplayName)')"
        }
    }
}

#finalize the where-object string by ensuring to look for all Unread Items
$inboxFilterString = $inboxFilterString -join ' -or '
if ($inboxFilterString.length -eq 0)
{
    $inboxFilterString = "(" + $emailFilterString + ")" + " -and " + $unreadFilterString
}
else
{
    $inboxFilterString = "(" + $inboxFilterString + " -or " + $emailFilterString + ")" + " -and " + $unreadFilterString
}
if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "General" -EventId 5 -LogMessage "Filtering Mailbox on: $inboxFilterString" -Severity "Information"}
$inboxFilterString = [scriptblock]::Create("$inboxFilterString")

#filter the inbox
$inbox = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView) | where-object $inboxFilterString | Sort-Object DateTimeReceived
if (($loggingLevel -ge 1)){New-SMEXCOEvent -Source "General" -EventId 2 -LogMessage "Messages to Process: $($inbox.Count)" -Severity "Information"; $messagesProcessed = 0}
# Custom Event Handler
if ($ceScripts) { Invoke-OnOpenInbox }

#build the formatting of a SWITCH statement using a Here-String
#!!!! This MUST maintain its formatting/indentation. DO NOT CHANGE !!!!
$switchBlockTemplate = @'
param($incomingValue)
switch -Regex ($incomingValue)
{

'@

#parse each message
foreach ($message in $inbox)
{
    #load the entire message
    $message.Load($propertySet)

    #initialize a variable to determine if valid update
    $isUpdate = $null

    #Process an Email
    if ($message.ItemClass -eq "IPM.Note")
    {
        $email = [PSCustomObject] @{
            From                = $message.From.Address
            To                  = $message.ToRecipients
            CC                  = $message.CcRecipients
            Subject             = $message.Subject
            Attachments         = $message.Attachments
            Body                = $message.Body.Text
            DateTimeSent        = $message.DateTimeSent
            DateTimeReceived    = $message.DateTimeReceived
            ID                  = $message.ID
            ConversationID      = $message.ConversationID
            ConversationTopic   = $message.ConversationTopic
            ItemClass           = $message.ItemClass
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessEmail }

        switch -Regex ($email.subject)
        {
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ir" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
            "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "sr" -workItem $result}else {new-workitem -message $email -wiType $defaultNewWorkItem}}
            "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "pr" -workItem $result}else {new-workitem -message $email -wiType $defaultNewWorkItem}}
            "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "cr" -workItem $result}else {new-workitem -message $email -wiType $defaultNewWorkItem}}

            #### activities ####
            "\[$raRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $raClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ra" -workItem $result}}
            "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ma" -workItem $result}}

            #### 3rd party classes, work items, etc. add here ####
            "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem -message $email -wiType $defaultNewWorkItem}}}

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "($replyRegex)(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if(!($isUpdate)){if($mergeReplies -eq $true){Confirm-WorkItem -message $email}else{new-workitem -message $email -wiType $defaultNewWorkItem}}}

            #### default action, create work item ####
            default {
                if (($smexcoSettingsCustomRules.CustomRuleMessageClassEnum.DisplayName -contains "IPM.Note") -and ($UseCustomRules -eq $true))
                {
                    #Custom Patterns has at least one defined use of IPM.Note, and Custom Rules are enabled
                    Test-EmailPattern -MessageClass $message.ItemClass -Email $email
                }
                else
                {
                    #No core SCSM Work Item was matched, Custom Rules may not have been used/have matched, create a New Work Item
                    New-WorkItem $email $defaultNewWorkItem
                }
            }
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessEmail }

        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve) | Out-Null
        if ($deleteAfterProcessing){$message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems) | Out-Null}
    }

    #### Process a Digitally Signed message ####
    elseif ($message.ItemClass -eq "IPM.Note.SMIME.MultipartSigned")
    {
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessSignedEmail }

        $response = Read-MIMEMessage -message $message

        $email = [PSCustomObject] @{
            From              = $response.From.address
            To                = $response.To
            CC                = $response.Cc
            Subject           = $response.Subject
            Attachments       = ($response.Attachments | Where-Object { $_.filename -ne "smime.p7s" })
            Body              = $response.TextBody.Trim()
            DateTimeSent      = $message.DateTimeSent
            DateTimeReceived  = $message.DateTimeReceived
            ID                = $message.ID
            ConversationID    = $message.ConversationID
            ConversationTopic = $message.ConversationTopic
            ItemClass         = $message.ItemClass
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessEmail }

        #Verify digital signature
        foreach ($sig in $response.Body.Verify($certStore))
        {
            try
            {
                $sigResult = $sig.Verify()
                if ($sigResult -eq $true)
                {
                    $validSig = $true
                    if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Cryptography" -EventID 2 -Severity "Information" -LogMessage "Digital signature is valid"}
                }
            }
            catch
            {
                $validSig = $false
                if ($loggingLevel -ge 2) {New-SMEXCOEvent -Source "Cryptography" -EventID 3 -Severity "Warning" -LogMessage "Digital signature could not be verified"}
            }
        }

        #The signature is valid OR signature is not valid and invalid signatures are set to process anyway
        if (($validSig) -or (($validSig -eq $false) -and ($ignoreInvalidDigitalSignature -eq $true)))
        {
            switch -Regex ($email.subject)
            {
                #### primary work item types ####
                "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ir" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "sr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "pr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "cr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}

                #### activities ####
                "\[$raRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $raClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ra" -workItem $result}}
                "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ma" -workItem $result}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "($replyRegex)(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if(!($isUpdate)){if($mergeReplies -eq $true){Confirm-WorkItem -message $email}else{new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### Email is going to invoke a custom action. The signature MUST be valid to proceed
                "\[$powershellKeyword]" {if ($validSig -and $ceScripts)
                {
                    Invoke-ValidDigitalSignatureAction
                    #you could then insert custom pwsh to parse email, sender, body, etc. and then
                    #restart a computer, bounce a windows service, ping a network device, call a 3rd party Rest API,
                    #or initiate a webhook for a 3rd party platform
                }}

                #### default action, create work item ####
                default {
                    if (($smexcoSettingsCustomRules.CustomRuleMessageClassEnum.DisplayName -contains "IPM.Note.SMIME.MultipartSigned") -and ($UseCustomRules -eq $true))
                    {
                        #Custom Patterns has at least one defined use of IPM.Note.SMIME.MultipartSigned, and Custom Rules are enabled
                        Test-EmailPattern -MessageClass $message.ItemClass -Email $email
                    }
                    else
                    {
                        #No core SCSM Work Item was matched, Custom Rules may not have been used/have matched, create a New Work Item
                        New-WorkItem $email $defaultNewWorkItem
                    }
                }
            }
        }
        #the signature is not valid and invalid signatures should not be processed, call custom actions if enabled
        else
        {
            #Custom Event Handler
            if ($ceScripts) { Invoke-InvalidDigitalSignatureAction }
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessEmail }

        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve) | Out-Null
        if ($deleteAfterProcessing){$message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems) | Out-Null}

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessSignedEmail }
    }

    #### Process an Encrypted message ####
    elseif ($message.ItemClass -eq "IPM.Note.SMIME")
    {
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }

        $response = Read-MIMEMessage -message $message
        try {$decryptedBody = $response.Body.Decrypt($certStore)} catch {if($loggingLevel -ge 3) {New-SMEXCOEvent -Source "Cryptography" -EventID 4 -Severity "Error" -LogMessage $_.Exception}}

        #Messaged is encrypted and/or signed
        if ($decryptedBody.ContentType.MimeType -eq "multipart/alternative")
        {
            #check to see if there are attachments
            $decryptedAttachments = $decryptedBody | Where-Object{$_.isattachment -eq $true}

            $email = [PSCustomObject] @{
                From              = $response.From.address
                To                = $response.To
                CC                = $response.Cc
                Subject           = $response.Subject
                Attachments       = $decryptedAttachments
                Body              = $decryptedBody.GetTextBody("Text")
                DateTimeSent      = $message.DateTimeSent
                DateTimeReceived  = $message.DateTimeReceived
                ID                = $message.ID
                ConversationID    = $message.ConversationID
                ConversationTopic = $message.ConversationTopic
                ItemClass         = $message.ItemClass
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEmail }

            switch -Regex ($email.subject)
            {
                #### primary work item types ####
                "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ir" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "sr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "pr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "cr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}

                #### activities ####
                "\[$raRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $raClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ra" -workItem $result}}
                "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ma" -workItem $result}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "($replyRegex)(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if(!($isUpdate)){if($mergeReplies -eq $true){Confirm-WorkItem -message $email}else{new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### default action, create work item ####
                default {
                    if (($smexcoSettingsCustomRules.CustomRuleMessageClassEnum.DisplayName -contains "IPM.Note.SMIME") -and ($UseCustomRules -eq $true))
                    {
                        #No Primary Work Item was matched, Custom Patterns has at least one defined use of IPM.Note.SMIME, and Custom Rules are enabled
                        Test-EmailPattern -MessageClass $message.ItemClass -Email $email
                    }
                    else
                    {
                        #No core SCSM Work Item was matched, Custom Rules may not have been used/have matched, create a New Work Item
                        New-WorkItem $email $defaultNewWorkItem
                    }
                }
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEmail }

            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve) | Out-Null
            if ($deleteAfterProcessing){$message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems) | Out-Null}

            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }
        }
        #Message is encrypted and/or signed, has at least 1 attachment that is an Exchange object
        if ($decryptedBody.ContentType.MimeType -eq "multipart/mixed")
        {
            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }

            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessSignedEmail }

            $decryptedAttachments = $decryptedBody | Where-Object{$_.isattachment -eq $true}

            $email = [PSCustomObject] @{
                From              = $response.From.address
                To                = $response.To
                CC                = $response.Cc
                Subject           = $response.Subject
                Attachments       = $decryptedAttachments
                Body              = $decryptedBody[0].GetTextBody("Text")
                DateTimeSent      = $message.DateTimeSent
                DateTimeReceived  = $message.DateTimeReceived
                ID                = $message.ID
                ConversationID    = $message.ConversationID
                ConversationTopic = $message.ConversationTopic
                ItemClass         = $message.ItemClass
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEmail }

            switch -Regex ($email.subject)
            {
                #### primary work item types ####
                "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ir" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "sr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "pr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "cr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}

                #### activities ####
                "\[$raRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $raClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ra" -workItem $result}}
                "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ma" -workItem $result}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "($replyRegex)(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if(!($isUpdate)){if($mergeReplies -eq $true){Confirm-WorkItem -message $email}else{new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### default action, create work item ####
                default {
                    if (($smexcoSettingsCustomRules.CustomRuleMessageClassEnum.DisplayName -contains "IPM.Note.SMIME") -and ($UseCustomRules -eq $true))
                    {
                        #No Primary Work Item was matched, Custom Patterns has at least one defined use of IPM.Note.SMIME, and Custom Rules are enabled
                        Test-EmailPattern -MessageClass $message.ItemClass -Email $email
                    }
                    else
                    {
                        #No core SCSM Work Item was matched, Custom Rules may not have been used/have matched, create a New Work Item
                        New-WorkItem $email $defaultNewWorkItem
                    }
                }
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEmail }

            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve) | Out-Null
            if ($deleteAfterProcessing){$message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems) | Out-Null}

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessSignedEmail }

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEncryptedEmail }
        }
        #Messaged is encrypted and/or signed
        if ($decryptedBody.ContentType.MimeType -eq "application/x-pkcs7-mime")
        {
            #check to see if there are attachments
            $digitalSignatures = $decryptedBody.Verify($certStore, [ref]$decryptedBody)
            $digitalSignatures | Foreach-Object {
                if ($_.Verify()) {
                    if ($loggingLevel -ge 4)
                    { New-SMEXCOEvent -Source "Cryptography" -EventID 2 -Severity "Information" -LogMessage "Digital signature is valid" }
                }
                else {
                    if ($loggingLevel -ge 2) {
                        New-SMEXCOEvent -Source "Cryptography" -EventID 3 -Severity "Warning" -LogMessage "Digital signature could not be verified"
                    }
                }
            }
            $decryptedBodyWOAttachments = $decryptedBody | Where-Object{($_.isattachment -eq $false)}
            $decryptedAttachments = if ($decryptedBody.ContentType.MimeType -eq "multipart/alternative") {$decryptedBody | Where-Object{$_.isattachment -eq $true}} else {$decryptedBody | Select-Object -skip 1}
            $digitalSignatures | Foreach-Object {if ($_.Verify()){New-SMEXCOEvent -Source "Cryptography" -EventID 2 -Severity "Information" -LogMessage "Digital signature is valid"} else {New-SMEXCOEvent -Source "Cryptography" -EventID 3 -Severity "Warning" -LogMessage "Digital signature could not be verified"}}
            $email = [PSCustomObject] @{
                From              = $response.From.address
                To                = $response.To
                CC                = $response.Cc
                Subject           = $response.Subject
                Attachments       = $decryptedAttachments
                Body              = $(try { $decryptedBodyWOAttachments.GetTextBody("Text").Trim() } catch { if ($decryptedBodyWOAttachments[0].Text.GetType().Name -eq "Object[]"){ $decryptedBodyWOAttachments[0][0].Text.Trim()} else{ $decryptedBodyWOAttachments[0].Text.Trim()} })
                DateTimeSent      = $message.DateTimeSent
                DateTimeReceived  = $message.DateTimeReceived
                ID                = $message.ID
                ConversationID    = $message.ConversationID
                ConversationTopic = $message.ConversationTopic
                ItemClass         = $message.ItemClass
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEmail }

            switch -Regex ($email.subject)
            {
                #### primary work item types ####
                "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ir" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "sr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "pr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}
                "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "cr" -workItem $result} else {new-workitem -message $email -wiType $defaultNewWorkItem}}

                #### activities ####
                "\[$raRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $raClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ra" -workItem $result}}
                "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){$isUpdate = $true; update-workitem -message $email -wiType "ma" -workItem $result}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "($replyRegex)(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if(!($isUpdate)){if($mergeReplies -eq $true){Confirm-WorkItem -message $email}else{new-workitem -message $email -wiType $defaultNewWorkItem}}}

                #### default action, create work item ####
                default {
                    if (($smexcoSettingsCustomRules.CustomRuleMessageClassEnum.DisplayName -contains "IPM.Note.SMIME") -and ($UseCustomRules -eq $true))
                    {
                        #No Primary Work Item was matched, Custom Patterns has at least one defined use of IPM.Note.SMIME, and Custom Rules are enabled
                        Test-EmailPattern -MessageClass $message.ItemClass -Email $email
                    }
                    else
                    {
                        #No core SCSM Work Item was matched, Custom Rules may not have been used/have matched, create a New Work Item
                        New-WorkItem $email $defaultNewWorkItem
                    }
                }
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEmail }

            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve) | Out-Null
            if ($deleteAfterProcessing){$message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems) | Out-Null}

            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }
        }
    }

    #Process a Calendar Meeting
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Request")
    {
        $appointment = [PSCustomObject] @{
            StartTime         = $message.Start
            EndTime           = $message.End
            To                = $message.ToRecipients
            From              = $message.From.Address
            Attachments       = $message.Attachments
            Subject           = $message.Subject
            DateTimeReceived  = $message.DateTimeReceived
            DateTimeSent      = $message.DateTimeSent
            Body              = $message.Body.Text
            ID                = $message.ID
            ConversationID    = $message.ConversationID
            ConversationTopic = $message.ConversationTopic
            ItemClass         = $message.ItemClass
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessAppointment }

        switch -Regex ($appointment.subject)
        {
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "ir" -workItem $result}}
            "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "sr" -workItem $result}}
            "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "pr" -workItem $result}}
            "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "cr" -workItem $result}}
            "\[$rrRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $rrClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "rr" -workItem $result}}

            #### activities ####
            "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "ma" -workItem $result}}
            "\[$paRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $paClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "pa" -workItem $result}}
            "\[$saRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $saClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "sa" -workItem $result}}
            "\[$daRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $daClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "da" -workItem $result}}

            #### 3rd party classes, work items, etc. add here ####

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem -message $appointment -wiType $defaultNewWorkItem $true; Set-WorkItemScheduledTime -calAppt $appointment -workItem $returnedNewWorkItemToSchedule ; $message.Accept($true)}
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessAppointment }
    }

    #Process a Calendar Meeting Cancellation
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Canceled")
    {
        $appointment = [PSCustomObject] @{
            StartTime         = $message.Start
            EndTime           = $message.End
            To                = $message.ToRecipients
            From              = $message.From.Address
            Attachments       = $message.Attachments
            Subject           = $message.Subject
            DateTimeReceived  = $message.DateTimeReceived
            DateTimeSent      = $message.DateTimeSent
            Body              = $message.Body.Text
            ID                = $message.ID
            ConversationID    = $message.ConversationID
            ConversationTopic = $message.ConversationTopic
            ItemClass         = $message.ItemClass
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessCancelMeeting }

        switch -Regex ($appointment.subject)
        {
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $irClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "ir" -workItem $result}}
            "\[$srRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $srClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "sr" -workItem $result}}
            "\[$prRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $prClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "pr" -workItem $result}}
            "\[$crRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $crClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "cr" -workItem $result}}
            "\[$rrRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $rrClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "rr" -workItem $result}}

            #### activities ####
            "\[$maRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $maClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "ma" -workItem $result}}
            "\[$paRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $paClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "pa" -workItem $result}}
            "\[$saRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $saClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "sa" -workItem $result}}
            "\[$daRegex[0-9]+\]" {$result = Get-WorkItem -workItemID $matches[0] -workItemClass $daClass; if ($result){Set-WorkItemScheduledTime -calAppt $appointment -workItem $result; $isUpdate = $true; Update-WorkItem -message $appointment -wiType "da" -workItem $result}}

            #### 3rd party classes, work items, etc. add here ####
            "([C][a][n][c][e][l][e][d][:])(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if($mergeReplies -eq $true){$result = Confirm-WorkItem -message $appointment -returnWorkItem $true; Set-WorkItemScheduledTime -calAppt $appointment -workItem $result} else{new-workitem -message $appointment -wiType $defaultNewWorkItem}}

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem -message $appointment -wiType $defaultNewWorkItem $true; Set-WorkItemScheduledTime -calAppt $appointment -workItem $returnedNewWorkItemToSchedule ; $message.Accept($true)}
        }

        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessCancelMeeting }

        #Move to deleted items
        $message.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
    }

    #Process a custom message class as defined through it's Custom Rules Pattern if it's enabled
    else
    {
        if ($UseCustomRules)
        {
            #The message is from an unknown Message Class. Load custom defined Message Classes based on Enums not defined in the SMLets Exchange Connector MP
            $customMessageClasses = (Get-SCSMEnumeration "SMLets.Exchange.Connector.MessageClassEnum$" | Get-SCSMChildEnumeration) | Select-Object id, displayname, @{Name = "ManagementPack"; Expression = {$_.Identifier.Domain[0]}} | Where-Object {$_.ManagementPack -ne "SMLets.Exchange.Connector"}
            if ($customMessageClasses.Count -ge 1)
            {
                #Identify if a custom pattern from External Ticketing in Settings UI matches the current Message's Class name
                $customMatchingClasses = $customMessageClasses | Where-Object {$_.DisplayName -eq $message.ItemClass}
                if ($customMatchingClasses)
                {
                    #No Primary Work Item was matched, Custom Patterns has at least one defined use of IPM.Note, and External Ticketing is enabled
                    $email = [PSCustomObject] @{
                        From                = $message.From.Address
                        To                  = $message.ToRecipients
                        CC                  = $message.CcRecipients
                        Subject             = $message.Subject
                        Attachments         = $message.Attachments
                        Body                = $message.Body.Text
                        DateTimeSent        = $message.DateTimeSent
                        DateTimeReceived    = $message.DateTimeReceived
                        ID                  = $message.ID
                        ConversationID      = $message.ConversationID
                        ConversationTopic   = $message.ConversationTopic
                        ItemClass           = $message.ItemClass
                    }
                    Test-EmailPattern -MessageClass $message.ItemClass -Email $email

                    #mark the message as read on Exchange, move to deleted items
                    $message.IsRead = $true
                    $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve) | Out-Null
                    if ($deleteAfterProcessing){$message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems) | Out-Null}
                }
            }
        }
        else
        {
            #Custom Rules are not enabled
            if ($loggingLevel -ge 4) {New-SMEXCOEvent -Source "Test-EmailPattern" -EventId 13 -Severity "Information" -LogMessage "Custom Rules are not enabled"}
        }
    }

    #increment the number of messages processed if logging is enabled
    if ($loggingLevel -ge 1){$messagesProcessed++; New-SMEXCOEvent -Source "General" -EventId 3 -LogMessage "Processed: $messagesProcessed of $($inbox.Count)" -Severity "Information"}
}

#log the total number of messages processed and the connector's total run time
if ($loggingLevel -ge 1)
{
    $endTime = Get-Date
    $runtime = $endTime - $startTime
    New-SMExcoEvent -Source "General" -EventID 6 -Severity "Information" -LogMessage "Processed $($inbox.Count) messages in:
    Minutes: $($runtime.TotalMinutes)
    Seconds: $($runtime.TotalSeconds)
    Milliseconds: $($runtime.TotalMilliseconds)"
}
