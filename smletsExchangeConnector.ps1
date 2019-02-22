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
Contributors: Martin Blomgren, Leigh Kilday, Tom Hendricks
Reviewers: Tom Hendricks, Brian Weist
Inspiration: The Cireson Community, Anders Asp, Stefan Roth, and (of course) Travis Wright for SMlets examples
Requires: PowerShell 4+, SMlets, and Exchange Web Services API (already installed on SCSM workflow server by virtue of stock Exchange Connector).
    3rd party option: If you're a Cireson customer and make use of their paid SCSM Portal with HTML Knowledge Base this will work as is
        if you aren't, you'll need to create your own Type Projection for Change Requests for the Add-ActionLogEntry
        function. Navigate to that function to read more. If you don't make use of their HTML KB, you'll want to keep $searchCiresonHTMLKB = $false
    Signed/Encrypted option: .NET 4.5 is required to use MimeKit.dll
Misc: The Release Record functionality does not exist in this as no out of box (or 3rd party) Type Projection exists to serve this purpose.
    You would have to create your own Type Projection in order to leverage this.
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

#region #### Configuration ####
#define the SCSM management server, this could be a remote name or localhost
$scsmMGMTServer = "localhost"
#if you are running this script in SMA or Orchestrator, you may need/want to present a credential object to the management server.  Leave empty, otherwise.
$scsmMGMTCreds = $null

#define/use SCSM WF credentials
#$exchangeAuthenticationType - "windows" or "impersonation" are valid inputs here.
    #Windows will use the credentials that start this script in order to authenticate to Exchange and retrieve messages
        #choosing this option only requires the $workflowEmailAddress variable to be defined
        #this is ideal if you'll be using Task Scheduler or SMA to initiate this
    #Impersonation will use the credentials that are defined here to connect to Exchange and retrieve messages
        #choosing this option requires the $workflowEmailAddress, $username, $password, and $domain variables to be defined
#UseAutoDiscover = Determines whether ($true) or not ($false) to connect to Exchange using autodiscover.  If $false, provide a URL for $ExchangeEndpoint
    #ExchangeEndpoint = A URL in the format of 'https://<yourservername.domain.tld>/EWS/Exchange.asmx' such as 'https://mail.contoso.com/EWS/Exchange.asmx'
$exchangeAuthenticationType = "windows"
$workflowEmailAddress = ""
$username = ""
$password = ""
$domain = ""
$UseAutodiscover = $true
$ExchangeEndpoint = ""

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
#fromKeyword = If $includeWholeEmail is set to true, messages will be parsed UNTIL they find this word
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
$defaultNewWorkItem = "ir"
$defaultIRTemplateName = "IR Template Name Goes Here"
$defaultSRTemplateName = "SR Template Name Goes Here"
$defaultPRTemplateName = "PR Template Name Goes Here"
$defaultCRTemplateName = "CR Template Name Goes Here"
$defaultIncidentResolutionCategory = ""
$defaultProblemResolutionCategory = ""
$defaultServiceRequestImplementationCategory = ""
$checkAttachmentSettings = $false
$minFileSizeInKB = "25"
$createUsersNotInCMDB = $true
$includeWholeEmail = $false
$attachEmailToWorkItem = $false
$voteOnBehalfOfGroups = $false
$fromKeyword = "From"
$UseMailboxRedirection = $false
$Mailboxes = @{
    "MyOtherMailbox@company.com" = @{"DefaultWiType"="SR";"IRTemplate"="My IR Template";"SRTemplate"="My SR Template";"PRTemplate"="My PR Template";"CRTemplate"="My CR Template"};
}
$CreateNewWorkItemWhenClosed = $false
$takeRequiresGroupMembership = $false
$crSupportGroupEnumGUID = ""
$maSupportGroupEnumGUID = ""
$redactPiiFromMessage = $false
$changeIncidentStatusOnReply = $false
$changeIncidentStatusOnReplyAffectedUser = "IncidentStatusEnum.Active$"
$changeIncidentStatusOnReplyAssignedTo = "IncidentStatusEnum.Active.Pending$"
$changeIncidentStatusOnReplyRelatedUser = "IncidentStatusEnum.Active$"

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
$processCalendarAppointment = $false
$processDigitallySignedMessages = $false
$processEncryptedMessages = $false
$certStore = "user"
$mergeReplies = $false

#optional, enable integration with Cireson Knowledge Base/Service Catalog
#this uses the now depricated Cireson KB API Search by Text, it works as of v7.x but should be noted it could be entirely removed in future portals
#$numberOfWordsToMatchFromEmailToRO = defines the minimum number of words that must be matched from an email/new work item before Request Offerings will be
    #suggested to the Affected User about them
#$numberOfWordsToMatchFromEmailToKA = defines the minimum number of words that must be matched from an email/new work item before Knowledge Articles will be
    #suggested to the Affected User about them
#searchAvailableCiresonPortalOfferings = search available Request Offerings within the Affected User's permission scope based words matched in
    #their email/new work item
#enableSetFirstResponseDateOnSuggestions = When Knowledge Article or Request Offering suggestions are made to the Affected User, you can optionally
    #set the First Response Date value on a New Work Item
#$ciresonPortalServer = URL that will be used to search for KB articles via invoke-restmethod. Make sure to leave the "/" after your tld!
#$ciresonPortalWindowsAuth = how invoke-restmethod should attempt to authenticate to your portal server.
    #Leave true if your portal uses Windows Auth, change to False for Forms authentication.
    #If using forms, you'll need to set the ciresonPortalUsername and Password variables. For ease, you could set this equal to the username/password defined above
$searchCiresonHTMLKB = $false
$numberOfWordsToMatchFromEmailToRO = 1
$numberOfWordsToMatchFromEmailToKA = 1
$searchAvailableCiresonPortalOfferings = $false
$enableSetFirstResponseDateOnSuggestions = $false
$ciresonPortalServer = "https://portalserver.domain.tld/"
$ciresonPortalWindowsAuth = $true
$ciresonPortalUsername = ""
$ciresonPortalPassword = ""

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
$enableSCSMAnnouncements = $false
$enableCiresonPortalAnnouncements = $false
$announcementKeyword = "announcement"
$approvedADGroupForSCSMAnnouncements = "my custom AD SCSM Authorized Announcers Users group"
$approvedUsersForSCSMAnnouncements = "myfirst.email@domain.com", "mysecond.address@domain.com"
$approvedMemberTypeForSCSMAnnouncer = "group"
$lowAnnouncemnentPriorityKeyword = "low"
$criticalAnnouncemnentPriorityKeyword = "high"
$lowAnnouncemnentExpirationInHours = 7
$normalAnnouncemnentExpirationInHours = 3
$criticalAnnouncemnentExpirationInHours = 1

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
#enableAzureCognitiveServicesPriorityScoring = If enabled with enableAzureCognitiveServicesForNewWI, the Sentiment Score will be used
    #to set the Impact & Urgency and/or Urgency $ Priority on Incidents or Service Requests. Bounds can be edited within
    #the Get-ACSWorkItemPriority function. 
#azureRegion = where Cognitive Services is deployed as seen in it's respective settings pane,
    #i.e. ukwest, eastus2, westus, northcentralus
#azureCogSvcTextAnalyticsAPIKey = API key for your cognitive services text analytics deployment. This is found in the settings pane for Cognitive Services in https://portal.azure.com
#minPercentToCreateServiceRequest = The minimum sentiment rating required to create a Service Request, a number less than this will create an Incident
$enableAzureCognitiveServicesForNewWI = $false
$minPercentToCreateServiceRequest = "95"
$enableAzureCognitiveServicesForKA = $false
$enableAzureCognitiveServicesForRO = $false
$enableAzureCognitiveServicesPriorityScoring = $false
$azureRegion = ""
$azureCogSvcTextAnalyticsAPIKey = ""

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
$enableKeywordMatchForNewWI = $false
$workItemTypeOverrideKeywords = "(?<!in )error|problem|fail|crash|\bjam\b|\bjammed\b|\bjamming\b|broke|froze|issue|unable"
$workItemOverrideType = "ir"

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
$enableSCOMIntegration = $false
$scomMGMTServer = ""
$approvedMemberTypeForSCOM = "group"
$approvedADGroupForSCOM = "my custom AD SCOM Authorized Users group"
$approvedUsersForSCOM = "myfirst.email@domain.com", "mysecond.address@domain.com"
$distributedApplicationHealthKeyword = "health"

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
$privateCommentKeyword = "private"

#define the path to the Exchange Web Services API and MimeKit
#the PII regex file and HTML Suggestion Template paths will only be leveraged if these features are enabled above.
#$htmlSuggestionTemplatePath must end with a "\"
$exchangeEWSAPIPath = "C:\Program Files\Microsoft\Exchange\Web Services\1.2\Microsoft.Exchange.WebServices.dll"
$mimeKitDLLPath = "C:\smletsExchangeConnector\mimekit.dll"
$piiRegexPath = "C:\smletsExchangeConnector\pii_regex.txt"
$htmlSuggestionTemplatePath = "c:\smletsexchangeconnector\htmlEmailTemplates\"

#enable logging per standard Exchange Connector registry keys
#valid options on that registry key are 1 to 7 where 7 is the most verbose
#$loggingLevel = (Get-ItemProperty "HKLM:\Software\Microsoft\System Center Service Manager Exchange Connector" -ErrorAction SilentlyContinue).LoggingLevel
#$loggingLevel = 1

#$ceScripts = invoke the Custom Events script, will optionally load custom/proprietary scripts as certain events occur.
    # set this equal to empty quotes ("") to turn custom events OFF
    # if using this feature, DO NOT USE QUOTES.  Start with a period/dot and then add the path to the script/runbook.
    # If running in SMA OR as a scheduled task with the custom events script in the same folder, use this format: . .\smletsExchangeConnector_CustomEvents.ps1
    # If running as a scheduled task and you have stored the events script in another folder, use this format: . C:\otherFolder\smletsExchangeConnector_CustomEvents.ps1'
$ceScripts = . .\smletsExchangeConnector_CustomEvents.ps1
#endregion #### Configuration ####

#region #### Process User Configs and Prep SMLets ####
# Ensure SMLets is loaded in the current session.
if (-Not (Get-Module SMLets)) {
    Import-Module SMLets
    # If the import is unsuccessful and PowerShell 5+ is installed, pull SMLets from the gallery and install.
    if ($PsVersionTable.PsVersion.Major -ge 5 -And (-Not (Get-Module SMLets))) {
        Find-Module SMLets | Install-Module
        Import-Module SMLets
    }
}

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
    $Mailboxes.add("$($workflowEmailAddress)", @{"DefaultWiType"=$defaultNewWorkItem;"IRTemplate"=$DefaultIRTemplateName;"SRTemplate"=$DefaultSRTemplateName;"PRTemplate"=$DefaultPRTemplateName;"CRTemplate"=$DefaultCRTemplateName})
}
else {
    $defaultIRTemplate = Get-SCSMObjectTemplate -DisplayName $DefaultIRTemplateName @scsmMGMTParams | where-object {$_.displayname -eq "$DefaultIRTemplateName"}
    $defaultSRTemplate = Get-SCSMObjectTemplate -DisplayName $DefaultSRTemplateName @scsmMGMTParams | where-object {$_.displayname -eq "$DefaultSRTemplateName"}
    $defaultPRTemplate = Get-SCSMObjectTemplate -DisplayName $DefaultPRTemplateName @scsmMGMTParams | where-object {$_.displayname -eq "$DefaultPRTemplateName"}
    $defaultCRTemplate = Get-SCSMObjectTemplate -DisplayName $DefaultCRTemplateName @scsmMGMTParams | where-object {$_.displayname -eq "$DefaultCRTemplateName"}
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

$userClass = get-scsmclass -name "System.User$" @scsmMGMTParams
$domainUserClass = get-scsmclass -name "System.Domain.User$" @scsmMGMTParams
$notificationClass = get-scsmclass -name "System.Notification.Endpoint$" @scsmMGMTParams

$irLowImpact = Get-SCSMEnumeration -name "System.WorkItem.TroubleTicket.ImpactEnum.Low$" @scsmMGMTParams
$irLowUrgency = Get-SCSMEnumeration -name "System.WorkItem.TroubleTicket.UrgencyEnum.Low$" @scsmMGMTParams
$irActiveStatus = Get-SCSMEnumeration -name "IncidentStatusEnum.Active$" @scsmMGMTParams

$affectedUserRelClass = get-scsmrelationshipclass -name "System.WorkItemAffectedUser$" @scsmMGMTParams
$assignedToUserRelClass  = Get-SCSMRelationshipClass -name "System.WorkItemAssignedToUser$" @scsmMGMTParams
$createdByUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItemCreatedByUser$" @scsmMGMTParams
$workResolvedByUserRelClass = Get-SCSMRelationshipClass -name "System.WorkItem.TroubleTicketResolvedByUser$" @scsmMGMTParams
$wiRelatesToCIRelClass = Get-SCSMRelationshipClass -name "System.WorkItemRelatesToConfigItem$" @scsmMGMTParams
$wiRelatesToWIRelClass = Get-SCSMRelationshipClass -name "System.WorkItemRelatesToWorkItem$" @scsmMGMTParams
$wiContainsActivityRelClass = Get-SCSMRelationshipClass -name "System.WorkItemContainsActivity$" @scsmMGMTParams
$sysUserHasPrefRelClass = Get-SCSMRelationshipClass -name "System.UserHasPreference$" @scsmMGMTParams

$fileAttachmentClass = Get-SCSMClass -Name "System.FileAttachment$" @scsmMGMTParams
$fileAttachmentRelClass = Get-SCSMRelationshipClass -name "System.WorkItemHasFileAttachment$" @scsmMGMTParams
$fileAddedByUserRelClass = Get-SCSMRelationshipClass -name "System.FileAttachmentAddedByUser$" @scsmMGMTParams
$managementGroup = New-Object Microsoft.EnterpriseManagement.EnterpriseManagementGroup $scsmMGMTServer

$irTypeProjection = Get-SCSMTypeProjection -name "system.workitem.incident.projectiontype$" @scsmMGMTParams
$srTypeProjection = Get-SCSMTypeProjection -name "system.workitem.servicerequestprojection$" @scsmMGMTParams
$prTypeProjection = Get-SCSMTypeProjection -name "system.workitem.problem.projectiontype$" @scsmMGMTParams
$crTypeProjection = Get-SCSMTypeProjection -Name "system.workitem.changerequestprojection$" @scsmMGMTParams

$userHasPrefProjection = Get-SCSMTypeProjection -name "System.User.Preferences.Projection$" @scsmMGMTParams

# Retrieve Support Group Class Extensions on CR/MA if defined
if ($maSupportGroupEnumGUID)
{
    $maSupportGroupPropertyName = ($maClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.EnumType -like "*$maSupportGroupEnumGUID*")}).Name
}
if ($crSupportGroupEnumGUID)
{
    $crSupportGroupPropertyName = ($crClass.GetProperties(1, 1) | where-object {($_.SystemType.Name -eq "Enum") -and ($_.EnumType -like "*$crSupportGroupEnumGUID*")}).Name
}
#endregion

#region #### Exchange Connector Functions ####
function New-WorkItem ($message, $wiType, $returnWIBool) 
{
    $from = $message.From
    $to = $message.To
    $cced = $message.CC
    $title = $message.subject
    $description = $message.body

    #removes PII if RedactPiiFromMessage is enabled
    if ($redactPiiFromMessage -eq $true)
    {
        $description = remove-PII $description
    }
    
    #if the message is longer than 4000 characters take only the first 4000.
    if ($description.length -ge "4000")
    {
        $description = $description.substring(0,4000)
    }

    #find Affected User from the From Address
    $relatedUsers = @()
    $affectedUser = Get-SCSMUserByEmailAddress -EmailAddress "$from"
    if ((!$affectedUser) -and ($createUsersNotInCMDB -eq $true)) {$affectedUser = create-userincmdb "$from"}

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
                    $relatedUser = create-userincmdb $to.address
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
                        $relatedUser = create-userincmdb $ToSMTP.address
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
                    $relatedUser = create-userincmdb $cced.address
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
                        $relatedUser = create-userincmdb $ccSMTP.address
                        $relatedUsers += $relatedUser
                    }
                }
                $x++
            }
        }
    }
    
    $TemplatesForThisMessage = Get-TemplatesByMailbox $message
    
    # Use the global default work item type or, if mailbox redirection is used, use the default work item type for the
    # specific mailbox that the current message was sent to. If Azure Cognitive Services is enabled
    # run the message through it to determine the Default Work Item type. Otherwise, use default if there is no match.
    if ($enableAzureCognitiveServicesForNewWI -eq $true)
    {
        $sentimentScore = Get-AzureEmailSentiment -messageToEvaluate $message.body
        
        #if the sentiment is greater than or equal to what is defined, create a Service Request. Optionally define Urgency/Priority in that order.
        if ($sentimentScore -ge [int]$minPercentToCreateServiceRequest)
        {
            $workItemType = "sr"
            if ($enableAzureCognitiveServicesPriorityScoring -eq $true)
            {
                $priorityEnumArray = Get-ACSWorkItemPriority -score $sentimentScore -wiClass "System.WorkItem.ServiceRequest"
            }
        }
        else #sentiment is lower than defined value, create an Incident. Optionally define Impact/Urgency in that order.
        {
            $workItemType = "ir"
            if ($enableAzureCognitiveServicesPriorityScoring -eq $true)
            {
                $priorityEnumArray = Get-ACSWorkItemPriority -score $sentimentScore -wiClass "System.WorkItem.Incident"
            }
        }
    }
    elseif ($enableKeywordMatchForNewWI -eq $true -and $(Test-KeywordsFoundInMessage $message) -eq $true) {
        #Keyword override is true and keyword(s) found in message
        $workItemType = $workItemOverrideType
    }
    elseif ($UseMailboxRedirection -eq $true) {
        $workItemType = if ($TemplatesForThisMessage) {$TemplatesForThisMessage["DefaultWiType"]} else {$defaultNewWorkItem}
    }
    else {
        $workItemType = $defaultNewWorkItem
    }
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-BeforeCreateAnyWorkItem }
    
    #create the Work Item based on the globally defined Work Item type and Template
    switch ($workItemType) 
    {
        "ir" {
                    if ($UseMailboxRedirection -eq $true -And $TemplatesForThisMessage.Count -gt 0) {
                        $IRTemplate = Get-ScsmObjectTemplate -DisplayName $($TemplatesForThisMessage["IRTemplate"]) @scsmMGMTParams
                    }
                    else {
                        $IRTemplate = $defaultIRTemplate
                    }
                    $newWorkItem = New-SCSMObject -Class $irClass -PropertyHashtable @{"ID" = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Incident")["Prefix"] + "{0}"; "Status" = $irActiveStatus; "Title" = $title; "Description" = $description; "Classification" = $null; "Impact" = $irLowImpact; "Urgency" = $irLowUrgency; "Source" = "IncidentSourceEnum.Email$"} -PassThru @scsmMGMTParams
                    $irProjection = Get-SCSMObjectProjection -ProjectionName $irTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Set-SCSMObjectTemplate -Projection $irProjection -Template $IRTemplate @scsmMGMTParams
                    if (($enableAzureCognitiveServicesPriorityScoring -eq $true) -and ($enableAzureCognitiveServicesForNewWI -eq $true)) {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Impact" = $priorityEnumArray[0]; "Urgency" = $priorityEnumArray[1]} @scsmMGMTParams}
                    Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description} @scsmMGMTParams
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
                        New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk @scsmMGMTParams
                        }
                    }
                    
                    #### Determine auto-response logic for Knowledge Base and/or Request Offering Search ####
                    $ciresonSuggestionURLs = Get-CiresonSuggestionURL -SuggestKA:$searchCiresonHTMLKB -AzureKA:$enableAzureCognitiveServicesForKA -SuggestRO:$searchAvailableCiresonPortalOfferings -AzureRO:$enableAzureCognitiveServicesForRO -WorkItem $newWorkItem -AffectedUser $affectedUser
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
                    else
                    {
                        #both options are set to $false
                        #don't suggest anything to the Affected User based on their recently created Default Work Item
                    }

                    # Custom Event Handler
                    if ($ceScripts) { Invoke-AfterCreateIR }
                    
                }
        "sr" {
                    if ($UseMailboxRedirection -eq $true -and $TemplatesForThisMessage.Count -gt 0) {
                        $SRTemplate = Get-ScsmObjectTemplate -DisplayName $($TemplatesForThisMessage["SRTemplate"]) @scsmMGMTParams
                    }
                    else {
                        $SRTemplate = $defaultSRTemplate
                    }
                    $newWorkItem = new-scsmobject -class $srClass -propertyhashtable @{"ID" = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.ServiceRequest")["Prefix"] + "{0}"; "Title" = $title; "Description" = $description; "Status" = "ServiceRequestStatusEnum.New$"} -PassThru @scsmMGMTParams
                    $srProjection = Get-SCSMObjectProjection -ProjectionName $srTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Apply-SCSMTemplate -Projection $srProjection -Template $SRTemplate
                    #Set-SCSMObjectTemplate -projection $srProjection -Template $SRTemplate @scsmMGMTParams
                    if (($enableAzureCognitiveServicesPriorityScoring -eq $true) -and ($enableAzureCognitiveServicesForNewWI -eq $true)) {Set-SCSMObject -SMObject $newWorkItem -PropertyHashtable @{"Urgency" = $priorityEnumArray[0]; "Priority" = $priorityEnumArray[1]} @scsmMGMTParams}
                    Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description} @scsmMGMTParams
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
                        New-SCSMRelationshipObject -Relationship $affectedUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
                    }
                    if ($relatedUsers)
                    {
                        foreach ($relatedUser in $relatedUsers)
                        {
                            New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $newWorkItem -Target $relatedUser -Bulk @scsmMGMTParams
                        }
                    }
                    
                    #### Determine auto-response logic for Knowledge Base and/or Request Offering Search ####
                    $ciresonSuggestionURLs = Get-CiresonSuggestionURL -SuggestKA:$searchCiresonHTMLKB -AzureKA:$enableAzureCognitiveServicesForKA -SuggestRO:$searchAvailableCiresonPortalOfferings -AzureRO:$enableAzureCognitiveServicesForRO -WorkItem $newWorkItem -AffectedUser $affectedUser
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
                    else
                    {
                        #both options are set to $false
                        #don't suggest anything to the Affected User based on their recently created Default Work Item
                    }
                    
                    # Custom Event Handler
                    if ($ceScripts) { Invoke-AfterCreateSR }					
                }
        "pr" {
                    if ($UseMailboxRedirection -eq $true -and $TemplatesForThisMessage.Count -gt 0) {
                        $PRTemplate = Get-ScsmObjectTemplate -DisplayName $($TemplatesForThisMessage["PRTemplate"]) @scsmMGMTParams
                    }
                    else {
                        $PRTemplate = $defaultPRTemplate
                    }
                    $newWorkItem = new-scsmobject -class $prClass -propertyhashtable @{"ID" = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Problem")["Prefix"] + "{0}"; "Title" = $title; "Description" = $description; "Status" = "ProblemStatusEnum.Active$"} -PassThru @scsmMGMTParams
                    $prProjection = Get-SCSMObjectProjection -ProjectionName $prTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                    if($message.Attachments){Attach-FileToWorkItem $message $newWorkItem.ID}
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $newWorkItem.ID}
                    Set-SCSMObjectTemplate -Projection $prProjection -Template $defaultPRTemplate @scsmMGMTParams
                    Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description} @scsmMGMTParams
                    #no Affected User to set on a Problem, set Created By using the Affected User object if it exists
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
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
                        $CRTemplate = Get-ScsmObjectTemplate -DisplayName $($TemplatesForThisMessage["CRTemplate"]) @scsmMGMTParams
                    }
                    else {
                        $CRTemplate = $defaultCRTemplate
                    }
                    $newWorkItem = new-scsmobject -class $crClass -propertyhashtable @{"ID" = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.ChangeRequest")["Prefix"] + "{0}"; "Title" = $title; "Description" = $description; "Status" = "ChangeStatusEnum.New$"} -PassThru @scsmMGMTParams
                    $crProjection = Get-SCSMObjectProjection -ProjectionName $crTypeProjection.Name -Filter "ID -eq $($newWorkItem.Name)" @scsmMGMTParams
                    #Set-SCSMObjectTemplate -Projection $crProjection -Template $defaultCRTemplate @scsmMGMTParams
                    Apply-SCSMTemplate -Projection $crProjection -Template $CRTemplate
                    Set-ScsmObject -SMObject $newWorkItem -PropertyHashtable @{"Description" = $description} @scsmMGMTParams
                    #The Affected User relationship exists on Change Requests, but does not exist on the CR Form out of box.
                    #Cireson SCSM Portal customers may wish to set the Sender as the Affected User so that it follows Incident/Service Request style functionality in that the Sender/User
                    #in question can see the CR in the "My Requests" section of their SCSM portal. This can be acheived by uncommenting the New-SCSMRelationshipObject seen below
                    #Set Created By using the Affected User object if it exists
                    if ($affectedUser)
                    {
                        New-SCSMRelationshipObject -Relationship $createdByUserRelClass -Source $newWorkItem -Target $affectedUser -Bulk @scsmMGMTParams
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
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-AfterCreateAnyWorkItem }

    if ($returnWIBool -eq $true)
    {
        return $newWorkItem
    }
}

function Update-WorkItem ($message, $wiType, $workItemID) 
{
    #removes PII if RedactPiiFromMessage is enable
    if ($redactPiiFromMessage -eq $true)
    {
        $message.body = remove-PII $message.body
    }
      
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
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-BeforeUpdateAnyWorkItem }
    
    #determine who left the comment
    $commentLeftBy = Get-SCSMUserByEmailAddress -EmailAddress "$($message.From)"
    if ((!$commentLeftBy) -and ($createUsersNotInCMDB -eq $true) ){$commentLeftBy = create-userincmdb $message.From}

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
            $workItem = get-scsmobject -class $irClass -filter "Name -eq '$workItemID'" @scsmMGMTParams

            try {$existingWiStatusName = $workItem.Status.Name} catch {}
            if ($CreateNewWorkItemWhenClosed -eq $true -And $existingWiStatusName -eq "IncidentStatusEnum.Closed") {
                $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" @scsmMGMTParams | foreach-object {Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $_ @scsmMGMTParams).sourceobject.id @scsmMGMTParams} | where-object {$_.Status -ne "IncidentStatusEnum.Closed"}
                if (($relatedWorkItemFromAttachmentSearch | get-unique).count -eq 1)
                {
                    Update-WorkItem -message $message -wiType "ir" -workItemID $relatedWorkItemFromAttachmentSearch.Name
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
                try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                if ($assignedToSMTP.TargetAddress -eq $affectedUserSMTP.TargetAddress){$assignedToSMTP = $null}
                #write to the Action log and take action on the Work Item if neccesary
                switch ($message.From)
                {
                    $affectedUserSMTP.TargetAddress {
                        if ($changeIncidentStatusOnReply) {Set-SCSMObject -SMObject $workItem -Property Status -Value "$changeIncidentStatusOnReplyAffectedUser"}
                        switch -Regex ($commentToAdd) {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterAcknowledge }}}
                            "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Resolved" -IsPrivate $false; if ($defaultIncidentResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property ResolutionCategory -Value $defaultIncidentResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                            "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Closed" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                            "\[$takeKeyword]" { 
                                $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.TierQueue.Id
                                if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                    New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Assign" -IsPrivate $false
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterTake }
                                }
                                else {
                                    #TODO: Send an email to let them know it failed?
                                }
                            }
                            "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Active$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}}
                            "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "\[$irRegex[0-9]+\]")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem $message "ir" $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk @scsmMGMTParams}catch{}; if ($ceScripts) { Invoke-AfterReactivate }}}
                            {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false}
                            default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false}
                        }
                    }
                    $assignedToSMTP.TargetAddress {
                        if ($changeIncidentStatusOnReply) {Set-SCSMObject -SMObject $workItem -Property Status -Value "$changeIncidentStatusOnReplyAssignedTo"}
                        switch -Regex ($commentToAdd) {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterAcknowledge }}}
                            "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Resolved" -IsPrivate $false; if ($defaultIncidentResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property ResolutionCategory -Value $defaultIncidentResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                            "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Closed" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                            "\[$takeKeyword]" { 
                                $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.TierQueue.Id
                                if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                    New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterTake }
                                }
                                else {
                                    #TODO: Send an email to let them know it failed?
                                }
                            }
                            "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Active$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}}
                            "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "\[$irRegex[0-9]+\]")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem $message "ir" $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk @scsmMGMTParams}catch{}; if ($ceScripts) { Invoke-AfterReactivate }}}
                            {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                            "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $true}
                            default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                        }
                    }
                    default {
                        if ($changeIncidentStatusOnReply) {Set-SCSMObject -SMObject $workItem -Property Status -Value "$changeIncidentStatusOnReplyRelatedUser"}
                        switch -Regex ($commentToAdd) {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterAcknowledge }}}
                            "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Resolved" -IsPrivate $false; if ($defaultIncidentResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property ResolutionCategory -Value $defaultIncidentResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                            "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "IncidentStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Closed" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                            "\[$takeKeyword]" { 
                                $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.TierQueue.Id
                                if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                    New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterTake }
                                }
                                else {
                                    #TODO: Send an email to let them know it failed?
                                }
                            }
                            "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "IncidentStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "IncidentStatusEnum.Active$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}}
                            "\[$reactivateKeyword]" {if (($workItem.Status.Name -eq "IncidentStatusEnum.Closed") -and ($message.Subject -match "\[$irRegex[0-9]+\]")){$message.subject = $message.Subject.Replace("[" + $Matches[0] + "]", ""); $returnedWorkItem = New-WorkItem $message "ir" $true; try{New-SCSMRelationshipObject -Relationship $wiRelatesToWIRelClass -Source $workItem -Target $returnedWorkItem -Bulk @scsmMGMTParams}catch{}; if ($ceScripts) { Invoke-AfterReactivate }}}
                            {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                            "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $true}
                            default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $null}
                        }
                    }
                }               
                #relate the user to the work item
                New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams
                #add any new attachments
                if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
            }

            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterUpdateIR }
        } 
        "sr" {
            $workItem = get-scsmobject -class $srClass -filter "Name -eq '$workItemID'" @scsmMGMTParams

            try {$existingWiStatusName = $workItem.Status.Name} catch {}
            if ($CreateNewWorkItemWhenClosed -eq $true -And $existingWiStatusName -eq "ServiceRequestStatusEnum.Closed") {
                $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" @scsmMGMTParams | foreach-object {Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $_ @scsmMGMTParams).sourceobject.id @scsmMGMTParams} | where-object {$_.Status -ne "ServiceRequestStatusEnum.Closed"}
                if (($relatedWorkItemFromAttachmentSearch | get-unique).count -eq 1)
                {
                    Update-WorkItem -message $message -wiType "sr" -workItemID $relatedWorkItemFromAttachmentSearch.Name
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
                try {$affectedUser = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $affectedUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                if($affectedUser){$affectedUserSMTP = Get-SCSMRelatedObject -SMObject $affectedUser @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                if ($assignedToSMTP.TargetAddress -eq $affectedUserSMTP.TargetAddress){$assignedToSMTP = $null}
                switch ($message.From)
                {
                    $affectedUserSMTP.TargetAddress {
                        switch -Regex ($commentToAdd)
                        {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false}}
                            "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                            "\[$takeKeyword]" {
                                $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.SupportGroup.Id
                                if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                    New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "Assign" -IsPrivate $false
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterTake }
                                }
                                else {
                                    #TODO: Send an email to let them know it failed?
                                }
                            }
                            "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"CompletedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Completed$"; "Notes" = "$commentToAdd"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($defaultServiceRequestImplementationCategory) {Set-SCSMObject -SMObject $workItem -Property ImplementationResults -Value $defaultServiceRequestImplementationCategory}; if ($ceScripts) { Invoke-AfterCompleted }}
                            "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Canceled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                            "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $affectedUser -Action "EndUserComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                            default {if($commentToAdd -match "#$privateCommentKeyword"){$isPrivateBool = $true}else{$isPrivateBool = $false};Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $isPrivateBool}
                        }
                    }
                    $assignedToSMTP.TargetAddress {
                        switch -Regex ($commentToAdd)
                        {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}}
                            "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                            "\[$takeKeyword]" {
                                $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.SupportGroup.Id
                                if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                    New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterTake }
                                }
                                else {
                                    #TODO: Send an email to let them know it failed?
                                }
                            }
                            "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"CompletedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Completed$"; "Notes" = "$commentToAdd"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($defaultServiceRequestImplementationCategory) {Set-SCSMObject -SMObject $workItem -Property ImplementationResults -Value $defaultServiceRequestImplementationCategory}; if ($ceScripts) { Invoke-AfterCompleted }}
                            "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Canceled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                            "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}           
                            "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $true}
                            default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                        }
                    }
                    default {
                        switch -Regex ($commentToAdd)
                        {
                            "\[$acknowledgedKeyword]" {if ($workItem.FirstResponseDate -eq $null){Set-SCSMObject -SMObject $workItem -Property FirstResponseDate -Value $message.DateTimeSent.ToUniversalTime() @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}}
                            "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                            "\[$takeKeyword]" {
                                $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.SupportGroup.Id
                                if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                    New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterTake }
                                }
                                else {
                                    #TODO: Send an email to let them know it failed?
                                }
                            }
                            "\[$completedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"CompletedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Completed$"; "Notes" = "$commentToAdd"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($defaultServiceRequestImplementationCategory) {Set-SCSMObject -SMObject $workItem -Property ImplementationResults -Value $defaultServiceRequestImplementationCategory}; if ($ceScripts) { Invoke-AfterCompleted }}
                            "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ServiceRequestStatusEnum.Canceled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                            "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ServiceRequestStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}           
                            "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $true}
                            default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $null}
                        }
                    }
                }
                #relate the user to the work item
                New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams
                #add any new attachments
                if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
            }
            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterUpdateSR }
        } 
        "pr" {
                    $workItem = get-scsmobject -class $prClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                    try {$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log
                    switch ($message.From)
                    {
                        $assignedToSMTP.TargetAddress {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Resolved" -IsPrivate $false; if ($defaultProblemResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property Resolution -Value $defaultProblemResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "\[$takeKeyword]" {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk;
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false;
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    if ($ceScripts){ Invoke-AfterTake }}
                                "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Active$" @scsmMGMTParams}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                            }
                        }
                        default {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$resolvedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ResolvedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Resolved$"; "ResolutionDescription" = "$commentToAdd"} @scsmMGMTParams; New-SCSMRelationshipObject -Relationship $workResolvedByUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Resolved" -IsPrivate $false; if ($defaultProblemResolutionCategory) {Set-SCSMObject -SMObject $workItem -Property Resolution -Value $defaultProblemResolutionCategory}; if ($ceScripts) { Invoke-AfterResolved }}
                                "\[$closedKeyword]" {Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"ClosedDate" = (Get-Date).ToUniversalTime(); "Status" = "ProblemStatusEnum.Closed$"} @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterClosed }}
                                "\[$takeKeyword]" {New-SCSMRelationshipObject -relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk;
                                    Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false;
                                    if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                    if ($ceScripts) { Invoke-AfterTake }}
                                "\[$reactivateKeyword]" {if ($workItem.Status.Name -eq "ProblemStatusEnum.Resolved") {Set-SCSMObject -SMObject $workItem -Property Status -Value "ProblemStatusEnum.Active$" @scsmMGMTParams}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Reactivate" -IsPrivate $false; if ($ceScripts) { Invoke-AfterReactivate }}
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                            }
                        }
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
                    
                    # Custom Event Handler
                    if ($ceScripts) { Invoke-AfterUpdatePR }
                }
        "cr" {
                    $workItem = get-scsmobject -class $crClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                    try{$assignedTo = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                    if($assignedTo){$assignedToSMTP = Get-SCSMRelatedObject -SMObject $assignedTo @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    #write to the Action log and take action on the Work Item if neccesary
                    switch ($message.From)
                    {
                        $assignedToSMTP.TargetAddress {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                "\[$takeKeyword]" { 
                                    $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$crSupportGroupPropertyName.Id
                                    if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                        New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "Assign" -IsPrivate $false
                                        if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        #TODO: Send an email to let them know it failed?
                                    }
                                }
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "EndUserComment" -IsPrivate $false}
                                "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $true}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $assignedTo -Action "AnalystComment" -IsPrivate $false}
                            }
                        }
                        default {
                            switch -Regex ($commentToAdd)
                            {
                                "\[$holdKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.OnHold$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterHold }}
                                "\[$cancelledKeyword]" {Set-SCSMObject -SMObject $workItem -Property Status -Value "ChangeStatusEnum.Cancelled$" @scsmMGMTParams; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false; if ($ceScripts) { Invoke-AfterCancelled }}
                                "\[$takeKeyword]" { 
                                    $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$crSupportGroupPropertyName.Id
                                    if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                        New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                        Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "Assign" -IsPrivate $false
                                        if ($workItem.FirstAssignedDate -eq $null) {Set-SCSMObject -SMObject $workItem -Property FirstAssignedDate -Value (Get-Date).ToUniversalTime() @scsmMGMTParams}
                                        # Custom Event Handler
                                        if ($ceScripts) { Invoke-AfterTake }
                                    }
                                    else {
                                        #TODO: Send an email to let them know it failed?
                                    }
                                }
                                {($commentToAdd -match [Regex]::Escape("["+$announcementKeyword+"]")) -and (Get-SCSMAuthorizedAnnouncer -sender $message.from -eq $true)} {if ($enableCiresonPortalAnnouncements) {Set-CiresonPortalAnnouncement -message $message -workItem $workItem}; if ($enableSCSMAnnouncements) {Set-CoreSCSMAnnouncement -message $message -workItem $workItem}; Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                                "#$privateCommentKeyword" {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $true}
                                default {Add-ActionLogEntry -WIObject $workItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "AnalystComment" -IsPrivate $false}
                            } 
                        }
                    }
                    #relate the user to the work item
                    New-SCSMRelationshipObject -Relationship $wiRelatesToCIRelClass -Source $workItem -Target $commentLeftBy -Bulk @scsmMGMTParams
                    #add any new attachments
                    if ($attachEmailToWorkItem -eq $true){Attach-EmailToWorkItem $message $workItem.ID}
                    
                    # Custom Event Handler
                    if ($ceScripts) { Invoke-AfterUpdateCR }
                }
        
        #### activities ####
        "ra" {
                    $workItem = get-scsmobject -class $raClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                    $reviewers = Get-SCSMRelatedObject -SMObject $workItem -Relationship $raHasReviewerRelClass @scsmMGMTParams
                    foreach ($reviewer in $reviewers)
                    {
                        $reviewingUser = get-scsmobject -class $userClass -filter "id -eq '$((Get-SCSMRelatedObject -SMObject $reviewer -Relationship $raReviewerIsUserRelClass @scsmMGMTParams).id)'" @scsmMGMTParams
                        $reviewingUserName = $reviewingUser.UserName #it is necessary to store this in its own variable for the AD filters to work correctly
                        $reviewingUserSMTP = Get-SCSMRelatedObject -SMObject $reviewingUser @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress

                        if ($commentToAdd.length -gt 256) { $decisionComment = $commentToAdd.substring(0,253)+"..." } else { $decisionComment = $commentToAdd }
                        
                        #Reviewer is a User
                        if ([bool] (Get-ADUser @adParams -filter {SamAccountName -eq $reviewingUserName}))
                        {
                            #approved
                            if (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$approvedKeyword]"))
                            {
                                    Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                    New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk @scsmMGMTParams
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterApproved }
                            }
                            #rejected
                            elseif (($reviewingUserSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                            {
                                    Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                    New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $reviewingUser -Bulk @scsmMGMTParams
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

                                $isReviewerGroupMember = Get-ADGroupMember -Server $AdRoot -Identity $reviewingUser.UserName -recursive @adParams | ? { $_.objectClass -eq "user" -and $_.name -eq $votedOnBehalfOfUser.UserName }

                                #approved on behalf of
                                if (($isReviewerGroupMember) -and ($commentToAdd -match "\[$approvedKeyword]"))
                                {
                                    Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Approved$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                    New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $votedOnBehalfOfUser -Bulk @scsmMGMTParams
                                    # Custom Event Handler
                                    if ($ceScripts) { Invoke-AfterApprovedOnBehalf }
                                
                                }
                                #rejected on behalf of
                                elseif (($isReviewerGroupMember) -and ($commentToAdd -match "\[$rejectedKeyword]"))
                                {
                                    Set-SCSMObject -SMObject $reviewer -PropertyHashtable @{"Decision" = "DecisionEnum.Rejected$"; "DecisionDate" = $message.DateTimeSent.ToUniversalTime(); "Comments" = $decisionComment} @scsmMGMTParams
                                    New-SCSMRelationshipObject -Relationship $raVotedByUserRelClass -Source $reviewer -Target $votedOnBehalfOfUser -Bulk @scsmMGMTParams
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
                                }
                                else {
                                    #not allowing votes on behalf of group, or user is not in an eligible group
                                }
                            }
                            else
                            {
                                #not a user or a group
                            }
                        }
                    }
                    
                    # Custom Event Handler
                    if ($ceScripts) { Invoke-AfterUpdateRA }
                }
        "ma" {
                    $workItem = get-scsmobject -class $maClass -filter "Name -eq '$workItemID'" @scsmMGMTParams
                    try {$activityImplementer = get-scsmobject -id (Get-SCSMRelatedObject -SMObject $workItem -Relationship $assignedToUserRelClass @scsmMGMTParams).id @scsmMGMTParams} catch {}
                    if ($activityImplementer){$activityImplementerSMTP = Get-SCSMRelatedObject -SMObject $activityImplementer @scsmMGMTParams | ?{$_.displayname -like "*SMTP"} | select-object TargetAddress}
                    
                    #take
                    switch -Regex ($commentToAdd)
                    {
                        "\[$takeKeyword]" { 
                            $memberOfSelectedTier = Get-TierMembership -UserSamAccountName $commentLeftBy.UserName -TierId $workItem.$maSupportGroupPropertyName.Id
                            if ($takeRequiresGroupMembership -eq $false -or $memberOfSelectedTier -eq $true) {
                                New-SCSMRelationshipObject -Relationship $assignedToUserRelClass -Source $workItem -Target $commentLeftBy @scsmMGMTParams -bulk
                                # Custom Event Handler
                                if ($ceScripts) { Invoke-AfterTake }
                            }
                            else {
                                #TODO: Send an email to let them know it failed?
                            }
                        }
                    }
                    #completed
                    if (($activityImplementerSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$completedKeyword]"))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Completed$"; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams
                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterCompleted }
                    }
                    #skipped
                    elseif (($activityImplementerSMTP.TargetAddress -eq $message.From) -and ($commentToAdd -match "\[$skipKeyword]"))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Status" = "ActivityStatusEnum.Skipped$"; "Skip" = $true; "ActualEndDate" = (get-date).ToUniversalTime(); "Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams
                        # Custom Event Handler
                        if ($ceScripts) { Invoke-AfterSkipped }
                    }
                    #not from the Activity Implementer, add to the MA Notes
                    elseif (($activityImplementerSMTP.TargetAddress -ne $message.From))
                    {
                        Set-SCSMObject -SMObject $workItem -PropertyHashtable @{"Notes" = "$($workItem.Notes)$($activityImplementer.Name) @ $(get-date): $commentToAdd `n"} @scsmMGMTParams
                    }
                    #no keywords, add to the Parent Work Item
                    elseif (($activityImplementerSMTP.TargetAddress -eq $message.From) -and (($commentToAdd -notmatch "\[$completedKeyword]") -or ($commentToAdd -notmatch "\[$skipKeyword]")))
                    {
                        $parentWorkItem = Get-SCSMWorkItemParent $workItem.Get_Id().Guid
                        switch ($parentWorkItem.Classname)
                        {
                            "System.WorkItem.ChangeRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                            "System.WorkItem.ServiceRequest" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                            "System.WorkItem.Incident" {Add-ActionLogEntry -WIObject $parentWorkItem -Comment $commentToAdd -EnteredBy $commentLeftBy -Action "EndUserComment" -IsPrivate $false}
                        }
                            
                    }
                    
                    # Custom Event Handler
                    if ($ceScripts) { Invoke-AfterUpdateMA }
                }
    } 
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-AfterUpdateAnyWorkItem }
}

function Attach-EmailToWorkItem ($message, $workItemID)
{
    # Get attachment limits and attachment count in ticket, if configured to
    if ($checkAttachmentSettings -eq $true) {
        $workItem = Get-ScsmObject @scsmMGMTParams -class $wiClass -filter "Name -eq $workItemID"
        $workItemSettings = Get-SCSMWorkItemSettings -WorkItemClass $workItem.ClassName

        # Get count of attachents already in ticket
        try {$existingAttachmentsCount = (Get-ScsmRelatedObject @scsmMGMTParams -SMObject $workItem -Relationship $fileAttachmentRelClass).Count} catch {}
    }
    
    $messageMime = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $MemoryStream = New-Object System.IO.MemoryStream($messageMime.MimeContent.Content,0,$messageMime.MimeContent.Content.Length)
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-BeforeAttachEmail }
    
    # if #checkAttachmentSettings -eq $true, test whether the email size (IN KB!) exceeds the limit and if the number of existing attachments is under the limit
    if ($checkAttachmentSettings -eq $false -or `
        (($MemoryStream.Length / 1024) -le $($workItemSettings["MaxAttachmentSize"]) -and $existingAttachmentsCount -le $($workItemSettings["MaxAttachments"])))
    {
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
        $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemID'" @scsmMGMTParams
        $WorkItemProjection.__base.Add($emailAttachment, $fileAttachmentRelClass.Target)
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
}

#inspired and modified from Stefan Roth here - https://stefanroth.net/2015/03/28/scsm-passing-attachments-via-web-service-e-g-sma-web-service/
function Attach-FileToWorkItem ($message, $workItemId)
{
    # Get attachment limits and attachment count in ticket, if configured to
    if ($checkAttachmentSettings -eq $true) {
        $workItem = Get-ScsmObject @scsmMGMTParams -class $wiClass -filter "Name -eq $workItemID"
        $workItemSettings = Get-SCSMWorkItemSettings -WorkItemClass $workItem.ClassName
        $attachMaxSize = $workItemSettings["MaxAttachmentSize"]
        $attachMaxCount = $workItemSettings["MaxAttachments"]

        # Get count of attachents already in ticket
        $existingAttachments = Get-ScsmRelatedObject @scsmMGMTParams -SMObject $workItem -Relationship $fileAttachmentRelClass
        # Only use at before the loop
        try {$existingAttachmentsCount = $existingAttachments.Count } catch {}
    }
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-BeforeAttachFiles }
    
    foreach ($attachment in $message.Attachments)
    {

        if ($attachment.gettype().BaseType.Name -like "Mime*")
        {
            $signedAttachArray = $attachment.ContentObject.Stream.ToArray()
            $base64attachment = [System.Convert]::ToBase64String($signedAttachArray)
            $AttachmentContent = [convert]::FromBase64String($base64attachment)
    
            #Create a new MemoryStream object out of the attachment data
            $MemoryStream = New-Object System.IO.MemoryStream($signedAttachArray,0,$signedAttachArray.Length)
    
            if ($MemoryStream.Length -gt $minFileSizeInKB+"kb" -and ($checkAttachmentSettings -eq $false `
                -or ($existingAttachments.Count -lt $attachMaxCount -And $MemoryStream.Length -le "$attachMaxSize"+"mb")))
            {
                #Create the attachment object itself and set its properties for SCSM
                $NewFile = new-object Microsoft.EnterpriseManagement.Common.CreatableEnterpriseManagementObject($ManagementGroup, $fileAttachmentClass)
                $NewFile.Item($fileAttachmentClass, "Id").Value = [Guid]::NewGuid().ToString()
                $NewFile.Item($fileAttachmentClass, "DisplayName").Value = $attachment.FileName
                #$NewFile.Item($fileAttachmentClass, "Description").Value = $attachment.Description
                #$NewFile.Item($fileAttachmentClass, "Extension").Value =   $attachment.Extension
                $NewFile.Item($fileAttachmentClass, "Size").Value =        $MemoryStream.Length
                $NewFile.Item($fileAttachmentClass, "AddedDate").Value =   [DateTime]::Now.ToUniversalTime()
                $NewFile.Item($fileAttachmentClass, "Content").Value =     $MemoryStream
        
                #Add the attachment to the work item and commit the changes
                $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemId'" @scsmMGMTParams
                $WorkItemProjection.__base.Add($NewFile, $fileAttachmentRelClass.Target)
                $WorkItemProjection.__base.Commit()
    
                #create the Attached By relationship if possible
                $attachedByUser = Get-SCSMUserByEmailAddress -EmailAddress "$($message.from)"
                if ($attachedByUser) 
                { 
                    New-SCSMRelationshipObject -Source $NewFile -Relationship $fileAddedByUserRelClass -Target $attachedByUser @scsmMGMTParams -Bulk
                    $existingAttachments.Count += 1
                }
            }
        }
        else
        {
            $attachment.Load()
            $base64attachment = [System.Convert]::ToBase64String($attachment.Content)
    
            #Convert the Base64String back to bytes
            $AttachmentContent = [convert]::FromBase64String($base64attachment)
    
            #Create a new MemoryStream object out of the attachment data
            $MemoryStream = New-Object System.IO.MemoryStream($AttachmentContent,0,$AttachmentContent.length)
    
            if ($MemoryStream.Length -gt $minFileSizeInKB+"kb" -and ($checkAttachmentSettings -eq $false `
                -or ($existingAttachments.Count -lt $attachMaxCount -And $MemoryStream.Length -le "$attachMaxSize"+"mb")))
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
                $WorkItemProjection = Get-SCSMObjectProjection "System.WorkItem.Projection" -Filter "id -eq '$workItemId'" @scsmMGMTParams
                $WorkItemProjection.__base.Add($NewFile, $fileAttachmentRelClass.Target)
                $WorkItemProjection.__base.Commit()
    
                #create the Attached By relationship if possible
                $attachedByUser = Get-SCSMUserByEmailAddress -EmailAddress "$($message.from)"
                if ($attachedByUser) 
                { 
                    New-SCSMRelationshipObject -Source $NewFile -Relationship $fileAddedByUserRelClass -Target $attachedByUser @scsmMGMTParams -Bulk
                    $existingAttachments.Count += 1
                }
            }
        }
    }
    # Custom Event Handler
    if ($ceScripts) { Invoke-AfterAttachFiles }
}

function Get-WorkItem ($workItemID, $workItemClass)
{
    #removes [] surrounding a Work Item ID if neccesary
    if ($workitemID.StartsWith("[") -and $workitemID.EndsWith("]"))
    {
        $workitemID = $workitemID.TrimStart("[").TrimEnd("]")
    }

    #get the work item
    $wi = get-scsmobject -Class $workItemClass -Filter "Name -eq '$workItemID'" @scsmMGMTParams
    return $wi
}

function Get-SCSMUserByEmailAddress ($EmailAddress)
{
    $userSMTPNotification = Get-SCSMObject -Class $notificationClass -Filter "TargetAddress -eq '$EmailAddress'" @scsmMGMTParams | sort-object lastmodified -Descending | select-object -first 1
    if ($userSMTPNotification) 
    { 
        $user = get-scsmobject -id (Get-SCSMRelationshipObject -ByTarget $userSMTPNotification @scsmMGMTParams).sourceObject.id @scsmMGMTParams
        return $user
    }
    else
    {
        return $null
    }
}

# Nested group membership check inspired by Piotr Lewandowski https://gallery.technet.microsoft.com/scriptcenter/Get-nested-group-15f725f2#content
function Get-TierMembership ($UserSamAccountName, $TierId) {
    $isMember = $false

    #define classes
    $mapCls = Get-ScsmClass @scsmMGMTParams -Name "Cireson.SupportGroupMapping"

    #pull the group based on support tier mapping
    $mapping = $mapCls | Get-ScsmObject @scsmMGMTParams | ? { $_.SupportGroupId.Guid -eq $TierId.Guid }
    $groupId = $mapping.AdGroupId

    #get the AD group object name
    $grpInScsm = (Get-ScsmObject @scsmMGMTParams -Id $groupId)
    $grpSamAccountName = $grpInScsm.UserName
    
    #determine which domain to query, in case of multiple domains and trusts
    $AdRoot = (Get-AdDomain @adParams -Identity $grpInScsm.Domain).DNSRoot

    if ($grpSamAccountName) {
        # Get the group membership
        [array]$members = Get-ADGroupMember @adParams -Server $AdRoot -Identity $grpSamAccountName -Recursive

        # loop through the members of the AD group that underpins this support group, and look for the user
        $members | % {
            if ($_.objectClass -eq "user" -and $_.Name -match $UserSamAccountName) {
                $isMember = $true
            }
        }
    }
    return $isMember
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
                $ActivityObject = Get-SCSMObject -Id $WorkItemGUID @scsmMGMTParams
            }
        
            #Retrieve Parent
            Write-Verbose -Message "[PROCESS] Activity: $($ActivityObject.Name)"
            Write-Verbose -Message "[PROCESS] Retrieving WI Parent"
            $ParentRelatedObject = Get-SCSMRelationshipObject -ByTarget $ActivityObject @scsmMGMTParams | ?{$_.RelationshipID -eq $wiContainsActivityRelClass.id.Guid}
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
                Get-SCSMWorkItemParent -WorkItemGUID $ParentObject.Id.GUID @scsmMGMTParams
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
    $newUser = New-SCSMObject -Class $domainUserClass -PropertyHashtable @{"domain" = "$domainAndTLD"; "username" = "$username"; "displayname" = "$userEmail"} @scsmMGMTParams -PassThru

    #create the user notification projection
    $userNoticeProjection = @{__CLASS = "$($domainUserClass.Name)";
                                __SEED = $newUser;
                                Notification = @{__CLASS = "$($notificationClass)";
                                                    __OBJECT = @{"ID" = $newID; "TargetAddress" = "$userEmail"; "DisplayName" = "E-mail address"; "ChannelName" = "SMTP"}
                                                }
                                }

    #create the user's email notification channel
    New-SCSMObjectProjection -Type "$($userHasPrefProjection.Name)" -Projection $userNoticeProjection @scsmMGMTParams
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-AfterUserCreatedInCMDB }
    
    return $newUser
}

#combined previous 4 individual comment functions featured from versions 1 to 1.4.3 into single function and introduced more Action Log functionality
#inspired and modified from Travis Wright here - https://blogs.technet.microsoft.com/servicemanager/2013/01/16/creating-membership-and-hosting-objectsrelationships-using-new-scsmobjectprojection-in-smlets/
#inspired and modified from Anders Asp here - http://www.scsm.se/?p=1423
#inspired and modified from Xapity here - http://www.xapity.com/single-post/2016/11/27/PowerShell-for-SCSM-Updating-the-Action-Log
function Add-ActionLogEntry {
    param (
        [parameter(Mandatory=$true, Position=0)]
        $WIObject,
        [parameter(Mandatory=$true, Position=1)] 
        [ValidateSet("Assign","AnalystComment","Closed","Escalated","EmailSent","EndUserComment","FileAttached","FileDeleted","Reactivate","Resolved","TemplateApplied")] 
        [string] $Action,
        [parameter(Mandatory=$true, Position=2)]
        [string] $Comment,
        [parameter(Mandatory=$true, Position=3)]
        [string] $EnteredBy,
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
        "System.WorkItem.Incident" {New-SCSMObjectProjection -Type "System.WorkItem.IncidentPortalProjection" -Projection $Projection @scsmMGMTParams}
        "System.WorkItem.ServiceRequest" {New-SCSMObjectProjection -Type "System.WorkItem.ServiceRequestProjection" -Projection $Projection @scsmMGMTParams}
        "System.WorkItem.Problem" {New-SCSMObjectProjection -Type "System.WorkItem.Problem.ProjectionType" -Projection $Projection @scsmMGMTParams}
        "System.WorkItem.ChangeRequest" {New-SCSMObjectProjection -Type "Cireson.ChangeRequest.ViewModel" -Projection $Projection @scsmMGMTParams}
    }
}

#if using windows authentication, retrieve a Cireson Web API token
function Get-CiresonPortalAPIToken
{
    $ciresonPortalCredentials = @{"username" = "$ciresonPortalUsername"; "password" = "$ciresonPortalPassword"; "languagecode" = "ENU" } | ConvertTo-Json
    $ciresonTokenURL = $ciresonPortalServer+"api/V3/Authorization/GetToken"
    $ciresonAPIToken = Invoke-RestMethod -uri $ciresonTokenURL -Method post -Body $ciresonPortalCredentials
    $ciresonAPIToken = "Token" + " " + $ciresonAPIToken
    return $ciresonAPIToken
}

#retrieve a user from SCSM through the Cireson Web Portal API
function Get-CiresonPortalUser ($username, $domain)
{
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
function Get-CiresonPortalGroup ($groupEmail)
{
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
    $ciresonPortalGroup = $cwpGroupResponse | select-object @{Name='AccessGroupId'; Expression={$_.Id}}, name | ?{$_.name -eq $($adGroup.Name)}
    return $ciresonPortalGroup
}

#retrieve all the announcements on the portal
function Get-CiresonPortalAnnouncements ($languageCode)
{
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
function Search-AvailableCiresonPortalOfferings ($searchQuery, $ciresonPortalUser)
{
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
            $wordsMatched = ($searchQuery.Split() | ?{($serviceCatalogResult.RequestOfferingTitle -match "\b$_\b") -or ($serviceCatalogResult.RequestOfferingDescription -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToRO)
            {
                $ciresonPortalRequestURL = "`"" + $ciresonPortalServer + "SC/ServiceCatalog/RequestOffering/" + $serviceCatalogResult.RequestOfferingId + "," + $serviceCatalogResult.ServiceOfferingId + "`""
                $RequestOfferingURL = "<a href=$ciresonPortalRequestURL/>$($serviceCatalogResult.RequestOfferingTitle)</a><br />"
                $requestOfferingSuggestion = New-Object System.Object
                $requestOfferingSuggestion | Add-Member -type NoteProperty -name RequestOfferingURL -value $RequestOfferingURL
                $requestOfferingSuggestion | Add-Member -type NoteProperty -name WordsMatched -value $wordsMatched
                $matchingRequestURLs += $requestOfferingSuggestion
            }
        }
        $matchingRequestURLs = ($matchingRequestURLs | sort-object WordsMatched -Descending).RequestOfferingURL
        return $matchingRequestURLs
    }
}

#search the Cireson KB based on content from a New Work Item and notify the Affected User
function Search-CiresonKnowledgeBase ($searchQuery, $ciresonPortalUser)
{
    $kbAPIurl = "api/V3/KnowledgeBase/GetHTMLArticlesFullTextSearch?userId=$($ciresonPortalUser.Id)&searchValue=$searchQuery&isManager=$([bool]$ciresonPortalUser.KnowledgeManager)&userLanguageCode=$($ciresonPortalUser.LanguageCode)"
    if ($ciresonPortalWindowsAuth -eq $true)
    {
        $kbResults = Invoke-RestMethod -Uri ($ciresonPortalServer+$kbAPIurl) -UseDefaultCredentials
    }
    else
    {
        $kbResults = Invoke-RestMethod -Uri ($ciresonPortalServer+$kbAPIurl) -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
    }

    $kbResults =  $kbResults | ?{$_.endusercontent -ne ""} | select-object articleid, title
    
    if ($kbResults)
    {
        $matchingKBURLs = @()
        foreach ($kbResult in $kbResults)
        {
            $wordsMatched = ($searchQuery.Split() | ?{($kbResult.title -match "\b$_\b")}).count
            if ($wordsMatched -ge $numberOfWordsToMatchFromEmailToKA)
            {
                $knowledgeSuggestion = New-Object System.Object
                $KnowledgeArticleURL = "<a href=$ciresonPortalServer" + "KnowledgeBase/View/$($kbResult.articleid)#/>$($kbResult.title)</a><br />"
                $knowledgeSuggestion | Add-Member -type NoteProperty -name KnowledgeArticleURL -value $KnowledgeArticleURL
                $knowledgeSuggestion | Add-Member -type NoteProperty -name WordsMatched -value $wordsMatched
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
    [cmdletbinding()]
    Param
    (
        [Parameter()]
        [switch]$SuggestKA,
        [Parameter()]
        [switch]$SuggestRO,
        [Parameter()]
        [switch]$AzureKA,
        [Parameter()]
        [switch]$AzureRO,
        [Parameter()]
        [object]$WorkItem,
        [Parameter()]
        [object]$AffectedUser
    )
    
    #retrieve the cireson portal user
    $portalUser = Get-CiresonPortalUser -username $AffectedUser.UserName -domain $AffectedUser.Domain

    #Define the initial keyword hashtable to use against the Cireson Web API
    $searchQueriesHash = @{"AzureRO" = "$($WorkItem.title.trim()) $($WorkItem.description)"; "AzureKA" = "$($WorkItem.title.trim()) $($WorkItem.description)"}

    #if at least 1 ACS feature is being used, retrieve the keywords from ACS
    if ($AzureKA -or $AzureRO)
    {
        $acsKeywordsToSet = (Get-AzureEmailKeywords -messageToEvaluate "$($WorkItem.title.trim()) $($WorkItem.description)") -join " "
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
        "SuggestKA" {$kbURLs = Search-CiresonKnowledgeBase -searchQuery $($searchQueriesHash["AzureKA"]) -ciresonPortalUser $portalUser}
        "SuggestRO" {$requestURLs = Search-AvailableCiresonPortalOfferings -searchQuery $($searchQueriesHash["AzureRO"]) -ciresonPortalUser $portalUser}
    }
    return $kbURLs, $requestURLs
}

#take suggestion URL arrays returned from Get-CiresonSuggestionURL, load custom HTML templates, and send results back out to the Affected User about their Work Item
function Send-CiresonSuggestionEmail
{
    [cmdletbinding()]
    Param
    (
        [Parameter()]
        [array]$KnowledgeBaseURLs,
        [Parameter()]
        [array]$RequestOfferingURLs,
        [Parameter()]
        [object]$WorkItem,
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
    $emailTemplate = try {$emailTemplate.Replace("{0}", $KnowledgeBaseURLs)} catch {}
    $emailTemplate = try {$emailTemplate.Replace("{1}", $RequestOfferingURLs)} catch {}
    $emailTemplate = try {$emailTemplate.Replace("{2}", $resolveMailTo)} catch {}

    #send the email to the affected user
    Send-EmailFromWorkflowAccount -subject "[$($WorkItem.id)] - $($WorkItem.title)" -body $emailTemplate -bodyType "HTML" -toRecipients $AffectedUserEmailAddress
    
    #if enabled, as part of the Suggested KA or RO process set the First Response Date on the Work Item
    if ($enableSetFirstResponseDateOnSuggestions) {Set-SCSMObject -SMObject $WorkItem -Property FirstResponseDate -Value (get-date).ToUniversalTime() @scsmMGMTParams}
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
    if ($calAppt.ItemClass -eq "IPM.Schedule.Meeting.Request")
    {
        #set the Scheduled Start/End dates on the Work Item
        $scheduledHashTable =  @{"ScheduledStartDate" = $calAppt.StartTime.ToUniversalTime(); "ScheduledEndDate" = $calAppt.EndTime.ToUniversalTime()}    
        
        switch ($wiType)
        {
            "ir" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "sr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "pr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "cr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "rr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}

            #activities
            "ma" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "pa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "sa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "da" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
        }
    }

    #the meeting request is a cancelation, null the values out of the Schedule Start/End Time
    elseif ($calAppt.ItemClass -eq "IPM.Schedule.Meeting.Canceled")
    {
        #set the Scheduled Start/End dates on the Work Item
        $scheduledHashTable =  @{"ScheduledStartDate" = $null; "ScheduledEndDate" = $null}    
        
        switch ($wiType)
        {
            "ir" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "sr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "pr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "cr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "rr" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}

            #activities
            "ma" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "pa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "sa" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
            "da" {Set-SCSMObject -SMObject $workItem -propertyhashtable $scheduledHashTable @scsmMGMTParams}
        }
    }
}

function Verify-WorkItem ($message, $returnWorkItem)
{
    #If emails are being attached to New Work Items, filter on the File Attachment Description that equals the Exchange Conversation ID as defined in the Attach-EmailToWorkItem function
    if ($attachEmailToWorkItem -eq $true)
    {
        $emailAttachmentSearchObject = Get-SCSMObject -Class $fileAttachmentClass -Filter "Description -eq 'ExchangeConversationID:$($message.ConversationID);'" @scsmMGMTParams | select-object -first 1 
        $relatedWorkItemFromAttachmentSearch = Get-SCSMObject -Id (Get-SCSMRelationshipObject -ByTarget $emailAttachmentSearchObject @scsmMGMTParams).sourceobject.id @scsmMGMTParams
        if ($emailAttachmentSearchObject -and $relatedWorkItemFromAttachmentSearch)
        {
            switch ($relatedWorkItemFromAttachmentSearch.ClassName)
            {
                "System.WorkItem.Incident" {Update-WorkItem -message $message -wiType "ir" -workItemID $relatedWorkItemFromAttachmentSearch.id; if($returnWorkItem){return $relatedWorkItemFromAttachmentSearch}}
                "System.WorkItem.ServiceRequest" {Update-WorkItem -message $message -wiType "sr" -workItemID $relatedWorkItemFromAttachmentSearch.id; if($returnWorkItem){return $relatedWorkItemFromAttachmentSearch}}
                "System.WorkItem.ChangeRequest" {Update-WorkItem -message $message -wiType "cr" -workItemID $relatedWorkItemFromAttachmentSearch.id; if($returnWorkItem){return $relatedWorkItemFromAttachmentSearch}}
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

function Read-MIMEMessage ($message)
{
    #Get the Mime Content of the message via MimeKit
    $messageWithMimeContent = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($exchangeService,$message.id,$mimeContentSchema)
    $mimeMessageMemoryStream = New-Object System.IO.MemoryStream($messageWithMimeContent.MimeContent.Content,0,$messageWithMimeContent.MimeContent.Content.Length)
    $parsedMimeMessage = New-Object MimeKit.MimeParser($mimeMessageMemoryStream)

    return $parsedMimeMessage
}

# Get-TemplatesByMailbox returns a hashtable with DefaultWiType, IRTemplate, SRTemplate, PRTemplate, and CRTemplate
function Get-TemplatesByMailbox ($message)
{      
    Write-Debug "To: $($message.To)"
    # There could be more than one addressee--loop through and match to our list
    foreach ($recipient in $message.To) {
        Write-Debug $recipient.Address
        
        # Break on the first match
        if ($Mailboxes[$recipient.Address]) {
            $MailboxToUse = $recipient.Address
            break
        }
    }

    if ($MailboxToUse) {
        Write-Debug "Redirection from known mailbox: $mailboxToUse.  Using custom templates."
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
            Write-Debug "Redirection from known mailbox: $mailboxToUse.  Found in CC field.  Using custom templates."
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
                Write-Debug "Redirection from known mailbox: $mailboxToUse.  Found in Return-Path field.  Using custom templates."
                return $Mailboxes[$MailboxToUse]
            }
            else {
                Write-Debug "No redirection from known mailbox.  Using Default templates"
                return $Mailboxes[$ScsmEmail]
            }
        }
    }
}

function Get-SCSMWorkItemSettings ($WorkItemClass) {   
    switch ($WorkItemClass) {
        "System.WorkItem.Incident" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.WorkItem.Incident.GeneralSetting"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxAttachments
            $maxSize = $settings.MaxAttachmentSize
            $prefix = $settings.PrefixForId
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.ServiceRequest" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ServiceRequestSettings"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.ServiceRequestPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.ChangeRequest" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ChangeSettings"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.SystemWorkItemChangeRequestIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Problem" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ProblemSettings"
            $settings = $settingCls | Get-ScsmObject @scsmMGMTParams
            $maxAttach = $settings.MaxFileAttachmentsCount
            $maxSize = $settings.MaxFileAttachmentSizeinKB
            $prefix = $settings.ProblemIdPrefix
            $prefixRegex = ""
            foreach ($char in $prefix.tochararray()) {$prefixRegex += "[" + $char + "]"}
        }
        "System.WorkItem.Release" {
            $settingCls = Get-ScsmClass @scsmMGMTParams -Name "System.GlobalSetting.ReleaseSettings"
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
function Test-KeywordsFoundInMessage ($message) {
    $found = $false
    #check the subject first
    $found = ($message.subject -match $workItemTypeOverrideKeywords)
    #if necessary, check the body
    if (-Not $found) {
        $found = ($message.body -match $workItemTypeOverrideKeywords)
    }
    return $found
}

#retrieve sender's ability to post announcement based on previously defined email addresses or an AD group
function Get-SCSMAuthorizedAnnouncer ($sender)
{
    switch ($approvedMemberTypeForSCSMAnnouncer)
    {
        "users" {if ($approvedUsersForSCSMAnnouncements -match $sender)
                    {
                        return $true
                    }
                    else
                    {
                        return $false
                    }        
                }
        "group" {$group = Get-ADGroup @adParams -Identity $approvedADGroupForSCSMAnnouncements
                    $adUser = Get-ADUser @adParams -Filter "EmailAddress -eq '$sender'"
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

function Set-CoreSCSMAnnouncement ($message, $workItem)
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
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($lowAnnouncemnentExpirationInHours)}
    }
    elseif ($message.body -match [Regex]::Escape("#$criticalAnnouncemnentPriorityKeyword"))
    {
        #high priority
        $scsmPriorityName = "Critical"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        $announcementBody = $announcementBody -replace "#$criticalAnnouncemnentPriorityKeyword", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($criticalAnnouncemnentExpirationInHours)}
    }
    else
    {
        #normal priority
        $scsmPriorityName = "Medium"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($normalAnnouncemnentExpirationInHours)}
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

function Set-CiresonPortalAnnouncement ($message, $workItem)
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
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($lowAnnouncemnentExpirationInHours)}
    }
    elseif ($message.body -match [Regex]::Escape("#$criticalAnnouncemnentPriorityKeyword"))
    {
        #high priority
        $ciresonPortalPriorityEnum = "F10A51C2-C569-4E64-8237-2B117D63DDB8"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        $announcementBody = $announcementBody -replace "#$criticalAnnouncemnentPriorityKeyword", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($criticalAnnouncemnentExpirationInHours)}
    }
    else
    {
        #normal priority
        $ciresonPortalPriorityEnum = "64096F7F-F8E0-491C-A7FE-94FEDDED4715"
        $announcementBody = $message.Body -replace "\[$announcementKeyword\]", ""
        if ($message.EndTime -eq $null) {$message.EndTime = $message.StartTime.AddHours($normalAnnouncemnentExpirationInHours)}
    }

    #Extract the groups that the message was sent to
    #rename the GroupID property to "AccessGroupID" so as to compare the difference later
    $groupEmails = @()
    $groupEmails += $message.To | ?{$_.MailboxType -ne "Mailbox"}
    $groupEmails += $message.Cc | ?{$_.MailboxType -ne "Mailbox"}
    $portalGroups = @()
    foreach ($groupEmail in $groupEmails)
    {
        $portalGroups += Get-CiresonPortalGroup -groupEmail $groupEmail.Name
    }

    #Get the user that is posting the announcement (from) from the SCSM/Cireson Portal to determine their language code to post the announcement
    $announcerSCSMObject = Get-SCSMUserByEmailAddress -EmailAddress "$($message.from)"
    $ciresonPortalAnnouncer = Get-CiresonPortalUser -username $announcerSCSMObject.username -domain $announcerSCSMObject.domain

    #Get any announcements that already exist for the Work Item
    $allPortalAnnouncements = Get-CiresonPortalAnnouncements -languageCode $ciresonPortalAnnouncer.LanguageCode
    $allPortalAnnouncements = $allPortalAnnouncements | ?{$_.title -match "\[" + $workitem.name + "\]"}

    #determine authentication to use (windows/forms)
    if ($allPortalAnnouncements)
    {
        #### there are announcements to create/update ####

        #### combine the announcement objects and group objects together and group by GroupAccessID, then find object groups that don't have an announcement id ####
        #Announcement array has an AccessGroupID that does not match a group from the message. create announcement for that group
        $groupsToCreateAnnouncements = ($portalGroups + $allPortalAnnouncements) | Group-Object -Property AccessGroupId | ?{$_.Count -eq 1} | Select-Object -Expand Group | ?{$_.Id -eq $null}

        #Announcement array has an AccessGroupID that contains a current group from the message.
        $groupsToUpdateAnnouncements = ($portalGroups + $allPortalAnnouncements) | Group-Object -Property AccessGroupId | Select-Object -Expand Group | ?{$_.Id -ne $null}

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
                $announcementResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials
            }
            else
            {
                $announcementResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
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
                $announcementResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials
            }
            else
            {
                $announcementResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
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
                $announcementResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -UseDefaultCredentials
            }
            else
            {
                $announcementResponse = Invoke-RestMethod -Uri ($ciresonPortalServer+$updateAnnouncementURL) -Method post -Body $announcement -Headers @{"Authorization"=Get-CiresonPortalAPIToken}
            }
        }
    }
    
    # Custom Event Handler
    if ($ceScripts) { Invoke-AfterSetPortalAnnouncement }
}

#modified from Ritesh Modi - https://blogs.msdn.microsoft.com/riteshmodi/2017/03/24/text-analytics-key-phrase-cognitive-services-powershell/
function Get-AzureEmailSentiment ($messageToEvaluate)
{
    #define cognitive services URLs
    $sentimentURI = "https://$azureRegion.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment"

    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "en"; "id" = "1"; "text" = "$messageToEvaluate" };
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final

    #invoke the Cognitive Services Sentiment API
    $sentimentResult = Invoke-RestMethod -Method Post -Uri $sentimentURI -Header @{ "Ocp-Apim-Subscription-Key" = $azureCogSvcTextAnalyticsAPIKey } -Body $messagePayload -ContentType "application/json"
    
    #return the percent score
    return ($sentimentResult.documents.score * 100)
}
function Get-ACSWorkItemPriority ($score, $wiClass)
{  
    #change boundaries as neccesary
    switch ($wiClass)
    {
        #impact/urgency
        "System.WorkItem.Incident" {
            switch ($score)
            {
                {$_ -ge 0 -and $_ -le 20} {$priorityCombo = "hi/hi"}
                {$_ -ge 20 -and $_ -le 60} {$priorityCombo = "med/med"}
                {$_ -ge 60 -and $_ -le 80} {$priorityCombo = "med/low"}
                {$_ -ge 80 -and $_ -le 90} {$priorityCombo = "low/med"}
                {$_ -ge 90 -and $_ -le 100} {$priorityCombo = "low/low"}
            }
        }
        #urgency/priority
        "System.WorkItem.ServiceRequest" {
            switch ($score)
            {
                {$_ -ge 0 -and $_ -le 20} {$priorityCombo = "imm/imm"}
                {$_ -ge 20 -and $_ -le 60} {$priorityCombo = "hi/med"}
                {$_ -ge 60 -and $_ -le 80} {$priorityCombo = "med/low"}
                {$_ -ge 80 -and $_ -le 90} {$priorityCombo = "med/med"}
                {$_ -ge 90 -and $_ -le 100} {$priorityCombo = "low/low"}
            }
        }
    }

    $priorityCalc = @()
    switch ($wiClass)
    {
        "System.WorkItem.Incident" {
            $priorityCalc + $priorityCombo.Split("/")[0].Replace("hi", "System.WorkItem.TroubleTicket.ImpactEnum.High$").Replace("med", "System.WorkItem.TroubleTicket.ImpactEnum.Medium$").Replace("low", "System.WorkItem.TroubleTicket.ImpactEnum.Low$");
            $priorityCalc + $priorityCombo.Split("/")[1].Replace("hi", "System.WorkItem.TroubleTicket.UrgencyEnum.High$").Replace("med", "System.WorkItem.TroubleTicket.UrgencyEnum.Medium$").Replace("low", "System.WorkItem.TroubleTicket.UrgencyEnum.Low$");
        }
        "System.WorkItem.ServiceRequest" {
            $priorityCalc + $priorityCombo.Split("/")[0].Replace("imm", "ServiceRequestUrgencyEnum.Immediate$").Replace("hi", "ServiceRequestUrgencyEnum.High$").Replace("med", "ServiceRequestUrgencyEnum.Medium$").Replace("low", "ServiceRequestUrgencyEnum.Low$");
            $priorityCalc + $priorityCombo.Split("/")[1].Replace("imm", "ServiceRequestPriorityEnum.Immediate$").Replace("hi", "ServiceRequestPriorityEnum.High$").Replace("med", "ServiceRequestPriorityEnum.Medium$").Replace("low", "ServiceRequestPriorityEnum.Low$");
        }
    }
    return $priorityCalc
}
function Get-AzureEmailKeywords ($messageToEvaluate)
{
    #define cognitive services URLs
    $keyPhraseURI = "https://$azureRegion.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases"

    #create the JSON request
    $documents = @()
    $requestHashtable = @{"language" = "en"; "id" = "1"; "text" = "$messageToEvaluate" };
    $documents += $requestHashtable
    $final = @{documents = $documents}
    $messagePayload = ConvertTo-Json $final

    #invoke the Text Analytics Keyword API
    $keywordResult = Invoke-RestMethod -Method Post -Uri $keyPhraseURI -Header @{ "Ocp-Apim-Subscription-Key" = $azureCogSvcTextAnalyticsAPIKey } -Body $messagePayload -ContentType "application/json" 

    #return the keywords
    return $keywordResult.documents.keyPhrases
}
#endregion

#region #### Modified version of Set-SCSMTemplateWithActivities from Morton Meisler seen here http://blog.ctglobalservices.com/service-manager-scsm/mme/set-scsmtemplatewithactivities-powershell-script/
function Update-SCSMPropertyCollection
{
    Param ([Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateObject]$Object =$(throw "Please provide a valid template object")) 
    
    #Regex - Find class from template object property between ! and ']
    $pattern = '(?<=!)[^!]+?(?=''\])'
    if (($Object.Path -match $pattern) -and (($Matches[0].StartsWith("System.WorkItem.Activity")) -or ($Matches[0].StartsWith("Microsoft.SystemCenter.Orchestrator")) -or ($Matches[0].StartsWith("Cireson.Powershell.Activity"))))
    {
        #Set prefix from activity class
        $prefix = (Get-SCSMWorkItemSettings -WorkItemClass $Matches[0])["Prefix"]
       
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
                Update-SCSMPropertyCollection -Object $obj
            }        
        }
    } 
}

function Apply-SCSMTemplate 
{
    Param ([Microsoft.EnterpriseManagement.Common.EnterpriseManagementObjectProjection]$Projection =$(throw "Please provide a valid projection object"), 
    [Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplate]$Template = $(throw 'Please provide an template object, ex. -template template'))
 
    #Get alias from system.workitem.library managementpack to set id property
    $templateMP = $Template.GetManagementPack() 
    $alias = $templateMP.References.GetAlias((Get-SCSMManagementPack system.workitem.library @scsmMGMTParams))
    
    #Update Activities in template
    foreach ($TemplateObject in $Template.ObjectCollection)
    {
        Update-SCSMPropertyCollection -Object $TemplateObject
    }
    #Apply update template
    Set-SCSMObjectTemplate -Projection $Projection -Template $Template -ErrorAction Stop @scsmMGMTParams
}

function Remove-PII ($body)
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

#endregion 

#region #### SCOM Request Functions ####
function Get-SCOMAuthorizedRequester ($sender)
{
    switch ($approvedMemberTypeForSCOM)
    {
        "users" {if ($approvedUsersForSCOM -match $sender)
                    {
                        return $true
                    }
                    else
                    {
                        return $false
                    }        
                }
        "group" {$group = Get-ADGroup @adParams -Identity $approvedADGroupForSCOM
                    $adUser = Get-ADUser @adParams -Filter "EmailAddress -eq '$sender'"
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

function Get-SCOMDistributedAppHealth ($message)
{
    #determine if the sender is authorized to make SCOM Health requests
    $isAuthorized = Get-SCOMAuthorizedRequester $message.From.Address

    if (($isAuthorized -eq $true))
    {
        #find the distributed application to search for based on the [Distributed App Name] from the email body
        #"\[(.*?)\]" - will match something [Service Manager] or [Operations Manager Management Group]
        if ($message.body -match "\[(.*?)\]"){$appName = $Matches[0].Replace("[", "").Replace("]", "")}
        else {<#body not [formed] correctly#>}

        #get Distributed Applications that meet search criteria
        $distributedApps = invoke-command -scriptblock {(Get-SCOMClass | Where-Object {$_.displayname -like "*$appName*"} | Get-SCOMMonitoringObject) | select-object Displayname, healthstate} -ComputerName $scomMGMTServer
        $healthySCOMApps = @()
        $unhealthySCOMApps = @()
        $notMonitoredSCOMApps = @()
        $unhealthySCOMAppsAlerts = @()
        $emailBody = @()

        #create, define, and load custom PS Object from SCOM DA Objects
        foreach ($distributedApp in $distributedApps)
        {
            #Healthy app - Green Agent state
            if ($distributedApp.HealthState -eq "Success")
            {
                $scomDAObject = New-Object System.Object
                $scomDAObject | Add-Member -Type NoteProperty –Name Name –Value $distributedApp.displayname
                $scomDAObject | Add-Member -Type NoteProperty –Name Status –Value "Healthy"
                $healthySCOMApps += $scomDAObject
                $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
            }
            #Unhealthy App - Red Agent state
            elseif ($result.HealthState -eq "Error")
            {
                $scomDAObject = New-Object System.Object
                $scomDAObject | Add-Member -Type NoteProperty –Name Name –Value $distributedApp.displayname
                $scomDAObject | Add-Member -Type NoteProperty –Name Status –Value "Critical"
                $unhealthySCOMApps += $scomDAObject
                $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
            }
            #Gray Agent state
            elseif ($result.HealthState -eq "Uninitialized")
            {
                $scomDAObject = New-Object System.Object
                $scomDAObject | Add-Member -Type NoteProperty –Name Name –Value $distributedApp.displayname
                $scomDAObject | Add-Member -Type NoteProperty –Name Status –Value "Not Monitored"
                $notMonitoredSCOMApps += $scomDAObject
                $emailBody += $scomDAObject.Name + " is " + $scomDAObject.Status + "<br />"
            } 
        }

        #if there are unhealthy apps/red agent states, get their Active alerts in SCOM
        if ($unhealthySCOMApps)
        {
            foreach ($unhealthySCOMApp in $unhealthySCOMApps)
            {
                $unhealthySCOMAppsAlerts += invoke-command -scriptblock {Get-SCOMClass | Where-Object {$_.displayname -like “$($unhealthySCOMApp.displayname)”} | Get-SCOMClassInstance | %{$_.GetRelatedMonitoringObjects()} | Get-SCOMAlert -ResolutionState 0} -computername $scomMGMTServer
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
        if ($certStore -eq "user")
        {
            $certStore = New-Object MimeKit.Cryptography.WindowsSecureMimeContext("CurrentUser")
        }
        else
        {
            $certStore = New-Object MimeKit.Cryptography.WindowsSecureMimeContext("LocalMachine")
        }
    }
    catch
    {
        #decrypting certificate or mimekit couldn't be loaded. Don't process signed/encrypted emails
        $processDigitallySignedMessages = $false
        $processEncryptedMessages = $false
    }
}

#load the Work Item regex patterns before the email processing loop
$irRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Incident")["PrefixRegex"]
$srRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.ServiceRequest")["PrefixRegex"]
$prRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Problem")["PrefixRegex"]
$crRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.ChangeRequest")["PrefixRegex"]
$maRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Activity.ManualActivity")["PrefixRegex"]
$paRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Activity.ParallelActivity")["PrefixRegex"]
$saRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Activity.SequentialActivity")["PrefixRegex"]
$daRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Activity.DependentActivity")["PrefixRegex"]
$raRegex = (get-scsmworkitemsettings -WorkItemClass "System.Workitem.Activity.ReviewActivity")["PrefixRegex"]

# Custom Event Handler
if ($ceScripts) { Invoke-BeforeConnect }

#define Exchange assembly and connect to EWS
[void] [Reflection.Assembly]::LoadFile("$exchangeEWSAPIPath")
$exchangeService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService
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
$inboxFilterString = [scriptblock]::Create("$inboxFilterString")

#filter the inbox
$inbox = $exchangeService.FindItems($inboxFolder.Id,$searchFilter,$itemView) | where-object $inboxFilterString | Sort-Object DateTimeReceived

# Custom Event Handler
if ($ceScripts) { Invoke-OnOpenInbox }

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
        $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessEmail }
        
        switch -Regex ($email.subject) 
        { 
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[$srRegex[0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[$prRegex[0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[$crRegex[0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
            #### activities ####
            "\[$raRegex[0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
            "\[$maRegex[0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

            #### 3rd party classes, work items, etc. add here ####
            "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "([R][E][:])(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

            #### default action, create work item ####
            default {new-workitem $email $defaultNewWorkItem} 
        }
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessEmail }

        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
        $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
    }

    #### Process a Digitally Signed message ####
    elseif ($message.ItemClass -eq "IPM.Note.SMIME.MultipartSigned")
    {
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessSignedEmail }
        
        $response = Read-MIMEMessage $message

        #check to see if there are attachments besides the smime.p7s signature
        $signedAttachments = $response.Attachments
        $signedAttachments = $signedAttachments | ?{$_.filename -ne "smime.p7s"}
   
        $email = New-Object System.Object 
        $email | Add-Member -type NoteProperty -name From -value $response.From.address
        $email | Add-Member -type NoteProperty -name To -value $response.To.Address
        $email | Add-Member -type NoteProperty -name CC -value $response.Cc.Address
        $email | Add-Member -type NoteProperty -name Subject -value $response.Subject
        $email | Add-Member -type NoteProperty -name Attachments -value $signedAttachments
        $email | Add-Member -type NoteProperty -name Body -value $response.TextBody
        $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
        $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
        $email | Add-Member -type NoteProperty -name ID -Value $message.Id
        $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationId
        $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
        $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessEmail }

        switch -Regex ($email.subject) 
        { 
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[$srRegex[0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[$prRegex[0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
            "\[$crRegex[0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
            #### activities ####
            "\[$raRegex[0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
            "\[$maRegex[0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

            #### 3rd party classes, work items, etc. add here ####
            "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

            #### Email is a Reply and does not contain a [Work Item ID]
            # Check if Work Item (Title, Body, Sender, CC, etc.) exists
            # and the user was replying too fast to receive Work Item ID notification
            "([R][E][:])(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

            #### default action, create work item ####
            default {new-workitem $email $defaultNewWorkItem} 
        }
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessEmail }
                
        #mark the message as read on Exchange, move to deleted items
        $message.IsRead = $true
        $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
        $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessSignedEmail }
    }

    #### Process an Encrypted message ####
    elseif ($message.ItemClass -eq "IPM.Note.SMIME")
    {
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }
        
        $response = Read-MIMEMessage $message
        $decryptedBody = $response.Body.Decrypt($certStore)

        #Messaged is encrypted
        if (($response.Body -ne $null) -and ($response.Body.SecureMimeType -eq "EnvelopedData") -and ($decryptedBody.TextBody))
        {         
            #check to see if there are attachments
            $decryptedAttachments = $decryptedBody | ?{$_.isattachment -eq $true}

            $email = New-Object System.Object 
            $email | Add-Member -type NoteProperty -name From -value $response.From.Address
            $email | Add-Member -type NoteProperty -name To -value $response.To.Address
            $email | Add-Member -type NoteProperty -name CC -value $response.Cc.Address
            $email | Add-Member -type NoteProperty -name Subject -value $response.Subject
            $email | Add-Member -type NoteProperty -name Attachments -value $decryptedAttachments
            $email | Add-Member -type NoteProperty -name Body -value $decryptedBody.TextBody
            $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
            $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
            $email | Add-Member -type NoteProperty -name ID -Value $message.Id
            $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationId
            $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
            $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEmail }
            
            switch -Regex ($email.subject) 
            { 
                #### primary work item types ####
                "\[$irRegex[0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[$srRegex[0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[$prRegex[0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[$crRegex[0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
                #### activities ####
                "\[$raRegex[0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
                "\[$maRegex[0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "([R][E][:])(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

                #### default action, create work item ####
                default {new-workitem $email $defaultNewWorkItem} 
            }
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEmail }

            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
            $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }
        }
        #Message is encrypted and signed
        else
        {
            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEncryptedEmail }
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessSignedEmail }
            
            $email = New-Object System.Object 
            $email | Add-Member -type NoteProperty -name From -value $response.From.Address
            $email | Add-Member -type NoteProperty -name To -value $response.To.Address
            $email | Add-Member -type NoteProperty -name CC -value $response.Cc.Address
            $email | Add-Member -type NoteProperty -name Subject -value $response.Subject
            $email | Add-Member -type NoteProperty -name Attachments -value $decryptedAttachments
            $email | Add-Member -type NoteProperty -name Body -value "This message is digitally encrypted and signed. Please see the related/attached item."
            $email | Add-Member -type NoteProperty -name DateTimeSent -Value $message.DateTimeSent
            $email | Add-Member -type NoteProperty -name DateTimeReceived -Value $message.DateTimeReceived
            $email | Add-Member -type NoteProperty -name ID -Value $message.Id
            $email | Add-Member -type NoteProperty -name ConversationID -Value $message.ConversationId
            $email | Add-Member -type NoteProperty -name ConversationTopic -Value $message.ConversationTopic
            $email | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-BeforeProcessEmail }
            
            switch -Regex ($email.subject) 
            { 
                #### primary work item types ####
                "\[$irRegex[0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){update-workitem $email "ir" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[$srRegex[0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){update-workitem $email "sr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[$prRegex[0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){update-workitem $email "pr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
                "\[$crRegex[0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){update-workitem $email "cr" $result.id} else {new-workitem $email $defaultNewWorkItem}}
 
                #### activities ####
                "\[$raRegex[0-9]+\]" {$result = get-workitem $matches[0] $raClass; if ($result){update-workitem $email "ra" $result.id}}
                "\[$maRegex[0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){update-workitem $email "ma" $result.id}}

                #### 3rd party classes, work items, etc. add here ####
                "\[$distributedApplicationHealthKeyword]" {if($enableSCOMIntegration -eq $true){$result = Get-SCOMDistributedAppHealth -message $email; if ($result -eq $false){new-workitem $email $defaultNewWorkItem}}}

                #### Email is a Reply and does not contain a [Work Item ID]
                # Check if Work Item (Title, Body, Sender, CC, etc.) exists
                # and the user was replying too fast to receive Work Item ID notification
                "([R][E][:])(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if($mergeReplies -eq $true){Verify-WorkItem $email} else{new-workitem $email $defaultNewWorkItem}}

                #### default action, create work item ####
                default {new-workitem $email $defaultNewWorkItem} 
            }
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEmail }
            
            #mark the message as read on Exchange, move to deleted items
            $message.IsRead = $true
            $hideInVar01 = $message.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve)
            $hideInVar02 = $message.Move([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::DeletedItems)
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessSignedEmail }
            
            # Custom Event Handler
            if ($ceScripts) { Invoke-AfterProcessEncryptedEmail }
        }
    }

    #Process a Calendar Meeting
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
        $appointment | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessAppointment }
        
        switch -Regex ($appointment.subject) 
        { 
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){schedule-workitem $appointment "ir" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "ir" -workItemID $result.name}}
            "\[$srRegex[0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){schedule-workitem $appointment "sr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "sr" -workItemID $result.name}}
            "\[$prRegex[0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){schedule-workitem $appointment "pr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "pr" -workItemID $result.name}}
            "\[$crRegex[0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){schedule-workitem $appointment "cr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "cr" -workItemID $result.name}}
            "\[$rrRegex[0-9]+\]" {$result = get-workitem $matches[0] $rrClass; if ($result){schedule-workitem $appointment "rr" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "rr" -workItemID $result.name}}

            #### activities ####
            "\[$maRegex[0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){schedule-workitem $appointment "ma" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "ma" -workItemID $result.name}}
            "\[$paRegex[0-9]+\]" {$result = get-workitem $matches[0] $paClass; if ($result){schedule-workitem $appointment "pa" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "pa" -workItemID $result.name}}
            "\[$saRegex[0-9]+\]" {$result = get-workitem $matches[0] $saClass; if ($result){schedule-workitem $appointment "sa" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "sa" -workItemID $result.name}}
            "\[$daRegex[0-9]+\]" {$result = get-workitem $matches[0] $daClass; if ($result){schedule-workitem $appointment "da" $result; $message.Accept($true); Update-WorkItem -message $appointment -wiType "da" -workItemID $result.name}}

            #### 3rd party classes, work items, etc. add here ####

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem $appointment $defaultNewWorkItem $true; schedule-workitem -calAppt $appointment -wiType $defaultNewWorkItem -workItem $returnedNewWorkItemToSchedule ; $message.Accept($true)} 
        }
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessAppointment }
    }

    #Process a Calendar Meeting Cancellation
    elseif ($message.ItemClass -eq "IPM.Schedule.Meeting.Canceled")
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
        $appointment | Add-Member -type NoteProperty -name ItemClass -Value $message.ItemClass
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-BeforeProcessCancelMeeting }

        switch -Regex ($appointment.subject) 
        { 
            #### primary work item types ####
            "\[$irRegex[0-9]+\]" {$result = get-workitem $matches[0] $irClass; if ($result){schedule-workitem $appointment "ir" $result; Update-WorkItem -message $appointment -wiType "ir" -workItemID $result.name}}
            "\[$srRegex[0-9]+\]" {$result = get-workitem $matches[0] $srClass; if ($result){schedule-workitem $appointment "sr" $result; Update-WorkItem -message $appointment -wiType "sr" -workItemID $result.name}}
            "\[$prRegex[0-9]+\]" {$result = get-workitem $matches[0] $prClass; if ($result){schedule-workitem $appointment "pr" $result; Update-WorkItem -message $appointment -wiType "pr" -workItemID $result.name}}
            "\[$crRegex[0-9]+\]" {$result = get-workitem $matches[0] $crClass; if ($result){schedule-workitem $appointment "cr" $result; Update-WorkItem -message $appointment -wiType "cr" -workItemID $result.name}}
            "\[$rrRegex[0-9]+\]" {$result = get-workitem $matches[0] $rrClass; if ($result){schedule-workitem $appointment "rr" $result; Update-WorkItem -message $appointment -wiType "rr" -workItemID $result.name}}

            #### activities ####
            "\[$maRegex[0-9]+\]" {$result = get-workitem $matches[0] $maClass; if ($result){schedule-workitem $appointment "ma" $result; Update-WorkItem -message $appointment -wiType "ma" -workItemID $result.name}}
            "\[$paRegex[0-9]+\]" {$result = get-workitem $matches[0] $paClass; if ($result){schedule-workitem $appointment "pa" $result; Update-WorkItem -message $appointment -wiType "pa" -workItemID $result.name}}
            "\[$saRegex[0-9]+\]" {$result = get-workitem $matches[0] $saClass; if ($result){schedule-workitem $appointment "sa" $result; Update-WorkItem -message $appointment -wiType "sa" -workItemID $result.name}}
            "\[$daRegex[0-9]+\]" {$result = get-workitem $matches[0] $daClass; if ($result){schedule-workitem $appointment "da" $result; Update-WorkItem -message $appointment -wiType "da" -workItemID $result.name}}

            #### 3rd party classes, work items, etc. add here ####
            "([C][a][n][c][e][l][e][d][:])(?!.*\[(($irRegex)|($srRegex)|($prRegex)|($crRegex)|($maRegex)|($raRegex))[0-9]+\])(.+)" {if($mergeReplies -eq $true){$result = Verify-WorkItem $appointment -returnWorkItem $true; schedule-workitem $appointment $defaultNewWorkItem $result} else{new-workitem $appointment $defaultNewWorkItem}}

            #### default action, create/schedule a new default work item ####
            default {$returnedNewWorkItemToSchedule = new-workitem $appointment $defaultNewWorkItem $true; schedule-workitem -calAppt $appointment -wiType $defaultNewWorkItem -workItem $returnedNewWorkItemToSchedule ; $message.Accept($true)} 
        }
        
        # Custom Event Handler
        if ($ceScripts) { Invoke-AfterProcessCancelMeeting }
        
        #Move to deleted items
        $message.Delete([Microsoft.Exchange.WebServices.Data.DeleteMode]::MoveToDeletedItems)
    }
}
