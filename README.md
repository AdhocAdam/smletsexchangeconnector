<p align="center">
  <img width="512" height="512" src="psExchIconReadme.png">
</p>

# SCSM Exchange Connector via SMlets
<p align="left">
  <img src="https://github.com/AdhocAdam/smletsexchangeconnector/workflows/Primary%20Branch%20-%20Build/badge.svg?branch=primary" />
  <img src="https://github.com/AdhocAdam/smletsexchangeconnector/workflows/Dev%20Branch%20-%20Build/badge.svg?branch=dev" />
</p>
<p align="left">
  <a href="https://securitylab.github.com/tools/codeql"><img src="https://github.com/AdhocAdam/smletsexchangeconnector/actions/workflows/codeql-analysis.yml/badge.svg" /></a>
</p>

This PowerShell script leverages the [SMlets module](https://www.powershellgallery.com/packages/SMLets/) to build an open and flexible Exchange Connector for controlling Microsoft System Center Service Manager 2016+


## So what is this for?
The stock Exchange Connector is a seperate download that enables SCSM deployments to leverage an Exchange mailbox to process updates to work items. While incredibly useful, some feel limited by its inability to be customized given its nature as a sealed management pack. This PowerShell script replicates all functionality of [Exchange Connector 4.0](https://www.microsoft.com/en-us/download/details.aspx?id=101579), introduces a host of new features, and most importantly enables SCSM Administrators to customize the solution to their needs.

## Who is this for?
This is aimed at SCSM administrators looking to further push the automation limits of what their SCSM deployment can do with inbound email processing. As such, you should be comfortable with PowerShell and navigating SCSM via SMlets.

## What new things can it do?
<table border="0">
  <tr>
    <td colspan="3"><i>Custom Rules (v4.0)</i></td>
  </tr>
  <tr>
    <td align="center" width="200"><img src ="/FeatureScreenshots/customRules.png" /></td>
    <td width="auto">Sync any external system with Service Manager entirely through email. Whether it's a vendor's ticketing system, Microsoft Azure Invoices, an alerting tool, a CRM that sends notifications, or even a project management tool. You can now manage all of it through Work Items or Config Items of any class. Looking for inspiration? Check out the <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Integrating-External-Ticketing-Systems">wiki</a>.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>File Attachments for Activities (v3.4)</i></td>
  </tr>
  <tr>
    <td align="center" width="200"><img src ="/FeatureScreenshots/activityFileAttachments.png" /></td>
    <td width="auto">Whether you choose to Complete, Fail, Approve, Reject, or just leave notes on an Activity. Emails for Activities that have File Attachments, will now make their way to the parent Work Item.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>History Pane (v3.3)</i></td>
  </tr>
  <tr>
    <td width="50%"><img src ="/FeatureScreenshots/history.png" /></td>
    <td width="auto">Whether you're changing regex patterns, updating machine learning scores, or modifying any other configuration in the connector. Now just like any Work Item or Config Item you can view the complete history of changes you and your peers have set.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>[pwsh] Keyword (v3.1)</i></td>
  </tr>
  <tr>
    <td align="center" width="200"><img src ="/FeatureScreenshots/pwshKeyword.png" /></td>
    <td width="auto">It's now possible with digital signatures and custom events to call Rest APIs, invoke Webhooks, and execute PowerShell <i>directly</i> from email. This feature makes use of Mimekit to verify the sender and then calls your CustomEvents file to take further action. Looking for ideas to get started? Head over to the <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Custom-Events-Examples#invoke-validdigitalsignatureaction">wiki</a>.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>Workflow Engine (v3.0)</i></td>
  </tr>
  <tr>
    <td align="center" width="200"><img src ="/FeatureScreenshots/workflow.png" /></td>
    <td width="auto">The connector can optionally can be driven through the Service Manager workflow engine and leverage Run as Accounts. All while maintaining the customization and flexibility you've come to expect through a PowerShell based connector. Setup instructions can be found <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Configuration-Examples#scenario-4-native-service-manager">here</a>.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>OAuth 2.0 for Exchange Online (v2.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/OAuth.png" /></td>
    <td width="auto">The connector supports OAuth 2.0 authentication via Azure AD Application registration to communicate with Exchange Online. Need help getting setup? The <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Configuration-Examples#scenario-3-exchange-onlineoffice-365-with-oauth">wiki</a> has your covered.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>Transcribe Audio Files to Text (v2.3)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/speech-api.png" /></td>
    <td width="auto">Your SCSM environment can now hear! Translate incoming *.wav or *.ogg on emails to text with Azure Speech services. After enabling audio files will have their text transcriptions written to the File Attachment's Description. Get pricing details <a href="https://azure.microsoft.com/en-us/services/cognitive-services/speech-services/">here</a>.
</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>Describe and Extract Text from Images (v2.3)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/ACSVision.png" /></td>
    <td width="auto">Give your SCSM environment eyes with integration into Azure Cognitive Services Vision API. Once enabled, any image attached to emails will be submitted to Azure and will be tagged with 5 words describing what the image is. In the event text is detected, Optical Character Recognition (OCR) will additionally be performed. These results are then written into the Description property of the attachment. Get pricing details <a href="https://azure.microsoft.com/en-us/services/cognitive-services/computer-vision/">here</a>.
</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>Predict Impacted Configuration Items with Azure Machine Learning (v2.2)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/machineLearningConfigItem.png" /></td>
    <td width="auto">Take Azure Machine Learning even further and enable the connector the ability to relate Impacted Configuration Items to a New Work Item in addition to Work Item Decisioning, Classification, and Routing. This feature is compatible with Configuration Items of <i>any</i> type whether it be:
<br /><br />
<ul>
  <li>stock Service Manager classes</li>
  <li>imported Operations Manager classes</li>
  <li>ANY 3rd party management pack classes</li>
  <li>or even your own custom in house developed classes</li>
</ul>
Ready to go? Get started over <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Azure-Machine-Learning-setup">here</a>.</td>
  </tr>
</table>
<table border="0">
  <tr>
    <td colspan="3"><i>Support for Cireson Watchlist (v2.1)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/watchlist.png" /></td>
    <td width="auto">Cireson customers can now use the new configurable [watch] and [stopwatch] keywords to control the watchlist directly from email. You can configure HTML links within your own notifications so these keywords function similiar to Announcements, Take, Review Activities, and more. Looking for <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/HTML-Email-Keyword-Examples">examples</a>?</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Settings Management Pack (v2.0)</i></td>
  </tr>
  <tr>
    <td width="50%"><img src ="/docs/images/CreateACustomMPSettingsPane/headerGraphicThumbnail.png" /></td>
    <td width="auto">The SMLets Exchange Connector features over 120+ configuration possibilities. With a brand new management pack, upgrades just easier! This new MP moves the configuration from the PowerShell script into a dedicated MP whose settings are preserved through upgrades. Interested to see how it was built? Check out the blog here on the repo. <a href="https://adhocadam.github.io/smletsexchangeconnector/Building-a-Custom-Settings-UI-for-Service-Manager-(part-1)/">Part 1</a> and <a href="https://adhocadam.github.io/smletsexchangeconnector/Building-a-Custom-Settings-UI-for-Service-Manager-(part-2)/">Part 2 </a></td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Language Translation for New Work Items with Azure Translation (v1.6)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/translate.png" /></td>
    <td width="auto">Remove the barriers to communication from Affected Users with integration to Azure Translate. When enabled, New Work Items will see the immediate creation of an End User Comment on the Action Log in a default language of your choosing. Ready to close the loop so the Assigned User can work with the Affected User in their native language? A new complimentary project has arrived so Comments left on the Action Log area translated, and then the other party is notified of the translated comment. Get the Advanced Action Log Notify management pack over <a href="https://github.com/AdhocAdam/scsmtranslate">here</a>. Read up on Azure Translate pricing <a href="https://azure.microsoft.com/en-us/pricing/details/cognitive-services/translator-text-api/">here</a> and view the list of supported languages <a href="https://docs.microsoft.com/en-us/azure/cognitive-services/translator/language-support">here</a></td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Intelligent Routing and Classification with Azure Machine Learning (v1.5)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/machineLearningClassification.png" /></td>
    <td width="auto">Further compliment "Dynamic Work Item Analyst Assignment" and "Azure Cognitive Services" features with integration for Azure Machine Learning. Using AML, you can now train Azure's artificial intelligence on your SCSM environment using your very own SCSM Data Warehouse to autonomously choose Work Item Type (Incident or Service Request), Work Item Classification, and Work Item Support Group on brand new Work Items that the connector creates. Azure pricing can be found <a href="https://azure.microsoft.com/en-us/pricing/details/machine-learning-studio/">here</a> and how to configure it for your environment/connector can be found <a href="https://github.com/AdhocAdam/smletsexchangeconnector/wiki/Azure-Machine-Learning-setup">here</a>.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Dynamic Work Item Analyst Assignment on New Work Item (v1.5)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/Assign.png" /></td>
    <td width="auto">When New Work Items are created, you can optionally enable the connector to assign them out to Analysts contained within the Default Template's Support Group. This can be done through several means - completely random, least assigned open Work Items, or in conjunction with the <a href="https://github.com/AdhocAdam/scsmoutofoffice">SCSM Out of Office Management pack</a> to prevent assignment when Analysts are out of the office. <i>(Requires at least the free Cireson SCSM Portal to establish the relationship between Support Groups and Analysts)</i></td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Change Incident Status based on Who Replied (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/changeIncidentSstatusOnReply.png" /></td>
    <td width="auto">You can now optionally enable this feature that changes the Status of Incidents when the Affected User, Assigned To, or Related User replies. This helps better reflect time within a specific status as Analysts flip to Pending and potential Pending subcategories. As an example, this means you can configure the connector so that the Assigned To updates it to Pending and the Affected User updates it to Active.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Dynamic Priority based on Azure Sentiment Analysis (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/ACSPriorityScoring.png" /></td>
    <td width="auto">If you're making use of Sentiment Analysis from Azure Cognitive Services, the returned score can now optionally dynamically drive the Impact, Urgency, and Priority on new Incidents or Service Requests. Sentiment boundaries can be edited via the Get-ACSWorkItemPriority function.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Custom Theming for Suggested Knowledge Article or Request Offering emails (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/HTMLSuggestionTemplate.png" /></td>
    <td width="auto">Now when emails go out suggesting Knowledge Articles or Request Offerings based on the email's subject and body, you can theme them to match your organization's respective style. Customizations made to these HTML templates will be preserved throughout updates of the connector.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Redact Sensitive Information (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/redacted.png" /></td>
    <td width="auto">Using a custom defined regex file you can optionally have the connector replace words of your regex choice with [redacted] in New Work Item Descriptions or Comments. Customizations made to this regex list will be preserved throughout updates of the connector.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Enhanced Action Logging (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/enhancedActionlog.png" /></td>
    <td width="auto">Now when using the [resolved], [closed], [take], and other relevant keywords the Work Item's Action Log matches the visual style of these actions as opposed to simply leaving an Analyst Comment.</td>
  </tr>
</table>

*More Keyword functionality for Manual Activites (v1.4)*
- You can now use the [take] keyword against Manual Activites.

*Set First Response Date on Request Offering/Knowledge Article Suggestion (v1.4)*
- If you've enabled Suggested Request Offerings from your Service Catalog or Knowledge Articles from your Knowledge Base, this feature can now optionally set the First Response Date when the connector sends suggestions back to the Affected User giving you more data points to report against.

<table border="0">
  <tr>
    <td colspan="3"><i>Microsoft Azure Cognitive Services integration (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="/FeatureScreenshots/ACSBrain.png" /></td>
    <td width="auto">Ever wish you could create an Incident or Service Request based on the nature of the email? How about using the suggested Knowledge Article/Request Offering feature of this connector, but achieve even faster processing times? Leveraging the power of Azure Cognitive Services emails can now be optionally run through Keyword Analysis for more concise search results and/or Sentiment Analysis for dynamic creation of an Incident or Service Request based on the Affected User's perceived attitude. <a href="https://azure.microsoft.com/en-us/pricing/details/cognitive-services/text-analytics/">Microsoft pricing information can be found here.</a> Not ready for Azure? You can alternatively define a list of keywords that will control whether or not an Incident or Service Request is created.
    </td>
  </tr>
</table>

*Create Problems or Change Requests by Default (v1.4)*
- Whether you're dealing with external vendors or streamlining process, you can now configure the connector to create either Problems or Change Requests by default.

*Process multiple mailboxes with unique templates and default ticket types (v1.4)*
- Redirect multiple mailboxes to a single inbox and this connector can process each message based on the default work item type and templates for the original mailbox, as though it was directly connected to each. This can be helpful for multiple teams that each have their own mailbox, for example, but does not require many instances of the connector running against all mailboxes as the MS Exchange Connector did. In addition, it can also process those mailboxes should be they be featured on the CC or BCC lines.

*Generate a new, related work item when comments are emailed to closed tickets (v1.4)*
- Was the affected user on vacation? Did they try to add to an old incident ticket for a new instance of the same issue? Now you can optionally generate Work Items when a user comments on a closed ticket. If enabled, a new work item will be generated that is related to the old work item and contains many of the same properties, but also contains the new details from their email message.

*More Keyword functionality for Service Requests (v1.4)*
- Service Requests can now use the [acknowledge] and [hold] keywords.

*[take] Keyword can now optionally test for membership in the Support Group (v1.4)*
- If this feature is enabled, the sender's membership in the current Support Group will be tested before they are assigned to a ticket (Incident, Service Request. Manual Activity and Change Request require a class extension). This requires the Cireson Analyst Portal and in turn Cireson Portal Group Mapping in order to establish relationship between Support Groups/Analysts.

<table border="0">
  <tr>
    <td colspan="3"><i>Vote on Behalf of Active Directory Group (v1.4)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/voteOnBehalfOf.png" /></td>
    <td width="auto">The connector can now optionally process Active Directory Groups featured on Review Activities. Now, when a member of the group either approves or rejects the vote, their vote will be counted on behalf of the group through the &quot;Voted By&quot; relationship and their comment appended to the vote. </td>
  </tr>
</table>

*File Attachment Limits (v1.4)*
- Now you can optionally enforce settings for attachments per Work Item type as seen in the Administration -> Settings pane for maximum size and attachment count

<table border="0">
  <tr>
    <td colspan="3"><i>Announcement Integration with Core SCSM and Cireson Portal (v1.3)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/announcements.png" /></td>
    <td width="auto">Using the configuable [announcement] keyword you can post announcements to core SCSM and/or the Cireson portal when this keyword is featured in the body of your message. The priority can be controlled through the optional and configurable #low or #high tags. Absence of these optional #hashtags will default to a normal priority announcement. If you're creating announcements in the Cireson Portal, an announcement will be created for each group discovered on the message's To or CC lines. Thus allowing you to email several distro groups and simultaneously create targeted announcements for them. Announcements created through either means (SCSM/Cireson Portal) can be updated in the future by simply keeping the [Work Item] in the subject of your [announcement] based messages. Finally, announcements can only be created when the Sender's email is in an allowed style list or the Sender is part of a configurable Active Directory group.</td>
  </tr>
</table>

*Search Cireson Portal Service Catalog (v1.3)*
- When enabled this feature will email the Affected User suggested Cireson Portal Request Offerings within the Affected User's Service Catalog permission scope based on the content of their email when a New Work Item is generated. This feature has been tested and confirmed working with v7.x and v8.x of their SCSM portal. If both this and the "Search Cireson HTML Knowledge Base" feature are enabled, only a single summarized email will be sent to the Affected User containing links to Knowledge Articles and Request Offerings.

<table border="0">
  <tr>
    <td colspan="3"><i>Send Outlook Meeting (v1.3)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/sendOutlookMeeting.png" /></td>
    <td width="auto">Building upon the previous version's "Schedule Work Item", this feature introduces a Cireson SCSM Portal Task (via CustomSpace/custom.js) that kicks open your local Outlook client to send meeting requests on Work Items to the Affected User and your workflow account (or just your workflow if no Affected User is present). In doing so, further rounding out the Schedule Work Item feature by setting Scheduled Start/End Dates when the connector processes Calendar items.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Digitally Signed and Encrypted Emails (v1.3)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/digitalCertificates.png" /></td>
    <td width="auto">Leveraging the open source <A href="https://github.com/jstedfast/MimeKit">MimeKit</A> project by Jeffrey Stedfast, the connector can now process digitally signed or encrypted emails just like regular mail. This requires an appropriate certificate in either the user's personal cert store or the local machine's personal cert store.</td>
  </tr>
</table>

*System Center Operations Manager (SCOM) Integration (v1.3)*
- Using the configurable [health] keyword in the subject, authorized users (as defined individually or through an Active Directory group) can request the overall Health and Alert counts of a [Distributed Application] by placing it in the body of the email. Unauthorized users will have New Default Work Items created within SCSM.

*Merge Replies from Related Users instead of Creating New Default Work Items (v1.2)*
- If a user emails the SCSM Workflow Account and also adds additional users to the To/CC lines those related users are automatically added to the Related Items tab of a New Work Item. However in these scenarios, it's possible that one or more of those users could reply within the same processing loop of the Exchange Connector. As a result, they will queue more emails to be turned into New Default Work Items. This feature aims to address the scenario by querying Exchange Inbox/Deleted Items for matching Conversation Topics and ConversationIDs, finding the original item in the thread, searching for the Work Item that already exists in SCSM, and then appending their Reply to its Action Log

<table border="0">
  <tr>
    <td colspan="3"><i>Schedule Work Items (v1.2)</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/outlookMeeting.png" /></td>
    <td width="auto">It's now possible to interact with SCSM via Outlook Calendar Meetings! When a Calendar Meeting is sent, the Scheduled Start Date and Scheduled End Date will be set on the Work Item based on the start/end times of the meeting. If the work item cannot be found/does not exist, a new default work item is created and it's scheduled start/end times set accordingly. Upon success, the meeting will be accepted onto the workflow account's calendar and the requester will receive confirmation of the booking. This introduces the possibility of leveraging the workflow's calendar as a central place to see all Scheduled Work Items. Using <a href="http://cireson.com/apps/outlook-console/">Cireson's Outlook Plugin</a>. When setting a Work Item reminder in Outlook, you can now CC your workflow account to update values on the Work Item.</td>
  </tr>
</table>

*Minimum File Attachment Size*
- You can set a minimum size in KB. In doing so, files less than the defined size will not be added to the work item (i.e. corporate signature graphics won't be added)

<table border="0">
  <tr>
    <td colspan="3"><i>File Attachment &quot;Attached by&quot;</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/attachedBy.png" /></td>
    <td width="auto">When an email is sent with attachments, the &quot;File Attachment Added By User&quot; relationship will be set based on the Sender if the user is found in the CMDB</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Incident, Service Request, Change Request, Problem</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/take.png" /></td>
    <td width="auto">[Take] - When emailing your workflow account, it will assign the Incident, Service Request, Change Request, or Problem to you (from address) when this keyword is featured in the body of the email.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Service Request and Change Request</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/cancel.png" /></td>
    <td width="auto">[Cancel] - Cancel the Service Request or Change Request when this keyword is featured in the body of the email</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Incident</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/reactivate.png" /></td>
    <td width="auto">[Reactivate] - When submitted to a Resolved Incident, it will be reactivated. When submitted to a Closed Incident, a New Incident will be created and the two related to one another.</td>
  </tr>
</table>

<table border="0">
  <tr>
    <td colspan="3"><i>Change Request</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/hold.png" /></td>
    <td width="auto">[Hold] - Place the Change Request On-Hold when this keyword is featured in the body of the email</td>
  </tr>
</table>

*Manual Activity*
- [Skipped] - Skip the activity when this keyword is featured in the body of the email
- Misc - Anyone who is not the implementer will have their email appended to the "Notes" area of the MA
- Misc - If the Implementer leaves a comment that is not [Skipped] or [Completed] the comment is added to the highest level Parent Work Item

*Review Activity*
- Any reviewer who leaves a comment that doesn't contain [Approved] or [Rejected] will have their comment added to the highest level Parent Work Item. This addresses a scenario where users not familiar with SCSM (i.e. departments outside of IT) respond back to the email thinking someone is reading the message on your workflow account. Now their comments aren't simply lost, but instead given the visibility they deserve!

<table border="0">
  <tr>
    <td colspan="3"><i>Incident and Service Request</i></td>
  </tr>
  <tr>
    <td width="200"><img src ="https://github.com/AdhocAdam/smletsExchangeConnector/blob/primary/FeatureScreenshots/privateComment.png" /></td>
    <td width="auto">#private - When the message is attached to the Action Log, it will be marked as private if #private is featured in the body of the message. This keyword is configurable.</td>
  </tr>
</table>

*Assigned To/Affected User relationships on the Action Log*
- When someone who isn't the Assigned To/Affected User leaves a comment on the Action Log the comment's "IsPrivate" flag is marked as null (this is a bug in the EC v3.0 and v3.1 that has yet to be addressed by Microsoft). As such Cireson's Action Log Notify has no qualifier to go of off. With this script, the same functionality is present but now can be altered to get in line with SCSM and Cireson's MP.

*Search Cireson HTML Knowledge Base*
- If you're a customer of Cireson and this feature is enabled, your respective Cireson Portal HTML KB will be searched when a New Work item is generated using its title and description. The Sender will be sent a summarized HTML email with links directly to those knowledge articles about their recently created Work Item using the Exchange EWS API defined therein. As an example email, I've included an email body that features a [Resolved] and [Cancelled] link should the Affected User wish to mark their Incident/Service Request accordingly in the event the KB addresses their request. It should be noted, this is using the Cireson Web API to get KB through a now deprecated function. While this works, it goes without saying if Cireson drops this in coming versions it would cease to work. It has been tested and confirmed working with v7.x of their SCSM HTML KB portal.
