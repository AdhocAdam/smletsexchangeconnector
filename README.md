
# SCSM Exchange Connector via SMlets
This PowerShell script leverages the [SMlets module](https://www.powershellgallery.com/packages/smlets/0.5.0.1) to build an open  Exchange Connector for controlling Microsoft System Center Service Manager 2012+


## So what is this for?
The stock Exchange Connector is a seperate download for SCSM 2012+ that enables SCSM deployments to leverage an Exchange mailbox to process updates to work items. While incredibly useful, some feel limited by the lack of customization given its nature as a sealed management pack. This PowerShell script replicates all functionality of [Exchange Connector 3.1](https://www.microsoft.com/en-ca/download/details.aspx?id=45291) introduces several new features, and enables SCSM Administrators to further customize the solution to their needs.

## What new things can it do?
Minimum File Attachment Size
- You can set a minimum size in KB. In doing so, files less than the defined size will not be added to the work item (i.e. corporate signature graphics won't be added)

File Attachment "Added by"
- When an email is sent with attachments, the "File Attachment Added By User" relationship will be set based on the Sender if the user is found in the CMDB

Incident, Service Request, Change Request, Problem
- [Take] - When emailing your workflow account, it will assign the Incident, Service Request, Change Request, or Problem to you (from address) when this keyword is featured in the body of the email.

Incident
- [Reactivate] - When submitted to a Resolved Incident, it will be reactivated. When submitted to a Closed Incident, a New Incident will be created and the two related to one another.

Change Request
- [Hold] - Place the Change Request On-Hold when this keyword is featured in the body of the email
- [Cancel] - Cancel the Change Request when this keyword is featured in the body of the email

Manual Activity
- [Skipped] - Skip the activity when this keyword is featured in the body of the email
- Misc - Anyone who is not the implementer will have their email appended to the "Notes" area of the MA
- Misc - If the Implementer leaves a comment that is not [Skipped] or [Completed] the comment is added to the highest level Parent Work Item

Review Activity
- Any reviewer who leaves a comment that doesn't contain [Approved] or [Rejected] will have their comment added to the highest level Parent Work Item. This addresses a scenario where users not familiar with SCSM (i.e. departments outside of IT) respond back to the email thinking someone is reading the message on your workflow account. Now their comments aren't simply lost, but instead given the visibility they deserve!

Incident and Service Request
- #private - When the message is attached to the action log, it will be marked as private if #private is featured in the body of the message.

Assigned To/Affected User relationships on the Action Log
- When someone who isn't the Assigned To/Affected User leaves a comment on the Action Log the comment's "IsPrivate" flag is marked as null (this is a bug in the EC v3.0 and v3.1 that has yet to be addressed by Microsoft). As such Cireson's Action Log Notify has no qualifier to go of off. With this script, the same functionality is present but now can be altered to get in line with SCSM and Cireson's MP.

Search Cireson HTML Knowledge Base
- If enabled, your respective Cireson Portal HTML KB will be searched when a New Work item is generated using its title and description. The Sender will be sent a summarized HTML email with links directly to those knowledge articles about their recently created Work Item using the Exchange EWS API defined therein. As an example email, I've included an email body that features a [Resolved] and [Cancelled] link should the Affected User wish to mark their Incident/Service Request accordingly in the event the KB addresses their request. It should be noted, this is using the Cireson Web API to get KB through a now deprecated function. While this works, it goes without saying if Cireson drops this in coming versions it would cease to work. It has been tested and confirmed working with v7.x of their SCSM HTML KB portal.
