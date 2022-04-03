**Describe the bug**
_e.g. Dynamic Analyst Assignment by Volume is selecting the Analyst with the most assigned Work Items_

**Help us reproduce the bug**
_You can export your configuration to CSV with the following PowerShell function and then upload to this issue. Environment specific and/or identifiable information such as Management Server names, API keys, Exchange URLs, Azure GUIDs are automatically excluded._
```powershell
function Get-SMExcoConfiguration ($computername)
{
    $mpClass = Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings$" -computername $computername
    $propertiesToExclude = "SCSMmgmtServer", "workflowemailaddress", "exchangeautodiscoverurl", "azureclientid", "azuretenantid", "securereferenceidews", "AzureCloudInstance",
        "ciresonportalurl", "SecureReferenceIdCiresonPortal",
        "SCSMAnnouncementApprovedMemberType", "SCSMApprovedAnnouncementGroupGUID", "SCSMApprovedAnnouncementGroupDisplayName", "SCSMApprovedAnnouncementUsers",
        "SCOMApprovedMemberType", "SCOMApprovedGroupGUID", "SCOMApprovedUsers", "SCOMmgmtServer",
        "ACSTextAnalyticsAPIKey", "AMLAPIKey", "ACSSpeechAPIKey", "ACSVisionAPIKey", "ACSTranslateAPIKey", "AMLurl",
        "DisplayName", "__InternalId", "classname", "typename", "name", "path", "fullname", "ManagementPackClassIds", "LeastDerivedNonAbstractManagementPackClassId", "TimeAdded", "LastModifiedBy", "Values", "IsNew", "haschanges", "id", "managementgroup", "managementgroupid", "groupsasdifferenttype", "viewname", "objectmode"
    $mpClass | Get-SCSMObject -computername $computername | Select-Object * -ExcludeProperty $propertiesToExclude
}

#### view the configuration in PowerShell to ensure no sensitive data is included
Get-SMExcoConfiguration -computername localhost

#### export the configuration to CSV
Get-SMExcoConfiguration -computername localhost  | Export-CSV C:\temp\SMExcoConfiguration.csv -NoTypeInformation
```

**Expected behavior**
_e.g. It should be pick the Analyst with the least amount of assigned Work Items_

**Location and Environment**
_e.g. Running v4 of the connector with SCSM 2019 UR2 via workflows_

**Additional context**
_Add any other context or details you wish to provide here_
