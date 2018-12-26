using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.EnterpriseManagement.UI.WpfWizardFramework;
using System.ComponentModel;
using Microsoft.EnterpriseManagement.UI.DataModel;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Configuration;
using Microsoft.Win32;


namespace SMLetsExchangeConnectorSettingsUI
{
    class AdminSettingWizardData : WizardData
    {
        #region Variables
        private String strscsmMGMTServer = String.Empty;
        private String strWorkflowEmailAddress = String.Empty;
        private Boolean boolEnableAutodiscover = false;
        private String strAutodiscoverURL = String.Empty;
        private Boolean boolCreateUsersInCMDB = false;
        private Boolean boolIncludeWholeEmail = false;
        private Boolean boolAttachEmailToWorkItem = false;
        private Boolean boolMaxFileSizeEnabled = false;
        private Boolean boolEnableCiresonIntegration = false;
        private Boolean boolEnableCiresonKBSearch = false;
        private Boolean boolEnableCiresonROSearch = false;

        private String strEWSFilePath = String.Empty;
        private String strMimeKitFilePath = String.Empty;

        private String strIRTemplateName = String.Empty;
        private String strSRTemplateName = String.Empty;
        private String strCRTemplateName = String.Empty;
        private String strPRTemplateName = String.Empty;

        private String strFromKW = String.Empty;
        private String strAcknowledgeKW = String.Empty;
        private String strReactivateKW = String.Empty;
        private String strResolvedKW = String.Empty;
        private String strClosedKW = String.Empty;
        private String strHoldKW = String.Empty;
        private String strTakeKW = String.Empty;
        private String strCancelledKW = String.Empty;
        private String strCompletedKW = String.Empty;
        private String strSkippedKW = String.Empty;
        private String strApproveKW = String.Empty;
        private String strRejectKW = String.Empty;

        private String strCiresonPortalURL = String.Empty;
        private String strMinWordsToMatchToSuggestRO = String.Empty;
        private Boolean boolEnableAzureCognitiveServices = false;
        private Boolean boolEnableAnnouncementIntegration = false;
        private Boolean boolEnableSCSMAnnouncements = false;
        private Boolean boolEnableCiresonAnnouncements = false;

        private String strAnnouncementKW = String.Empty;
        private String strAuthorizedAnnouncer = String.Empty;
        private String strLowPriorityAnnouncementKW = String.Empty;
        private String strMedPriorityAnnouncementKW = String.Empty;
        private String strHighPriorityAnnouncementKW = String.Empty;
        private String strLowPriorityExpInHours = String.Empty;
        private String strMedPriorityExpInHours = String.Empty;
        private String strHighPriorityExpInHours = String.Empty;

        private Boolean boolProcessCalendarAppointments = false;
        private Boolean boolProcessMergeReplies = false;
        private Boolean boolProcessClosedWorkItemsToNewWorkItems = false;
        private Boolean boolProcessEncryptedEmails = false;
        private Boolean boolProcessDigitallySignedEmails = false;

        private String strDefaultWorkItem = String.Empty;
        private String strMinFileSize = String.Empty;

        private Boolean boolEnableSCOMIntegration = false;
        private String strAuthSCOMApproverType = String.Empty;
        private String strSCOMmmgtServer = String.Empty;
        private String strHealthKW = String.Empty;

        private String strAzureRegion = String.Empty;
        private String strazureCognitiveServicesAPIKey = String.Empty;
        private String strMinimumPercentToCreateServiceRequest = String.Empty;
        private Guid guidEnterpriseManagementObjectID = Guid.Empty;

        public String scsmMGMTServer
        {
            get
            {
                return this.strscsmMGMTServer;
            }
            set
            {
                if (this.strscsmMGMTServer != value)
                {
                    this.strscsmMGMTServer = value;
                }
            }
        }

        public String workflowEmailAddress
        {
            get
            {
                return this.strWorkflowEmailAddress;
            }
            set
            {
                if (this.strWorkflowEmailAddress != value)
                {
                    this.strWorkflowEmailAddress = value;
                }
            }
        }

        public Boolean isAutodiscoverEnabled
        {
            get
            {
                return this.boolEnableAutodiscover;
            }
            set
            {
                if (this.boolEnableAutodiscover != value)
                {
                    this.boolEnableAutodiscover = value;
                }
            }
        }

        public Boolean createUsersNotInCMDB
        {
            get
            {
                return this.boolCreateUsersInCMDB;
            }
            set
            {
                if (this.boolCreateUsersInCMDB != value)
                {
                    this.boolCreateUsersInCMDB = value;
                }
            }
        }

        public Boolean includeWholeEmail
        {
            get
            {
                return this.boolIncludeWholeEmail;
            }
            set
            {
                if (this.boolIncludeWholeEmail != value)
                {
                    this.boolIncludeWholeEmail = value;
                }
            }
        }

        public Boolean attachEmailToWorkItem
        {
            get
            {
                return this.boolAttachEmailToWorkItem;
            }
            set
            {
                if (this.boolAttachEmailToWorkItem != value)
                {
                    this.boolAttachEmailToWorkItem = value;
                }
            }
        }

        public String autodiscoverURL
        {
            get
            {
                return this.strAutodiscoverURL;
            }
            set
            {
                if (this.strAutodiscoverURL != value)
                {
                    this.strAutodiscoverURL = value;
                }
            }
        }

        public Boolean isMaxFileSizeAttachmentsEnabled
        {
            get
            {
                return this.boolMaxFileSizeEnabled;
            }
            set
            {
                if (this.boolMaxFileSizeEnabled != value)
                {
                    this.boolMaxFileSizeEnabled = value;
                }
            }
        }

        public String minFileAttachmentSize
        {
            get
            {
                return this.strMinFileSize;
            }
            set
            {
                if (this.strMinFileSize != value)
                {
                    this.strMinFileSize = value;
                }
            }
        }

        public String ewsFilePath
        {
            get
            {
                return this.strEWSFilePath;
            }
            set
            {
                if (this.strEWSFilePath != value)
                {
                    this.strEWSFilePath = value;
                }
            }
        }

        public String mimekitFilePath
        {
            get
            {
                return this.strMimeKitFilePath;
            }
            set
            {
                if (this.strMimeKitFilePath != value)
                {
                    this.strMimeKitFilePath = value;
                }
            }
        }

        public String irTemplateName
        {
            get
            {
                return this.strIRTemplateName;
            }
            set
            {
                if (this.strIRTemplateName != value)
                {
                    this.strIRTemplateName = value;
                }
            }
        }

        public String srTemplateName
        {
            get
            {
                return this.strSRTemplateName;
            }
            set
            {
                if (this.strSRTemplateName != value)
                {
                    this.strSRTemplateName = value;
                }
            }
        }

        public String crTemplateName
        {
            get
            {
                return this.strCRTemplateName;
            }
            set
            {
                if (this.strCRTemplateName != value)
                {
                    this.strCRTemplateName = value;
                }
            }
        }

        public String prTemplateName
        {
            get
            {
                return this.strPRTemplateName;
            }
            set
            {
                if (this.strPRTemplateName != value)
                {
                    this.strPRTemplateName = value;
                }
            }
        }

        public String fromKeyword
        {
            get
            {
                return this.strFromKW;
            }
            set
            {
                if (this.strFromKW != value)
                {
                    this.strFromKW = value;
                }
            }
        }

        public String acknowledgeKeyword
        {
            get
            {
                return this.strAcknowledgeKW;
            }
            set
            {
                if (this.strAcknowledgeKW != value)
                {
                    this.strAcknowledgeKW = value;
                }
            }
        }

        public String reactivateKeyword
        {
            get
            {
                return this.strReactivateKW;
            }
            set
            {
                if (this.strReactivateKW != value)
                {
                    this.strReactivateKW = value;
                }
            }
        }

        public String resolveKeyword
        {
            get
            {
                return this.strResolvedKW;
            }
            set
            {
                if (this.strResolvedKW != value)
                {
                    this.strResolvedKW = value;
                }
            }
        }

        public String closedKeyword
        {
            get
            {
                return this.strClosedKW;
            }
            set
            {
                if (this.strClosedKW != value)
                {
                    this.strClosedKW = value;
                }
            }
        }

        public String holdKeyword
        {
            get
            {
                return this.strHoldKW;
            }
            set
            {
                if (this.strHoldKW != value)
                {
                    this.strHoldKW = value;
                }
            }
        }

        public String takeKeyword
        {
            get
            {
                return this.strTakeKW;
            }
            set
            {
                if (this.strTakeKW != value)
                {
                    this.strTakeKW = value;
                }
            }
        }

        public String cancelledKeyword
        {
            get
            {
                return this.strCancelledKW;
            }
            set
            {
                if (this.strCancelledKW != value)
                {
                    this.strCancelledKW = value;
                }
            }
        }

        public String completedKeyword
        {
            get
            {
                return this.strCompletedKW;
            }
            set
            {
                if (this.strCompletedKW != value)
                {
                    this.strCompletedKW = value;
                }
            }
        }

        public String skippedKeyword
        {
            get
            {
                return this.strSkippedKW;
            }
            set
            {
                if (this.strSkippedKW != value)
                {
                    this.strSkippedKW = value;
                }
            }
        }

        public String approvedKeyword
        {
            get
            {
                return this.strApproveKW;
            }
            set
            {
                if (this.strApproveKW != value)
                {
                    this.strApproveKW = value;
                }
            }
        }

        public String rejectedKeyword
        {
            get
            {
                return this.strRejectKW;
            }
            set
            {
                if (this.strRejectKW != value)
                {
                    this.strRejectKW = value;
                }
            }
        }

        public Boolean processCalendarAppointments
        {
            get
            {
                return this.boolProcessCalendarAppointments;
            }
            set
            {
                if (this.boolProcessCalendarAppointments != value)
                {
                    this.boolProcessCalendarAppointments = value;
                }
            }
        }

        public Boolean processMergeReplies
        {
            get
            {
                return this.boolProcessMergeReplies;
            }
            set
            {
                if (this.boolProcessMergeReplies != value)
                {
                    this.boolProcessMergeReplies = value;
                }
            }
        }

        public Boolean processClosedWorkItemsToNewWorkItems
        {
            get
            {
                return this.boolProcessClosedWorkItemsToNewWorkItems;
            }
            set
            {
                if (this.boolProcessClosedWorkItemsToNewWorkItems != value)
                {
                    this.boolProcessClosedWorkItemsToNewWorkItems = value;
                }
            }
        }

        public Boolean processEncryptedEmails
        {
            get
            {
                return this.boolProcessEncryptedEmails;
            }
            set
            {
                if (this.boolProcessEncryptedEmails != value)
                {
                    this.boolProcessEncryptedEmails = value;
                }
            }
        }

        public Boolean processDigitallySignedEmails
        {
            get
            {
                return this.boolProcessDigitallySignedEmails;
            }
            set
            {
                if (this.boolProcessDigitallySignedEmails != value)
                {
                    this.boolProcessDigitallySignedEmails = value;
                }
            }
        }

        public Boolean isAzureCognitiveServicesEnabled
        {
            get
            {
                return this.boolEnableAzureCognitiveServices;
            }
            set
            {
                if (this.boolEnableAzureCognitiveServices != value)
                {
                    this.boolEnableAzureCognitiveServices = value;
                }
            }
        }

        public Boolean isCiresonIntegrationEnabled
        {
            get
            {
                return this.boolEnableCiresonIntegration;
            }
            set
            {
                if (this.boolEnableCiresonIntegration != value)
                {
                    this.boolEnableCiresonIntegration = value;
                }
            }
        }

        public Boolean isCiresonKBSearchEnabled
        {
            get
            {
                return this.boolEnableCiresonKBSearch;
            }
            set
            {
                if (this.boolEnableCiresonKBSearch != value)
                {
                    this.boolEnableCiresonKBSearch = value;
                }
            }
        }

        public Boolean isCiresonROSearchEnabled
        {
            get
            {
                return this.boolEnableCiresonROSearch;
            }
            set
            {
                if (this.boolEnableCiresonROSearch != value)
                {
                    this.boolEnableCiresonROSearch = value;
                }
            }
        }

        public String minWordCountToSuggestRO
        {
            get
            {
                return this.strMinWordsToMatchToSuggestRO;
            }
            set
            {
                if (this.strMinWordsToMatchToSuggestRO != value)
                {
                    this.strMinWordsToMatchToSuggestRO = value;
                }
            }
        }

        public String ciresonPortalURL
        {
            get
            {
                return this.strCiresonPortalURL;
            }
            set
            {
                if (this.strCiresonPortalURL != value)
                {
                    this.strCiresonPortalURL = value;
                }
            }
        }

        public Boolean isAnnouncementIntegrationEnabled
        {
            get
            {
                return this.boolEnableAnnouncementIntegration;
            }
            set
            {
                if (this.boolEnableAnnouncementIntegration != value)
                {
                    this.boolEnableAnnouncementIntegration = value;
                }
            }
        }

        public Boolean isSCSMAnnouncementsEnabled
        {
            get
            {
                return this.boolEnableSCSMAnnouncements;
            }
            set
            {
                if (this.boolEnableSCSMAnnouncements != value)
                {
                    this.boolEnableSCSMAnnouncements = value;
                }
            }
        }

        public Boolean isCiresonAnnouncementsEnabled
        {
            get
            {
                return this.boolEnableCiresonAnnouncements;
            }
            set
            {
                if (this.boolEnableCiresonAnnouncements != value)
                {
                    this.boolEnableCiresonAnnouncements = value;
                }
            }
        }

        public String announcementKeyword
        {
            get
            {
                return this.strAnnouncementKW;
            }
            set
            {
                if (this.strAnnouncementKW != value)
                {
                    this.strAnnouncementKW = value;
                }
            }
        }

        public String authorizedAnnouncementApproverType
        {
            get
            {
                return this.strAuthorizedAnnouncer;
            }
            set
            {
                if (this.strAuthorizedAnnouncer != value)
                {
                    this.strAuthorizedAnnouncer = value;
                }
            }
        }

        public String lowPriorityAnnouncementKeyword
        {
            get
            {
                return this.strLowPriorityAnnouncementKW;
            }
            set
            {
                if (this.strLowPriorityAnnouncementKW != value)
                {
                    this.strLowPriorityAnnouncementKW = value;
                }
            }
        }

        public String mediumPriorityAnnouncementKeyword
        {
            get
            {
                return this.strMedPriorityAnnouncementKW;
            }
            set
            {
                if (this.strMedPriorityAnnouncementKW != value)
                {
                    this.strMedPriorityAnnouncementKW = value;
                }
            }
        }

        public String highPriorityAnnouncementKeyword
        {
            get
            {
                return this.strHighPriorityAnnouncementKW;
            }
            set
            {
                if (this.strHighPriorityAnnouncementKW != value)
                {
                    this.strHighPriorityAnnouncementKW = value;
                }
            }
        }

        public String lowPriorityExpiresInHours
        {
            get
            {
                return this.strLowPriorityExpInHours;
            }
            set
            {
                if (this.strLowPriorityExpInHours != value)
                {
                    this.strLowPriorityExpInHours = value;
                }
            }
        }

        public String mediumPriorityExpiresInHours
        {
            get
            {
                return this.strMedPriorityExpInHours;
            }
            set
            {
                if (this.strMedPriorityExpInHours != value)
                {
                    this.strMedPriorityExpInHours = value;
                }
            }
        }

        public String highPriorityExpiresInHours
        {
            get
            {
                return this.strHighPriorityExpInHours;
            }
            set
            {
                if (this.strHighPriorityExpInHours != value)
                {
                    this.strHighPriorityExpInHours = value;
                }
            }
        }

        public String authorizedSCOMApproverType
        {
            get
            {
                return this.strAuthSCOMApproverType;
            }
            set
            {
                if (this.strAuthSCOMApproverType != value)
                {
                    this.strAuthSCOMApproverType = value;
                }
            }
        }

        public Boolean isSCOMIntegrationEnabled
        {
            get
            {
                return this.boolEnableSCOMIntegration;
            }
            set
            {
                if (this.boolEnableSCOMIntegration != value)
                {
                    this.boolEnableSCOMIntegration = value;
                }
            }
        }

        public String SCOMServer
        {
            get
            {
                return this.strSCOMmmgtServer;
            }
            set
            {
                if (this.strSCOMmmgtServer != value)
                {
                    this.strSCOMmmgtServer = value;
                }
            }
        }

        public String healthKeyword
        {
            get
            {
                return this.strHealthKW;
            }
            set
            {
                if (this.strHealthKW != value)
                {
                    this.strHealthKW = value;
                }
            }
        }

        public String azureRegion
        {
            get
            {
                return this.strAzureRegion;
            }
            set
            {
                if (this.strAzureRegion != value)
                {
                    this.strAzureRegion = value;
                }
            }
        }

        public String defaultWorkItem
        {
            get
            {
                return this.strDefaultWorkItem;
            }
            set
            {
                if (this.strDefaultWorkItem != value)
                {
                    this.strDefaultWorkItem = value;
                }
            }
        }

        public String azureCognitiveServicesAPIKey
        {
            get
            {
                return this.strazureCognitiveServicesAPIKey;
            }
            set
            {
                if (this.strazureCognitiveServicesAPIKey != value)
                {
                    this.strazureCognitiveServicesAPIKey = value;
                }
            }
        }

        public String minimumPercentToCreateServiceRequest
        {
            get
            {
                return this.strMinimumPercentToCreateServiceRequest;
            }
            set
            {
                if (this.strMinimumPercentToCreateServiceRequest != value)
                {
                    this.strMinimumPercentToCreateServiceRequest = value;
                }
            }
        }

        public Guid EnterpriseManagementObjectID
        {
            get
            {
                return this.guidEnterpriseManagementObjectID;
            }
            set
            {
                if (this.guidEnterpriseManagementObjectID != value)
                {
                    this.guidEnterpriseManagementObjectID = value;
                }
            }
        }

        #endregion

        //load the currently defined settings in the management pack
        internal AdminSettingWizardData(EnterpriseManagementObject emoAdminSetting)
        {
            //Get the server name to connect to and then connect 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

            //Get the SMLets Settings class
            ManagementPackClass smletsExchangeConnectorSettingsClass = emg.EntityTypes.GetClass(new Guid("6FFB0179-D405-DCCB-FC28-59845DC4A4BA"));

            //General Settings
            this.scsmMGMTServer = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmMGMTServer"].ToString();
            this.workflowEmailAddress = emoAdminSetting[smletsExchangeConnectorSettingsClass, "workflowEmailAddress"].ToString();
            //this.autodiscoverURL = emoAdminSetting[smletsExchangeConnectorSettingsClass, "exchangeAutodiscoverURL"].ToString();

            //autodiscover
            try { this.isAutodiscoverEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "useAutoDiscover"].ToString()); }
            catch { this.isAutodiscoverEnabled = false; }

            //create users not found in CMDB
            try { this.createUsersNotInCMDB = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "createUsersNotInCMDB"].ToString()); }
            catch { this.createUsersNotInCMDB = false; }

            //include whole email on Action Log
            try { this.includeWholeEmail = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "includeWholeEmail"].ToString()); }
            catch { this.includeWholeEmail = false; }

            //attach email to work item on the related items tab
            try { this.attachEmailToWorkItem = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "attachEmailToWorkItem"].ToString()); }
            catch { this.attachEmailToWorkItem = false; }

            //DLL Paths for EWS and Mimekit
            this.ewsFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "locationExchangeWebServicesDLL"].ToString();
            this.mimekitFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "locationMimeKitDLL"].ToString();

            //Template selection
            this.defaultWorkItem = emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultWorkItem"].ToString();
            this.irTemplateName = emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultIncidentTemplate"].ToString();
            this.srTemplateName = emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultServiceRequestTemplate"].ToString();
            this.crTemplateName = emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultProblemTemplate"].ToString();
            this.prTemplateName = emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultChangeRequestTemplate"].ToString();

            //File Attachments
            if (emoAdminSetting[smletsExchangeConnectorSettingsClass, "minimumFileAttachmentSize"].ToString() == "(null)") { this.minFileAttachmentSize = "25"; }
            else { this.minFileAttachmentSize = emoAdminSetting[smletsExchangeConnectorSettingsClass, "minimumFileAttachmentSize"].ToString(); }
            try { this.isMaxFileSizeAttachmentsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enforceFileAttachmentSettings"].ToString()); }
            catch { this.isMaxFileSizeAttachmentsEnabled = false; }

            //Keywords
            this.fromKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordFrom"].ToString();
            this.acknowledgeKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordAcknowledge"].ToString();
            this.reactivateKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordReactivate"].ToString();
            this.resolveKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordResolved"].ToString();
            this.closedKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordClosed"].ToString();
            this.holdKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordHold"].ToString();
            this.takeKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordTake"].ToString();
            this.cancelledKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordCancelled"].ToString();
            this.completedKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordCompleted"].ToString();
            this.skippedKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordSkipped"].ToString();
            this.approvedKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordApprove"].ToString();
            this.rejectedKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordReject"].ToString();

            //Cognitive Services
            this.azureRegion = emoAdminSetting[smletsExchangeConnectorSettingsClass, "azureTextAnalyticsRegion"].ToString();
            this.azureCognitiveServicesAPIKey = emoAdminSetting[smletsExchangeConnectorSettingsClass, "azureCognitiveServicesTextAnalyticsAPIKey"].ToString();
            this.minimumPercentToCreateServiceRequest = emoAdminSetting[smletsExchangeConnectorSettingsClass, "minimumPercentToCreateServiceRequest"].ToString();
            try { this.isAzureCognitiveServicesEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableAzureCognitiveServicesTextAnalytics"].ToString()); }
            catch { this.isAzureCognitiveServicesEnabled = false; }

            //Cireson Integration
            this.ciresonPortalURL = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ciresonPortalURL"].ToString();
            this.minWordCountToSuggestRO = emoAdminSetting[smletsExchangeConnectorSettingsClass, "numberOfWordsToMatchFromEmailToRequestOffering"].ToString();
            try { this.isCiresonIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableCiresonIntegration"].ToString()); }
            catch { this.isCiresonIntegrationEnabled = false; }

            //Ciresn Knowledge Base suggestions
            try { this.isCiresonKBSearchEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "ciresonSearchKnowledgeBase"].ToString()); }
            catch { this.isCiresonKBSearchEnabled = false; }

            //Cireson Service Catalog suggestions
            try { this.isCiresonROSearchEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "ciresonSearchRequestOfferings"].ToString()); }
            catch { this.isCiresonROSearchEnabled = false; }

            //Announcements
            this.announcementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeyword"].ToString();
            this.authorizedAnnouncementApproverType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmAnnouncementApprovedMemberType"].ToString();
            this.lowPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeywordLow"].ToString();
            this.mediumPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeywordMedium"].ToString();
            this.highPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeywordHigh"].ToString();
            this.lowPriorityExpiresInHours = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementPriorityLowExpirationInHours"].ToString();
            this.mediumPriorityExpiresInHours = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementPriorityNormalExpirationInHours"].ToString();
            this.highPriorityExpiresInHours = emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementPriorityCriticalExpirationInHours"].ToString();
            try { this.isAnnouncementIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableAnnouncements"].ToString()); }
            catch { this.isAnnouncementIntegrationEnabled = false; }

            try { this.isSCSMAnnouncementsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableSCSMAnnouncements"].ToString()); }
            catch { this.isSCSMAnnouncementsEnabled = false; }

            try { this.isCiresonAnnouncementsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableCiresonAnnouncements"].ToString()); }
            catch { this.isCiresonAnnouncementsEnabled = false; }

            //SCOM Integration
            try { this.isSCOMIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableSCOMIntegration"].ToString()); }
            catch { this.isSCOMIntegrationEnabled = false; }
            this.SCOMServer = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scomMGMTServer"].ToString();
            this.healthKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scomDAHealthKeyword"].ToString();
            this.authorizedSCOMApproverType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "scomApprovedMemberType"].ToString();

            //Processing Logic
            try { this.processCalendarAppointments = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "processCalendarAppointments"].ToString()); }
            catch { this.processCalendarAppointments = false; }
            try { this.processMergeReplies = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "mergeReplies"].ToString()); }
            catch { this.processMergeReplies = false; }
            try { this.processClosedWorkItemsToNewWorkItems = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "createNewWorkItemIfWorkItemClosed"].ToString()); }
            catch { this.processClosedWorkItemsToNewWorkItems = false; }
            try { this.processEncryptedEmails = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "processDigitallyEncryptedMessages"].ToString()); }
            catch { this.processEncryptedEmails = false; }
            try { this.processDigitallySignedEmails = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "processDigitallySignedMessages"].ToString()); }
            catch { this.processDigitallySignedEmails = false; }

            this.EnterpriseManagementObjectID = emoAdminSetting.Id;
        }

        //save the values back to the management pack
        public override void AcceptChanges(WizardMode wizardMode)
        {
            //Get the server name to connect to and connect 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);
            ManagementPackClass smletsExchangeConnectorSettingsClass = emg.EntityTypes.GetClass(new Guid("6FFB0179-D405-DCCB-FC28-59845DC4A4BA"));

            //Get the Connector object using the object ID 
            EnterpriseManagementObject emoAdminSetting = emg.EntityObjects.GetObject<EnterpriseManagementObject>(this.EnterpriseManagementObjectID, ObjectQueryOptions.Default);

            //General Settings
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmMGMTServer"].Value = this.scsmMGMTServer;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "workflowEmailAddress"].Value = this.workflowEmailAddress;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "useAutoDiscover"].Value = this.isAutodiscoverEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "exchangeAutodiscoverURL"].Value = this.autodiscoverURL;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "createUsersNotInCMDB"].Value = this.createUsersNotInCMDB;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "includeWholeEmail"].Value = this.includeWholeEmail;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "attachEmailToWorkItem"].Value = this.attachEmailToWorkItem;

            //DLL Paths for EWS and Mimekit
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "locationExchangeWebServicesDLL"].Value = this.ewsFilePath;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "locationMimeKitDLL"].Value = this.mimekitFilePath;

            //Templates
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultWorkItem"].Value = this.defaultWorkItem;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultIncidentTemplate"].Value = this.irTemplateName;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultServiceRequestTemplate"].Value = this.srTemplateName;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultProblemTemplate"].Value = this.prTemplateName;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "defaultChangeRequestTemplate"].Value = this.crTemplateName;

            //File Attachments
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "minimumFileAttachmentSize"].Value = this.minFileAttachmentSize;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enforceFileAttachmentSettings"].Value = this.isMaxFileSizeAttachmentsEnabled;

            //Keywords
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordFrom"].Value = this.fromKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordAcknowledge"].Value = this.acknowledgeKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordReactivate"].Value = this.reactivateKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordResolved"].Value = this.resolveKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordClosed"].Value = this.closedKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordHold"].Value = this.holdKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordTake"].Value = this.takeKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordCancelled"].Value = this.cancelledKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordCompleted"].Value = this.completedKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordSkipped"].Value = this.skippedKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordApprove"].Value = this.approvedKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmKeywordReject"].Value = this.rejectedKeyword;

            //Set Cognitive Services
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableAzureCognitiveServicesTextAnalytics"].Value = this.isAzureCognitiveServicesEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "azureTextAnalyticsRegion"].Value = this.azureRegion;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "azureCognitiveServicesTextAnalyticsAPIKey"].Value = this.azureCognitiveServicesAPIKey;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "minimumPercentToCreateServiceRequest"].Value = this.minimumPercentToCreateServiceRequest;

            //Announcements
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableAnnouncements"].Value = this.isAnnouncementIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableSCSMAnnouncements"].Value = this.isSCSMAnnouncementsEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scsmAnnouncementApprovedMemberType"].Value = this.authorizedAnnouncementApproverType;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableCiresonAnnouncements"].Value = this.isCiresonAnnouncementsEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeyword"].Value = this.announcementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeywordLow"].Value = this.lowPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeywordMedium"].Value = this.mediumPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementKeywordHigh"].Value = this.highPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementPriorityLowExpirationInHours"].Value = this.lowPriorityExpiresInHours;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementPriorityNormalExpirationInHours"].Value = this.mediumPriorityExpiresInHours;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "announcementPriorityCriticalExpirationInHours"].Value = this.highPriorityExpiresInHours;

            //Processing Logic
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "processCalendarAppointments"].Value = this.processCalendarAppointments;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "mergeReplies"].Value = this.processMergeReplies;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "createNewWorkItemIfWorkItemClosed"].Value = this.processClosedWorkItemsToNewWorkItems;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "processDigitallyEncryptedMessages"].Value = this.processEncryptedEmails;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "processDigitallySignedMessages"].Value = this.processDigitallySignedEmails;

            //SCOM Integration
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableSCOMIntegration"].Value = this.isSCOMIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scomMGMTServer"].Value = this.SCOMServer;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scomDAHealthKeyword"].Value = this.healthKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "scomApprovedMemberType"].Value = this.authorizedSCOMApproverType;

            //Cireson Integration
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "enableCiresonIntegration"].Value = this.isCiresonIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ciresonSearchKnowledgeBase"].Value = this.isCiresonKBSearchEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ciresonSearchRequestOfferings"].Value = this.isCiresonROSearchEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ciresonPortalURL"].Value = this.ciresonPortalURL;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "numberOfWordsToMatchFromEmailToRequestOffering"].Value = this.minWordCountToSuggestRO;

            //Update the MP
            emoAdminSetting.Commit();
            this.WizardResult = WizardResult.Success;
        }
    }
}