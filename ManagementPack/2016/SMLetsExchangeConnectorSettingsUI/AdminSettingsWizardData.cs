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
using Microsoft.EnterpriseManagement.ConnectorFramework;


namespace SMLetsExchangeConnectorSettingsUI
{
    class AdminSettingWizardData : WizardData
    {
        #region Variable Declaration
        //looking to collapse some/all of these quickly? CTRL + MM
        //Yes, that's really two M back to back

        //general settings page
        private String strscsmMGMTServer = String.Empty;
        private String strWorkflowEmailAddress = String.Empty;
        private Boolean boolEnableAutodiscover = false;
        private String strAutodiscoverURL = String.Empty;
        private Boolean boolEnableExchangeOnline = false;
        private String strAzureTenantID = String.Empty;
        private String strAzureAppID = String.Empty;
        //run as account - exchange web services
        ManagementPackSecureReference runasaccountews;

        //paths
        private String strEWSFilePath = String.Empty;
        private String strSMExcoFilePath = String.Empty;
        private String strMimeKitFilePath = String.Empty;
        private String strPIIRegexFilePath = String.Empty;
        private String strCiresonHTMLSuggestionTemplatesPath = String.Empty;
        private String strCustomEvents = String.Empty;

        //processing logic
        private Boolean boolCreateUsersInCMDB = false;
        private Boolean boolIncludeWholeEmail = false;
        private Boolean boolAttachEmailToWorkItem = false;
        private Boolean boolProcessADGroupVote = false;
        private Boolean boolProcessMergeReplies = false;
        private Boolean boolMailboxRedirection = false;
        private Boolean boolEnforceSupportGroupTake = false;
        private Boolean boolProcessClosedWorkItemsToNewWorkItems = false;
        private Boolean boolChangeIncidentStatusOnReply = false;
        private Boolean boolRemovePII = false;
        private String strDynamicAnalyst = String.Empty;

        //processing logic - additional mail types
        private Boolean boolProcessCalendarAppointments = false;
        private Boolean boolProcessEncryptedEmails = false;
        private Boolean boolProcessDigitallySignedEmails = false;
        private String strCertStore = String.Empty;

        //processing logic - additional mail types
        private String strExternalIRCommentPrivate = String.Empty;
        private String strExternalSRCommentPrivate = String.Empty;
        private String strExternalIRCommentType = String.Empty;
        private String strExternalSRCommentType = String.Empty;

        //processing logic, lists: Incident Status based on Who Replied
        ManagementPackEnumeration defaultIncidentStatusOnAUReplyEnum;
        ManagementPackEnumeration defaultIncidentStatusOnATReplyEnum;
        ManagementPackEnumeration defaultIncidentStatusOnRelReplyEnum;
        
        //processing logic, lists: Incident Resolution, Service Request, and Problem Category
        ManagementPackEnumeration defaultIncidentResolutionCategoryEnum;
        ManagementPackEnumeration defaultServiceRequestImplementationCategoryEnum;
        ManagementPackEnumeration defaultProblemResolutionCategoryEnum;

        //processing logic, lists:  Support Groups for Change Request, Manual Activity, and Problem Extended Classes
        ManagementPackProperty manualActivitySupportGroupExtensionEnum; 
        ManagementPackProperty changeRequestSupportGroupExtensionEnum;
        ManagementPackProperty problemSupportGroupExtensionEnum;

        //file attachments
        private String strMinFileSize = "0";
        private Boolean boolMaxFileSizeEnabled = false;

        //templates - default work item
        private String strDefaultWorkItem = String.Empty;

        //templates - Incident, Service Request, Change, and Problem
        ManagementPackObjectTemplate defaultIncidentTemplate;
        ManagementPackObjectTemplate defaultServiceRequestTemplate;
        ManagementPackObjectTemplate defaultChangeRequestTemplate;
        ManagementPackObjectTemplate defaultProblemTemplate;
        
        //mailbox redirection
        private String[] multipleMailboxesToAdd;
        private String[] multipleMailboxesToRemove;

        //keywords
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
        private String strHealthKW = String.Empty;
        private String strAnnouncementKW = String.Empty;
        private String strPrivateKW = String.Empty;

        //cireson
        private Boolean boolCiresonMPExists = false;
        private String strCiresonPortalURL = String.Empty;
        private Int32 strMinWordsToMatchToSuggestRO = 0;
        private Int32 strMinWordCountToSuggestKA = 0;
        private String strAddWatchlistKW = String.Empty;
        private String strRemoveWatchlistKW = String.Empty;
        private Boolean boolEnableAnnouncementIntegration = false;
        private Boolean boolEnableSCSMAnnouncements = false;
        private Boolean boolEnableCiresonAnnouncements = false;
        private Boolean boolEnableCiresonIntegration = false;
        private Boolean boolEnableCiresonKBSearch = false;
        private Boolean boolEnableCiresonROSearch = false;
        private Boolean boolEnableCiresonFirstResponseDateOnSuggestions = false;
        //run as account - cireson portal
        ManagementPackSecureReference runasaccountciresonportal;

        //announcements
        private String strAuthorizedAnnouncer = String.Empty;
        private String scsmApprovedAnnouncers = String.Empty;
        private String strLowPriorityAnnouncementKW = String.Empty;
        private String strMedPriorityAnnouncementKW = String.Empty;
        private String strHighPriorityAnnouncementKW = String.Empty;
        private String strLowPriorityExpInHours = "0";
        private String strMedPriorityExpInHours = "0";
        private String strHighPriorityExpInHours = "0";
        private IDataItem scsmApprovedAnnouncementGroup;
        
        //scom
        private Boolean boolEnableSCOMIntegration = false;
        private String strAuthSCOMApproverType = String.Empty;
        private String strSCOMmmgtServer = String.Empty;
        private String strAuthorizedSCOMUsers = String.Empty;   
        private IDataItem scomApprovedGroup;

        //artificial intelligence, option 0: Disabled
        private Boolean boolArtificialIntelligence = false;

        //artificial intelligence, option 1: ACS
        private Boolean boolEnableAzureCognitiveServices = false;
        private Boolean boolEnableACSForKA = false;
        private Boolean boolEnableACSForRO = false;
        private Boolean boolEnableACSForPriorityScoring = false;
        private String strAzureRegion = String.Empty;
        private String strazureCognitiveServicesAPIKey = String.Empty;
        private String strBoundaries = String.Empty;
        ManagementPackProperty incidentACSSentimentExtensionDec;
        ManagementPackProperty serviceRequestACSSentimentExtensionDec;
        private String strMinimumPercentToCreateServiceRequest = "0";
        ManagementPackEnumeration incidentImpactEnum;
        ManagementPackEnumeration incidentUrgencyEnum;
        ManagementPackEnumeration serviceRequestUrgencyEnum;
        ManagementPackEnumeration serviceRequestPriorityEnum;

        //artificial intelligencee, option 2: Keyword matching
        private Boolean boolEnableKeywordMatching = false;
        private String strKWMatchingRegexPattern = String.Empty;
        private String strKWWorkItemOverride = String.Empty;

        //artificial intelligencee, option 3: Azure Machine Learning
        private Boolean boolEnableAzureMachineLearning = false;
        private String strAMLAPIKey = String.Empty;
        private String strAMLWebServiceURL = String.Empty;
        private String strAMLWIConfidenceClassName = String.Empty;
        private String strAMLClassificationConfidenceClassName = String.Empty;
        private String strAMLSupportGroupConfidenceClassName = String.Empty;
        private String decAMLWIMinConfidence = "0";
        private String decAMLClassificationMinConfidence = "0";
        private String decAMLSupportGroupMinConfidence = "0";
        private String decAMLConfigItemMinConfidence = "0";

            //AML incident custom decimal extensions
            ManagementPackProperty incidentAMLWIConfidenceExtensionDec;
            ManagementPackProperty incidentAMLWITypePredictionExtensionGUID;
            ManagementPackProperty incidentAMLClassifcationConfidenceExtensionDec;
            ManagementPackProperty incidentAMLSupportGroupIConfidenceExtensionDec;
            ManagementPackProperty incidentAMLClassificationPredictionExtensionEnum;
            ManagementPackProperty incidentAMLSupportGroupPredictionExtensionEnum;

            //AML service request custom decimal extensions
            ManagementPackProperty serviceRequestAMLWIConfidenceExtensionDec;
            ManagementPackProperty serviceRequestAMLWITypePredictionExtensionGUID;
            ManagementPackProperty serviceRequestAMLClassifcationConfidenceExtensionDec;
            ManagementPackProperty serviceRequestAMLSupportGroupIConfidenceExtensionDec;
            ManagementPackProperty serviceRequestAMLClassificationPredictionExtensionEnum;
            ManagementPackProperty serviceRequestAMLSupportGroupPredictionExtensionEnum;

        //azure translate
        private Boolean boolEnableAzureTranslate = false;
        private String strAzureTranslateAPIKey = String.Empty;
        private String strAzureTranslateDefaultLanguageCode = String.Empty;

        //azure vision
        private Boolean boolEnableAzureVision = false;
        private String strAzureVisionAPIKey = String.Empty;
        private String strAzureVisionRegion = String.Empty;

        //azure speech
        private Boolean boolEnableAzureSpeech = false;
        private String strAzureSpeechAPIKey = String.Empty;
        private String strAzureSpeechRegion = String.Empty;
        
        //smlets exchange connector workflow mp
        private Boolean boolSMexcoWFEnabled = false;
        private String strSMExcoWFInterval = "0";

        //management pack guid
        private Guid guidEnterpriseManagementObjectID = Guid.Empty;
        #endregion

        //notify on property changed
        //https://docs.microsoft.com/en-us/dotnet/framework/wpf/data/how-to-implement-property-change-notification
        public event PropertyChangedEventHandler PropertyChanged;
        protected void NotifyPropertyChanged(string propchanger)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propchanger));
            }
        }

        #region Variable Get/Set.
        //WPF Forms bind to these and they are referenced in load/save.
        //Apart from the slightly different names in their initial declaration above
        //these can be told apart as they all begin with Capital letters

        //general settings page
        public String SCSMmanagementServer
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

        public String WorkflowEmailAddress
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

        public Boolean IsAutodiscoverEnabled
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

        public String ExchangeAutodiscoverURL
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
        
        public Boolean IsExchangeOnline
        {
            get
            {
                return this.boolEnableExchangeOnline;
            }
            set
            {
                if (this.boolEnableExchangeOnline != value)
                {
                    this.boolEnableExchangeOnline = value;
                }
            }
        }
        
        public String AzureTenantID
        {
            get
            {
                return this.strAzureTenantID;
            }
            set
            {
                if (this.strAzureTenantID != value)
                {
                    this.strAzureTenantID = value;
                }
            }
        }

        public String AzureClientID
        {
            get
            {
                return this.strAzureAppID;
            }
            set
            {
                if (this.strAzureAppID != value)
                {
                    this.strAzureAppID = value;
                }
            }
        }
        
        //run as account secure references
        public ManagementPackSecureReference RunAsAccountEWS
        {
            get
            {
                return runasaccountews;
            }
            set
            {
                if (this.runasaccountews != value)
                {
                    runasaccountews = value;
                    NotifyPropertyChanged("runAsAccount");
                }
            }
        }
        public IList<ManagementPackSecureReference> SecureRunAsAccounts { get; set; }

        //file attachments
        public Boolean IsMaxFileSizeAttachmentsEnabled
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

        public String MinFileAttachmentSize
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

        //file paths
        public String EWSFilePath
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
        
        public String SMExcoFilePath
        {
            get
            {
                return this.strSMExcoFilePath;
            }
            set
            {
                if (this.strSMExcoFilePath != value)
                {
                    this.strSMExcoFilePath = value;
                }
            }
        }

        public String MimeKitFilePath
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

        public String PIIRegexFilePath
        {
            get
            {
                return this.strPIIRegexFilePath;
            }
            set
            {
                if (this.strPIIRegexFilePath != value)
                {
                    this.strPIIRegexFilePath = value;
                }
            }
        }

        public String HTMLSuggestionTemplatesFilePath
        {
            get
            {
                return this.strCiresonHTMLSuggestionTemplatesPath;
            }
            set
            {
                if (this.strCiresonHTMLSuggestionTemplatesPath != value)
                {
                    this.strCiresonHTMLSuggestionTemplatesPath = value;
                }
            }
        }

        public String CustomEventsFilePath
        {
            get
            {
                return this.strCustomEvents;
            }
            set
            {
                if (this.strCustomEvents != value)
                {
                    this.strCustomEvents = value;
                }
            }
        }

        //templates
        public ManagementPackObjectTemplate DefaultIncidentTemplate
        {
            get
            {
                return defaultIncidentTemplate;
            }
            set
            {
                if (this.defaultIncidentTemplate != value)
                {
                    defaultIncidentTemplate = value;
                    NotifyPropertyChanged("defaultIncidentTemplateGUID");
                }
            }
        }
        public IList<ManagementPackObjectTemplate> IncidentTemplates { get; set; }
        public ManagementPackObjectTemplate DefaultServiceRequestTemplate
        {
            get
            {
                return defaultServiceRequestTemplate;
            }
            set
            {
                if (this.defaultServiceRequestTemplate != value)
                {
                    defaultServiceRequestTemplate = value;
                    NotifyPropertyChanged("defaultServiceRequestTemplateGUID");
                }
            }
        }
        public IList<ManagementPackObjectTemplate> ServiceRequestTemplates { get; set; }
        public ManagementPackObjectTemplate DefaultChangeRequestTemplate
        {
            get
            {
                return defaultChangeRequestTemplate;
            }
            set
            {
                if (this.defaultChangeRequestTemplate != value)
                {
                    defaultChangeRequestTemplate = value;
                    NotifyPropertyChanged("defaultChangeRequestTemplateGUID");
                }
            }
        }
        public IList<ManagementPackObjectTemplate> ChangeRequestTemplates { get; set; }
        public ManagementPackObjectTemplate DefaultProblemTemplate
        {
            get
            {
                return defaultProblemTemplate;
            }
            set
            {
                if (this.defaultProblemTemplate != value)
                {
                    defaultProblemTemplate = value;
                    NotifyPropertyChanged("defaultProblemTemplateGUID");
                }
            }
        }
        public IList<ManagementPackObjectTemplate> ProblemTemplates { get; set; }
        
        //processing logic
        public Boolean CreateUsersNotFoundtInCMDB
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

        public Boolean IncludeWholeEmail
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

        public Boolean AttachEmailToWorkItem
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

        public Boolean ProcessCalendarAppointments
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

        public Boolean ProcessMergeReplies
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

        public Boolean ProcessADGroupVote
        {
            get
            {
                return this.boolProcessADGroupVote;
            }
            set
            {
                if (this.boolProcessADGroupVote != value)
                {
                    this.boolProcessADGroupVote = value;
                }
            }
        }

        public Boolean EnforceSupportGroupTake
        {
            get
            {
                return this.boolEnforceSupportGroupTake;
            }
            set
            {
                if (this.boolEnforceSupportGroupTake != value)
                {
                    this.boolEnforceSupportGroupTake = value;
                }
            }
        }

        public Boolean ProcessClosedWorkItemsToNewWorkItems
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

        public Boolean ChangeIncidentStatusOnReply
        {
            get
            {
                return this.boolChangeIncidentStatusOnReply;
            }
            set
            {
                if (this.boolChangeIncidentStatusOnReply != value)
                {
                    this.boolChangeIncidentStatusOnReply = value;
                }
            }
        }

        public Boolean RemovePII
        {
            get
            {
                return this.boolRemovePII;
            }
            set
            {
                if (this.boolRemovePII != value)
                {
                    this.boolRemovePII = value;
                }
            }
        }

        public String DynamicAnalystAssignment
        {
            get
            {
                return this.strDynamicAnalyst;
            }
            set
            {
                if (this.strDynamicAnalyst != value)
                {
                    this.strDynamicAnalyst = value;
                }
            }
        }

        public Boolean ProcessEncryptedEmails
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

        public Boolean ProcessDigitallySignedEmails
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

        public String CertificateStore
        {
            get
            {
                return this.strCertStore;
            }
            set
            {
                if (this.strCertStore != value)
                {
                    this.strCertStore = value;
                }
            }
        }

        public String ExternalIRCommentPrivate
        {
            get
            {
                return this.strExternalIRCommentPrivate;
            }
            set
            {
                if (this.strExternalIRCommentPrivate != value)
                {
                    this.strExternalIRCommentPrivate = value;
                }
            }
        }

        public String ExternalIRCommentType
        {
            get
            {
                return this.strExternalIRCommentType;
            }
            set
            {
                if (this.strExternalIRCommentType != value)
                {
                    this.strExternalIRCommentType = value;
                }
            }
        }

        public String ExternalSRCommentPrivate
        {
            get
            {
                return this.strExternalSRCommentPrivate;
            }
            set
            {
                if (this.strExternalSRCommentPrivate != value)
                {
                    this.strExternalSRCommentPrivate = value;
                }
            }
        }

        public String ExternalSRCommentType
        {
            get
            {
                return this.strExternalSRCommentType;
            }
            set
            {
                if (this.strExternalSRCommentType != value)
                {
                    this.strExternalSRCommentType = value;
                }
            }
        }

        //processing logic, support group extensions
        public IList<ManagementPackProperty> ManualActivityEnumExtensions { get; set; }
        public IList<ManagementPackProperty> ChangeRequestEnumExtensions { get; set; }
        public IList<ManagementPackProperty> ProblemEnumExtensions { get; set; }
        public ManagementPackProperty ManualActivitySupportGroupExtensionEnum
        {
            get
            {
                return manualActivitySupportGroupExtensionEnum;
            }
            set
            {
                if (this.manualActivitySupportGroupExtensionEnum != value)
                {
                    manualActivitySupportGroupExtensionEnum = value;
                }
            }
        }
        public ManagementPackProperty ProblemSupportGroupExtensionEnum
        {
            get
            {
                return problemSupportGroupExtensionEnum;
            }
            set
            {
                if (this.problemSupportGroupExtensionEnum != value)
                {
                    problemSupportGroupExtensionEnum = value;
                }
            }
        }
        public ManagementPackProperty ChangeRequestSupportGroupExtensionEnum
        {
            get
            {
                return changeRequestSupportGroupExtensionEnum;
            }
            set
            {
                if (this.changeRequestSupportGroupExtensionEnum != value)
                {
                    changeRequestSupportGroupExtensionEnum = value;
                }
            }
        }

        //processing logic, change incident status who replied
        public IList<ManagementPackEnumeration> IncidentStatusEnums { get; set; }
        public ManagementPackEnumeration DefaultIncidentStatusOnAUReplyEnum
        {
            get
            {
                return defaultIncidentStatusOnAUReplyEnum;
            }
            set
            {
                if (this.defaultIncidentStatusOnAUReplyEnum != value)
                {
                    defaultIncidentStatusOnAUReplyEnum = value;
                    NotifyPropertyChanged("incidentStatusOnAffectedUserReply");
                }
            }
        }
        public ManagementPackEnumeration DefaultIncidentStatusOnATReplyEnum
        {
            get
            {
                return defaultIncidentStatusOnATReplyEnum;
            }
            set
            {
                if (this.defaultIncidentStatusOnATReplyEnum != value)
                {
                    defaultIncidentStatusOnATReplyEnum = value;
                    NotifyPropertyChanged("incidentStatusOnAssignedToReply");
                }
            }
        }
        public ManagementPackEnumeration DefaultIncidentStatusOnRelReplyEnum
        {
            get
            {
                return defaultIncidentStatusOnRelReplyEnum;
            }
            set
            {
                if (this.defaultIncidentStatusOnRelReplyEnum != value)
                {
                    defaultIncidentStatusOnRelReplyEnum = value;
                    NotifyPropertyChanged("incidentStatusOnRelatedUserReply");
                }
            }
        }

        //processing logic, default resolution categories
        public ManagementPackEnumeration DefaultIncidentResolutionCategoryEnum
        {
            get
            {
                return defaultIncidentResolutionCategoryEnum;
            }
            set
            {
                if (this.defaultIncidentResolutionCategoryEnum != value)
                {
                    defaultIncidentResolutionCategoryEnum = value;
                    NotifyPropertyChanged("incidentResolutionCategory");
                }
            }
        }
        public IList<ManagementPackEnumeration> IncidentResolutionCategoryEnums { get; set; }
        public ManagementPackEnumeration DefaultServiceRequestImplementationCategoryEnum
        {
            get
            {
                return defaultServiceRequestImplementationCategoryEnum;
            }
            set
            {
                if (this.defaultServiceRequestImplementationCategoryEnum != value)
                {
                    defaultServiceRequestImplementationCategoryEnum = value;
                    NotifyPropertyChanged("serviceRequestImplementationCategory");
                }
            }
        }
        public IList<ManagementPackEnumeration> ServiceRequestImplementationCategoryEnums { get; set; }
        public ManagementPackEnumeration DefaultProblemResolutionCategoryEnum
        {
            get
            {
                return defaultProblemResolutionCategoryEnum;
            }
            set
            {
                if (this.defaultProblemResolutionCategoryEnum != value)
                {
                    defaultProblemResolutionCategoryEnum = value;
                    NotifyPropertyChanged("problemResolutionCategory");
                }
            }
        }
        public IList<ManagementPackEnumeration> ProblemResolutionCategoryEnums { get; set; }
        
        //multi-mailbox
        public Boolean UseMailboxRedirection
        {
            get
            {
                return this.boolMailboxRedirection;
            }
            set
            {
                if (this.boolMailboxRedirection != value)
                {
                    this.boolMailboxRedirection = value;
                }
            }
        }
        public String[] MultipleMailboxesToAdd
        {
            get
            {
                return this.multipleMailboxesToAdd;
            }
            set
            {
                if (this.multipleMailboxesToAdd != value)
                {
                    this.multipleMailboxesToAdd = value;
                }
            }
        }
        public IList<EnterpriseManagementObject> AdditionalMailboxes { get; set; }
        public String[] MultipleMailboxesToRemove
        {
            get
            {
                return this.multipleMailboxesToRemove;
            }
            set
            {
                if (this.multipleMailboxesToRemove != value)
                {
                    this.multipleMailboxesToRemove = value;
                }
            }
        }

        //keywords
        public String KeywordFrom
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

        public String KeywordAcknowledge
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

        public String KeywordReactivate
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

        public String KeywordResolve
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

        public String KeywordClose
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

        public String KeywordHold
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

        public String KeywordTake
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

        public String KeywordCancel
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

        public String KeywordComplete
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

        public String KeywordSkip
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

        public String KeywordApprove
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

        public String KeywordReject
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

        public String KeywordAnnouncement
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

        public String KeywordPrivate
        {
            get
            {
                return this.strPrivateKW;
            }
            set
            {
                if (this.strPrivateKW != value)
                {
                    this.strPrivateKW = value;
                }
            }
        }

        public String KeywordHealth
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

        //cireson
        public Boolean IsCiresonPortalMPImported
        {
            get
            {
                return this.boolCiresonMPExists;
            }
            set
            {
                if (this.boolCiresonMPExists != value)
                {
                    this.boolCiresonMPExists = value;
                }
            }
        }

        public String KeywordAddWatchlist
        {
            get
            {
                return this.strAddWatchlistKW;
            }
            set
            {
                if (this.strAddWatchlistKW != value)
                {
                    this.strAddWatchlistKW = value;
                }
            }
        }

        public String KeywordRemoveWatchlist
        {
            get
            {
                return this.strRemoveWatchlistKW;
            }
            set
            {
                if (this.strRemoveWatchlistKW != value)
                {
                    this.strRemoveWatchlistKW = value;
                }
            }
        }

        public Boolean IsCiresonIntegrationEnabled
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

        public Boolean IsCiresonKBSearchEnabled
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

        public Boolean IsCiresonROSearchEnabled
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

        public Boolean IsCiresonFirstResponseDateOnSuggestionsEnabled
        {
            get
            {
                return this.boolEnableCiresonFirstResponseDateOnSuggestions;
            }
            set
            {
                if (this.boolEnableCiresonFirstResponseDateOnSuggestions != value)
                {
                    this.boolEnableCiresonFirstResponseDateOnSuggestions = value;
                }
            }
        }

        public Int32 MinWordCountToSuggestRO
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

        public Int32 MinWordCountToSuggestKA
        {
            get
            {
                return this.strMinWordCountToSuggestKA;
            }
            set
            {
                if (this.strMinWordCountToSuggestKA != value)
                {
                    this.strMinWordCountToSuggestKA = value;
                }
            }
        }

        public String CiresonPortalURL
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
        
        public ManagementPackSecureReference RunAsAccountCiresonPortal
        {
            get
            {
                return runasaccountciresonportal;
            }
            set
            {
                if (this.runasaccountciresonportal != value)
                {
                    runasaccountciresonportal = value;
                }
            }
        }

        //announcements
        public Boolean IsAnnouncementIntegrationEnabled
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

        public Boolean IsSCSMAnnouncementsEnabled
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

        public Boolean IsCiresonAnnouncementsEnabled
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

        public String SCSMApprovedAnnouncers
        {
            get
            {
                return this.scsmApprovedAnnouncers;
            }
            set
            {
                if (this.scsmApprovedAnnouncers != value)
                {
                    this.scsmApprovedAnnouncers = value;
                }
            }
        }

        public String AuthorizedAnnouncementApproverType
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

        public String LowPriorityAnnouncementKeyword
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

        public String MediumPriorityAnnouncementKeyword
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

        public String HighPriorityAnnouncementKeyword
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

        public String LowPriorityExpiresInHours
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

        public String MediumPriorityExpiresInHours
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

        public String HighPriorityExpiresInHours
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

        public IDataItem SCSMApprovedAnnouncementGroup
        {
            get
            {
                return this.scsmApprovedAnnouncementGroup;
            }
            set
            {
                if (this.scsmApprovedAnnouncementGroup != value)
                {
                    this.scsmApprovedAnnouncementGroup = value;
                }
            }
        }
        public String SCSMApprovedGroupDisplayName { get; set; }
        public Guid SCSMApprovedGroupGUID { get; set; }
        
        //scom
        public String AuthorizedSCOMApproverType
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

        public String AuthorizedSCOMUsers
        {
            get
            {
                return this.strAuthorizedSCOMUsers;
            }
            set
            {
                if (this.strAuthorizedSCOMUsers != value)
                {
                    this.strAuthorizedSCOMUsers = value;
                }
            }
        }

        public Boolean IsSCOMIntegrationEnabled
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

        public IDataItem SCOMApprovedGroup
        {
            get
            {
                return this.scomApprovedGroup;
            }
            set
            {
                if (this.scomApprovedGroup != value)
                {
                    this.scomApprovedGroup = value;
                }
            }
        }
        public String SCOMApprovedGroupDisplayName { get; set; }
        public Guid SCOMApprovedGroupGUID { get; set; }

        //artificial intelligence, option 0: disabled
        public Boolean IsArtificialIntelligenceEnabled
        {
            get
            {
                return this.boolArtificialIntelligence;
            }
            set
            {
                if (this.boolArtificialIntelligence != value)
                {
                    this.boolArtificialIntelligence = value;
                }
            }
        }

        //artificial intelligence, option 1: ACS
        public Boolean IsAzureCognitiveServicesEnabled
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

        public String ACSRegion
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

        public String DefaultWorkItem
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

        public String AzureCognitiveServicesAPIKey
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

        public String AzureCognitiveServicesBoundaries
        {
            get
            {
                return this.strBoundaries;
            }
            set
            {
                if (this.strBoundaries != value)
                {
                    this.strBoundaries = value;
                }
            }
        }
        

        public ManagementPackProperty ACSIncidentSentimentDecExtension
        {
            get
            {
                return incidentACSSentimentExtensionDec;
            }
            set
            {
                if (this.incidentACSSentimentExtensionDec != value)
                {
                    incidentACSSentimentExtensionDec = value;
                }
            }
        }

        public ManagementPackProperty ACSServiceRequestSentimentDecExtension
        {
            get
            {
                return serviceRequestACSSentimentExtensionDec;
            }
            set
            {
                if (this.serviceRequestACSSentimentExtensionDec != value)
                {
                    serviceRequestACSSentimentExtensionDec = value;
                }
            }
        }

        public String MinimumPercentToCreateServiceRequest
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

        public IList<ManagementPackEnumeration> IncidentImpactEnums { get; set; }
        public ManagementPackEnumeration IncidentImpactEnum
        {
            get
            {
                return incidentImpactEnum;
            }
            set
            {
                if (this.incidentImpactEnum != value)
                {
                    incidentImpactEnum = value;
                }
            }
        }

        public IList<ManagementPackEnumeration> IncidentUrgencyEnums { get; set; }
        public ManagementPackEnumeration IncidentUrgencyEnum
        {
            get
            {
                return incidentUrgencyEnum;
            }
            set
            {
                if (this.incidentUrgencyEnum != value)
                {
                    incidentUrgencyEnum = value;
                }
            }
        }

        public IList<ManagementPackEnumeration> ServiceRequestUrgencyEnums { get; set; }
        public ManagementPackEnumeration ServiceRequestUrgencyEnum
        {
            get
            {
                return serviceRequestUrgencyEnum;
            }
            set
            {
                if (this.serviceRequestUrgencyEnum != value)
                {
                    serviceRequestUrgencyEnum = value;
                }
            }
        }

        public IList<ManagementPackEnumeration> ServiceRequestPriorityEnums { get; set; }
        public ManagementPackEnumeration ServiceRequestPriorityEnum
        {
            get
            {
                return serviceRequestPriorityEnum;
            }
            set
            {
                if (this.serviceRequestPriorityEnum != value)
                {
                    serviceRequestPriorityEnum = value;
                }
            }
        }

        public Boolean IsACSForKAEnabled
        {
            get
            {
                return this.boolEnableACSForKA;
            }
            set
            {
                if (this.boolEnableACSForKA != value)
                {
                    this.boolEnableACSForKA = value;
                }
            }
        }

        public Boolean IsACSForROEnabled
        {
            get
            {
                return this.boolEnableACSForRO;
            }
            set
            {
                if (this.boolEnableACSForRO != value)
                {
                    this.boolEnableACSForRO = value;
                }
            }
        }

        public Boolean IsACSForPriorityScoringEnabled
        {
            get
            {
                return this.boolEnableACSForPriorityScoring;
            }
            set
            {
                if (this.boolEnableACSForPriorityScoring != value)
                {
                    this.boolEnableACSForPriorityScoring = value;
                }
            }
        }

        //artificial intelligencee, option 2: Keyword matching
        public Boolean IsKeywordMatchingEnabled
        {
            get
            {
                return this.boolEnableKeywordMatching;
            }
            set
            {
                if (this.boolEnableKeywordMatching != value)
                {
                    this.boolEnableKeywordMatching = value;
                }
            }
        }

        public String KeywordMatchingRegexPattern
        {
            get
            {
                return this.strKWMatchingRegexPattern;
            }
            set
            {
                if (this.strKWMatchingRegexPattern != value)
                {
                    this.strKWMatchingRegexPattern = value;
                }
            }
        }

        public String KeywordWorkItemTypeOverride
        {
            get
            {
                return this.strKWWorkItemOverride;
            }
            set
            {
                if (this.strKWWorkItemOverride != value)
                {
                    this.strKWWorkItemOverride = value;
                }
            }
        }

        //artificial intelligencee, option 3: Azure Machine Learning
        public Boolean IsAzureMachineLearningEnabled
        {
            get
            {
                return this.boolEnableAzureMachineLearning;
            }
            set
            {
                if (this.boolEnableAzureMachineLearning != value)
                {
                    this.boolEnableAzureMachineLearning = value;
                }
            }
        }

        public String AzureMachineLearningAPIKey
        {
            get
            {
                return this.strAMLAPIKey;
            }
            set
            {
                if (this.strAMLAPIKey != value)
                {
                    this.strAMLAPIKey = value;
                }
            }
        }

        public String AzureMachineLearningURL
        {
            get
            {
                return this.strAMLWebServiceURL;
            }
            set
            {
                if (this.strAMLWebServiceURL != value)
                {
                    this.strAMLWebServiceURL = value;
                }
            }
        }

        public IList<ManagementPackProperty> IncidentDecExtensions { get; set; }
        public IList<ManagementPackProperty> IncidentEnumExtensions { get; set; }
        public IList<ManagementPackProperty> IncidentStringExtensions { get; set; }
        public IList<ManagementPackProperty> ServiceRequestDecExtensions { get; set; }
        public IList<ManagementPackProperty> ServiceRequestEnumExtensions { get; set; }
        public IList<ManagementPackProperty> ServiceRequestStringExtensions { get; set; }
        public ManagementPackProperty AMLIncidentConfidenceDecExtension
        {
            get
            {
                return incidentAMLWIConfidenceExtensionDec;
            }
            set
            {
                if (this.incidentAMLWIConfidenceExtensionDec != value)
                {
                    incidentAMLWIConfidenceExtensionDec = value;
                }
            }
        }
        public ManagementPackProperty AMLIncidentWIPredictionExtension
        {
            get
            {
                return incidentAMLWITypePredictionExtensionGUID;
            }
            set
            {
                if (this.incidentAMLWITypePredictionExtensionGUID != value)
                {
                    incidentAMLWITypePredictionExtensionGUID = value;
                }
            }
        }
        public ManagementPackProperty AMLServiceRequestWIPredictionExtension
        {
            get
            {
                return serviceRequestAMLWITypePredictionExtensionGUID;
            }
            set
            {
                if (this.serviceRequestAMLWITypePredictionExtensionGUID != value)
                {
                    serviceRequestAMLWITypePredictionExtensionGUID = value;
                }
            }
        }
        public ManagementPackProperty AMLIncidentClassificationConfidenceDecExtension
        {
            get
            {
                return incidentAMLClassifcationConfidenceExtensionDec;
            }
            set
            {
                if (this.incidentAMLClassifcationConfidenceExtensionDec != value)
                {
                    incidentAMLClassifcationConfidenceExtensionDec = value;
                }
            }
        }
        public ManagementPackProperty AMLIncidentSupportGroupConfidenceDecExtension
        {
            get
            {
                return incidentAMLSupportGroupIConfidenceExtensionDec;
            }
            set
            {
                if (this.incidentAMLSupportGroupIConfidenceExtensionDec != value)
                {
                    incidentAMLSupportGroupIConfidenceExtensionDec = value;
                }
            }
        }
        public ManagementPackProperty AMLIncidentClassificationPredictionEnumExtension
        {
            get
            {
                return incidentAMLClassificationPredictionExtensionEnum; ;
            }
            set
            {
                if (this.incidentAMLClassificationPredictionExtensionEnum != value)
                {
                    incidentAMLClassificationPredictionExtensionEnum = value;
                }
            }
        }
        public ManagementPackProperty AMLIncidentSupportGroupPredictionEnumExtension
        {
            get
            {
                return incidentAMLSupportGroupPredictionExtensionEnum;
            }
            set
            {
                if (this.incidentAMLSupportGroupPredictionExtensionEnum != value)
                {
                    incidentAMLSupportGroupPredictionExtensionEnum = value;
                }
            }
        }

        public ManagementPackProperty AMLServiceRequestConfidenceDecExtension
        {
            get
            {
                return serviceRequestAMLWIConfidenceExtensionDec;
            }
            set
            {
                if (this.serviceRequestAMLWIConfidenceExtensionDec != value)
                {
                    serviceRequestAMLWIConfidenceExtensionDec = value;
                }
            }
        }
        public ManagementPackProperty AMLServiceRequestClassificationConfidenceDecExtension
        {
            get
            {
                return serviceRequestAMLClassifcationConfidenceExtensionDec;
            }
            set
            {
                if (this.serviceRequestAMLClassifcationConfidenceExtensionDec != value)
                {
                    serviceRequestAMLClassifcationConfidenceExtensionDec = value;
                }
            }
        }
        public ManagementPackProperty AMLServiceRequestSupportGroupConfidenceDecExtension
        {
            get
            {
                return serviceRequestAMLSupportGroupIConfidenceExtensionDec;
            }
            set
            {
                if (this.serviceRequestAMLSupportGroupIConfidenceExtensionDec != value)
                {
                    serviceRequestAMLSupportGroupIConfidenceExtensionDec = value;
                }
            }
        }
        public ManagementPackProperty AMLServiceRequestClassificationPredictionEnumExtension
        {
            get
            {
                return serviceRequestAMLClassificationPredictionExtensionEnum;
            }
            set
            {
                if (this.serviceRequestAMLClassificationPredictionExtensionEnum != value)
                {
                    serviceRequestAMLClassificationPredictionExtensionEnum = value;
                }
            }
        }
        public ManagementPackProperty AMLServiceRequestSupportGroupPredictionEnumExtension
        {
            get
            {
                return serviceRequestAMLSupportGroupPredictionExtensionEnum;
            }
            set
            {
                if (this.serviceRequestAMLSupportGroupPredictionExtensionEnum != value)
                {
                    serviceRequestAMLSupportGroupPredictionExtensionEnum = value;
                }
            }
        }

        public String AzureMachineLearningWIConfidence
        {
            get
            {
                return this.decAMLWIMinConfidence;
            }
            set
            {
                if (this.decAMLWIMinConfidence != value)
                {
                    this.decAMLWIMinConfidence = value;
                }
            }
        }

        public String AzureMachineLearningClassificationConfidence
        {
            get
            {
                return this.decAMLClassificationMinConfidence;
            }
            set
            {
                if (this.decAMLClassificationMinConfidence != value)
                {
                    this.decAMLClassificationMinConfidence = value;
                }
            }
        }

        public String AzureMachineLearningSupportGroupConfidence
        {
            get
            {
                return this.decAMLSupportGroupMinConfidence;
            }
            set
            {
                if (this.decAMLSupportGroupMinConfidence != value)
                {
                    this.decAMLSupportGroupMinConfidence = value;
                }
            }
        }

        public String AzureMachineLearningAffectedConfigItemConfidence
        {
            get
            {
                return this.decAMLConfigItemMinConfidence;
            }
            set
            {
                if (this.decAMLConfigItemMinConfidence != value)
                {
                    this.decAMLConfigItemMinConfidence = value;
                }
            }
        }

        //azure translate
        public Boolean IsAzureTranslationEnabled
        {
            get
            {
                return this.boolEnableAzureTranslate;
            }
            set
            {
                if (this.boolEnableAzureTranslate != value)
                {
                    this.boolEnableAzureTranslate = value;
                }
            }
        }

        public String AzureTranslateAPIKey
        {
            get
            {
                return this.strAzureTranslateAPIKey;
            }
            set
            {
                if (this.strAzureTranslateAPIKey != value)
                {
                    this.strAzureTranslateAPIKey = value;
                }
            }
        }

        public String AzureTranslateDefaultLanguageCode
        {
            get
            {
                return this.strAzureTranslateDefaultLanguageCode;
            }
            set
            {
                if (this.strAzureTranslateDefaultLanguageCode != value)
                {
                    this.strAzureTranslateDefaultLanguageCode = value;
                }
            }
        }

        //azure vision
        public Boolean IsAzureVisionEnabled
        {
            get
            {
                return this.boolEnableAzureVision;
            }
            set
            {
                if (this.boolEnableAzureVision != value)
                {
                    this.boolEnableAzureVision = value;
                }
            }
        }

        public String AzureVisionAPIKey
        {
            get
            {
                return this.strAzureVisionAPIKey;
            }
            set
            {
                if (this.strAzureVisionAPIKey != value)
                {
                    this.strAzureVisionAPIKey = value;
                }
            }
        }

        public String AzureVisionRegion
        {
            get
            {
                return this.strAzureVisionRegion;
            }
            set
            {
                if (this.strAzureVisionRegion != value)
                {
                    this.strAzureVisionRegion = value;
                }
            }
        }

        //azure speech
        public Boolean IsAzureSpeechEnabled
        {
            get
            {
                return this.boolEnableAzureSpeech;
            }
            set
            {
                if (this.boolEnableAzureSpeech != value)
                {
                    this.boolEnableAzureSpeech = value;
                }
            }
        }

        public String AzureSpeechAPIKey
        {
            get
            {
                return this.strAzureSpeechAPIKey;
            }
            set
            {
                if (this.strAzureSpeechAPIKey != value)
                {
                    this.strAzureSpeechAPIKey = value;
                }
            }
        }

        public String AzureSpeechRegion
        {
            get
            {
                return this.strAzureSpeechRegion;
            }
            set
            {
                if (this.strAzureSpeechRegion != value)
                {
                    this.strAzureSpeechRegion = value;
                }
            }
        }
        
        //smlets exchange connector workflow
        public Boolean IsSMExcoWorkflowEnabled
        {
            get
            {
                return this.boolSMexcoWFEnabled;
            }
            set
            {
                if (this.boolSMexcoWFEnabled != value)
                {
                    this.boolSMexcoWFEnabled = value;
                }
            }
        }

        public String SMExcoIntervalSeconds
        {
            get
            {
                return this.strSMExcoWFInterval;
            }
            set
            {
                if (this.strSMExcoWFInterval != value)
                {
                    this.strSMExcoWFInterval = value;
                }
            }
        }

        //management pack guid
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

        //This Loads the settings that have been set in the management pack. In some cases for user protection, they load values from the function below
        //so on Save some kind of value is committed. This provides the means to suggest a value once, but ignore it afterwards.
        internal AdminSettingWizardData(EnterpriseManagementObject emoAdminSetting)
        {
            //Get the server name to connect to and then connect 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

            /*Get the SMLets Settings class and Additional Mailbox class by their GUID
            ##PowerShell SMlets equivalent:
                Get-SCSMClass -name "SMLets"
                Get-SCSMClass -id "a0022e87-75a8-65ee-4581-d923ff06a564"
                Get-SCSMClass -id "8b132799-638b-cc6f-469a-8f0638f917c0"
            */
            ManagementPackClass smletsExchangeConnectorSettingsClass = emg.EntityTypes.GetClass(new Guid("a0022e87-75a8-65ee-4581-d923ff06a564"));
            ManagementPackClass smletsExchangeConnectorAdditionalMailboxClass = emg.EntityTypes.GetClass(new Guid("8b132799-638b-cc6f-469a-8f0638f917c0"));
            
            //attempt to find the Cireson Portal MP. In doing so, a variable now exists in memory to identify the MPs existance.
            try
            {
                ManagementPack ciresonPortal = emg.ManagementPacks.GetManagementPack(new Guid("7d0aff53-6bf5-465d-87bd-c83412d5806a"));
                this.IsCiresonPortalMPImported = true;
            }
            catch { this.IsCiresonPortalMPImported = false; }
            
            //General Settings
            this.SCSMmanagementServer = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMmgmtServer"].ToString();
            this.WorkflowEmailAddress = emoAdminSetting[smletsExchangeConnectorSettingsClass, "WorkflowEmailAddress"].ToString();
            this.ExchangeAutodiscoverURL = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExchangeAutodiscoverURL"].ToString();
            try { this.IsExchangeOnline = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseExchangeOnline"].ToString()); }
            catch { this.IsExchangeOnline = false; }
            this.AzureClientID = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AzureClientID"].ToString();
            this.AzureTenantID = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AzureTenantID"].ToString();
            
            //Run as Account List
            var allSecureReferences = emg.Security.GetSecureReferences();
            List<ManagementPackSecureReference> securereferences = new List<ManagementPackSecureReference>();
            foreach (ManagementPackSecureReference secRef in allSecureReferences)
            {
                if ((secRef.GetManagementPack().Name == "ServiceManager.LinkingFramework.Configuration") || (secRef.GetManagementPack().Name == "ServiceManager.Core.Library"))
                {
                    securereferences.Add(secRef);
                }
            }
            this.SecureRunAsAccounts = securereferences.ToList();

            //Run as Account - Exchange
            try
            {
                this.RunAsAccountEWS = emg.Security.GetSecureReference(new Guid(emoAdminSetting[smletsExchangeConnectorSettingsClass, "SecureReferenceIdEWS"].ToString()));
            }
            catch
            { 
                //a run as account for exchange is not defined
            }

            //Autodiscover
            try { this.IsAutodiscoverEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseAutoDiscover"].ToString()); }
            catch { this.IsAutodiscoverEnabled = false; }

            //DLL Paths - EWS, Mimekit, PII regex, HTML Suyggestions, Custom Events
            this.EWSFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathEWSDLL"].ToString();
            this.SMExcoFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathSMExcoPS"].ToString();
            this.MimeKitFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathMimeKitDLL"].ToString();
            this.PIIRegexFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathPIIRegex"].ToString();
            this.HTMLSuggestionTemplatesFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathHTMLSuggestionTemplates"].ToString();
            this.CustomEventsFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathCustomEvents"].ToString();

            //Processing Logic
            //Create users not found in CMDB
            try { this.CreateUsersNotFoundtInCMDB = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "CreateUsersNotInCMDB"].ToString()); }
            catch { this.CreateUsersNotFoundtInCMDB = false; }

            //Processing Logic - Process on behalf of AD groups
            try { this.ProcessADGroupVote = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "VoteOnBehalfOfADGroup"].ToString()); }
            catch { this.ProcessADGroupVote = false; }

            //Processing Logic - Take requires Support Group membership as defined by Cireson Support Group Mapping
            try { this.EnforceSupportGroupTake = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "TakeRequiresSupportGroupMembership"].ToString()); }
            catch { this.EnforceSupportGroupTake = false; }

            //Processing Logic - include whole email on Action Log
            try { this.IncludeWholeEmail = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "IncludeWholeEmail"].ToString()); }
            catch { this.IncludeWholeEmail = false; }

            //Processing Logic - attach email to work item on the related items tab
            try { this.AttachEmailToWorkItem = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "AttachEmailToWorkItem"].ToString()); }
            catch { this.AttachEmailToWorkItem = false; }

            //Processing Logic - Mail processing
            try { this.ProcessCalendarAppointments = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "ProcessCalendarAppointments"].ToString()); }
            catch { this.ProcessCalendarAppointments = false; }
            try { this.ProcessMergeReplies = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "MergeReplies"].ToString()); }
            catch { this.ProcessMergeReplies = false; }
            try { this.ProcessClosedWorkItemsToNewWorkItems = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "CreateNewWorkItemIfWorkItemClosed"].ToString()); }
            catch { this.ProcessClosedWorkItemsToNewWorkItems = false; }
            try { this.ProcessEncryptedEmails = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "ProcessDigitallyEncryptedMessages"].ToString()); }
            catch { this.ProcessEncryptedEmails = false; }
            try { this.ProcessDigitallySignedEmails = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "ProcessDigitallySignedMessages"].ToString()); }
            catch { this.ProcessDigitallySignedEmails = false; }
            this.DynamicAnalystAssignment = emoAdminSetting[smletsExchangeConnectorSettingsClass, "DynamicAnalystAssignmentType"].ToString();
            this.CertificateStore = emoAdminSetting[smletsExchangeConnectorSettingsClass, "CertificateStore"].ToString();

            //Processing Logic - related user comment control
            this.ExternalIRCommentType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentTypeIR"].ToString();
            this.ExternalSRCommentType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentTypeSR"].ToString();
            this.ExternalIRCommentPrivate = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentPrivacyIR"].ToString();
            this.ExternalSRCommentPrivate = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentPrivacySR"].ToString();

            //Processing Logic - remove PII from work items per regex file
            try { this.RemovePII = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "RemovePII"].ToString()); }
            catch { this.RemovePII = false; }

            /*Processing Logic - Custom Support Group enum extensions
                Get the Manual Activity, Change Request, and Problem classes, then get all of their properties
                that do not come from a Microsoft management pack
            ##PowerShell SMlets equivalent:
                $maClass = Get-SCSMClass -name "System.WorkItem.Activity.ManualActivity$"
                $maClass.GetProperties(1, 1) | Where-Object {($_.SystemType.Name -eq "enum") -and ($_.identifier -notlike "*System.WorkItem.Activity.Library*")}
            */
            ManagementPackClass incidentRequestClass = emg.EntityTypes.GetClass(new Guid("a604b942-4c7b-2fb2-28dc-61dc6f465c68"));
            ManagementPackClass changeRequestClass = emg.EntityTypes.GetClass(new Guid("e6c9cf6e-d7fe-1b5d-216c-c3f5d2c7670c"));
            ManagementPackClass manualActivityClass = emg.EntityTypes.GetClass(new Guid("7ac62bd4-8fce-a150-3b40-16a39a61383d"));
            ManagementPackClass problemClass = emg.EntityTypes.GetClass(new Guid("422afc88-5eff-f4c5-f8f6-e01038cde67f"));
            IList<ManagementPackProperty> crPropertyList = changeRequestClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            IList<ManagementPackProperty> maPropertyList = manualActivityClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            IList<ManagementPackProperty> prPropertyList = problemClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            List<ManagementPackProperty> crTempPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> maTempPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> prTempPropertyList = new List<ManagementPackProperty>();
            //Load the Class Extension lists into a temporary list as long as the Property type is an enum and does not come from the stock Class
            //Load the Drop Down's currently Selected Item in the list if the stored GUID matches a property's respective id
            foreach (ManagementPackProperty crproperty in crPropertyList)
            {
                if ((crproperty.Type.ToString() == "enum") && ((crproperty.GetManagementPack().Name != "System.WorkItem.ChangeRequest.Library")))
                {
                    crTempPropertyList.Add(crproperty);
                }
                try
                {
                    if (crproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "CRSupportGroupGUID"].Value)
                    {
                        this.ChangeRequestSupportGroupExtensionEnum = crproperty;
                    }
                }
                catch { }
            }
            foreach (ManagementPackProperty maproperty in maPropertyList)
            {
                if ((maproperty.Type.ToString() == "enum") && (!(maproperty.Identifier.ToString().Contains("System.WorkItem.Activity.Library"))))
                {
                    maTempPropertyList.Add(maproperty);
                }
                try
                {
                    if (maproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "MASupportGroupGUID"].Value)
                    {
                        this.ManualActivitySupportGroupExtensionEnum = maproperty;
                    }
                }
                catch { }
            }
            foreach (ManagementPackProperty prproperty in prPropertyList)
            {
                if ((prproperty.Type.ToString() == "enum") && ((prproperty.GetManagementPack().Name != "System.WorkItem.Problem.Library")) && ((prproperty.GetManagementPack().Name != "System.WorkItem.Library")))
                {
                    prTempPropertyList.Add(prproperty);
                }
                try
                {
                    if (prproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "PRSupportGroupGUID"].Value)
                    {
                        this.ProblemSupportGroupExtensionEnum = prproperty;
                    }
                }
                catch { }
            }
            //Processing Logic - load the Class Extension Lists from the temporary lists
            this.ChangeRequestEnumExtensions = crTempPropertyList.ToList();
            this.ManualActivityEnumExtensions = maTempPropertyList.ToList();
            this.ProblemEnumExtensions = prTempPropertyList.ToList();

            //Processing Logic - Change Incident Status on Reply
            try { this.ChangeIncidentStatusOnReply = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "ChangeIncidentStatusOnReply"].ToString()); }
            catch { this.ChangeIncidentStatusOnReply = false; }

            /*Processing Logic - Get lists (enums) by their parent GUID to build the Drop Downs
             * https://github.com/SMLets/SMLets/blob/55f1bac3bc7a7011a461b24f6d7787ba89fe2624/SMLets.Shared/Code/EntityTypes.cs#L288
            ##PowerShell SMlets equivalents:
                $irStatusEnum = Get-SCSMEnumeration -id "89b34802-671e-e422-5e38-7dae9a413ef8"
                $irStatusEnum = Get-SCSMEnumeration -name "IncidentStatusEnum$"
                Get-SCSMChildEnumeration -Enumeration $irStatusEnum
            */
            //Order By, How-To via C# = https://docs.microsoft.com/en-us/dotnet/api/system.linq.enumerable.orderby?view=netframework-4.7.2
            this.IncidentStatusEnums = emg.EntityTypes.GetChildEnumerations(new Guid("89b34802-671e-e422-5e38-7dae9a413ef8"), TraversalDepth.Recursive);
            this.IncidentStatusEnums = this.IncidentStatusEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();

            //Processing Logic - get the currently defined Incident Statuses by Reply
            this.DefaultIncidentStatusOnAUReplyEnum = (ManagementPackEnumeration)emoAdminSetting[null, "IncidentStatusOnAffectedUserReply"].Value;
            this.DefaultIncidentStatusOnATReplyEnum = (ManagementPackEnumeration)emoAdminSetting[null, "IncidentStatusOnAssignedToReply"].Value;
            this.DefaultIncidentStatusOnRelReplyEnum = (ManagementPackEnumeration)emoAdminSetting[null, "IncidentStatusOnRelatedUserReply"].Value;

            //Processing Logic - get Incident Resolution Categories by GUID to build the Drop Down
            this.IncidentResolutionCategoryEnums = emg.EntityTypes.GetChildEnumerations(new Guid("72674491-02cb-1d90-a48f-1b269eb83602"), TraversalDepth.Recursive);
            this.IncidentResolutionCategoryEnums = this.IncidentResolutionCategoryEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();

            //Processing Logic - get Service Request Implementation Categories by GUID to build the Drop Down
            this.ServiceRequestImplementationCategoryEnums = emg.EntityTypes.GetChildEnumerations(new Guid("4ea37c27-9b24-615a-94da-510539371f4c"), TraversalDepth.Recursive);
            this.ServiceRequestImplementationCategoryEnums = this.ServiceRequestImplementationCategoryEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();

            //Processing Logic - get Problem Resolution Categories by GUID to build the Drop Down
            this.ProblemResolutionCategoryEnums = emg.EntityTypes.GetChildEnumerations(new Guid("52a0bfb0-b7e6-d16e-d06e-97ce62b4a335"), TraversalDepth.Recursive);
            this.ProblemResolutionCategoryEnums = this.ProblemResolutionCategoryEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();

            //Processing Logic - get the currently defined Incident Resolution Category by the currently stored GUID value
            this.DefaultIncidentResolutionCategoryEnum = (ManagementPackEnumeration)emoAdminSetting[null, "IncidentResolutionCategory"].Value;

            //Processing Logic - get the currently defined Service Request Implementation Category by the currently stored GUID value
            this.DefaultServiceRequestImplementationCategoryEnum = (ManagementPackEnumeration)emoAdminSetting[null, "ServiceRequestImplementationCategory"].Value;

            //Processing Logic - get the currently defined Problem Resolution Category by the currently stored GUID value
            this.DefaultProblemResolutionCategoryEnum = (ManagementPackEnumeration)emoAdminSetting[null, "ProblemResolutionCategory"].Value;

            //Multiple Mailbox Routing/Redirection
            try { this.UseMailboxRedirection = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseMailboxRedirection"].ToString()); }
            catch { this.UseMailboxRedirection = false; }

            //File Attachments
            if (this.MinFileAttachmentSize != null) { this.MinFileAttachmentSize = emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinimumFileAttachmentSize"].ToString(); }
            else { this.MinFileAttachmentSize = strMinFileSize; }

            try { this.IsMaxFileSizeAttachmentsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnforceFileAttachmentSettings"].ToString()); }
            catch { this.IsMaxFileSizeAttachmentsEnabled = false; }

            //Templates
            this.DefaultWorkItem = emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultWorkItemType"].ToString();

                /*
                ##PowerShell SMlets equivalent requires use of retrieving Type Projection, and 3 semi-C# style calls to connect to the management group,
                  perform a near identical ManagementPackTemplateCritieria search as seen below in C#, and finally a C# style call to retrieve said templates
                  as Get-SCSMObjectTemplate does not accept a critiera parameter:
                    $irTypeProjection = Get-SCSMTypeProjection -name "system.workitem.incident.projectiontype$"
                    $srTypeProjection = Get-SCSMTypeProjection -name "system.workitem.servicerequestprojection$"
                    $prTypeProjection = Get-SCSMTypeProjection -name "system.workitem.problem.projectiontype$"
                    $crTypeProjection = Get-SCSMTypeProjection -Name "system.workitem.changerequestprojection$"

                    $managementGroup = new-object Microsoft.EnterpriseManagement.EnterpriseManagementGroup "localhost"
                    $incidentTemplateCritiera = new-object Microsoft.EnterpriseManagement.Configuration.ManragementPackObjectTemplateCriteria("TypeID = '$($irTypeProjection.ID)'")
                    $serviceRequestTemplateCritiera = new-object Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateCriteria("TypeID = '$($srTypeProjection.ID)'")
                    $changeRequestTemplateCritiera = new-object Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateCriteria("TypeID = '$($crTypeProjection.ID)'")
                    $problemTemplateCritiera = new-object Microsoft.EnterpriseManagement.Configuration.ManagementPackObjectTemplateCriteria("TypeID = '$($prTypeProjection.ID)'")

                    $managementGroup.Templates.GetObjectTemplates($changeRequestTemplateCritiera)
                 * 
                 However what makes this different is that we need to retrieve Templates as they fall under any kind of Type Projection that could have
                 been created. To do this, we need to obtain all of the Type Projections per a Work Item class. Then cycle through the above for each
                 returned Type Projection.
                */
                //first, we need to build Type Projection criteria per the 4 Primary Work Item classes, their GUIDs are as follows within "Type ="
                ManagementPackTypeProjectionCriteria irTPCriteria = new ManagementPackTypeProjectionCriteria("Type = 'a604b942-4c7b-2fb2-28dc-61dc6f465c68'");
                ManagementPackTypeProjectionCriteria srTPCriteria = new ManagementPackTypeProjectionCriteria("Type = '04b69835-6343-4de2-4b19-6be08c612989'");
                ManagementPackTypeProjectionCriteria crTPCriteria = new ManagementPackTypeProjectionCriteria("Type = 'e6c9cf6e-d7fe-1b5d-216c-c3f5d2c7670c'");
                ManagementPackTypeProjectionCriteria prTPCriteria = new ManagementPackTypeProjectionCriteria("Type = '422afc88-5eff-f4c5-f8f6-e01038cde67f'");

                //second, we need to get all of the Type Projections that could be getting used for a particular class. This ensures that if anyone has
                //custom Type Projections or has modified stock forms through the Authoring tool that we make those Templates available for selection
                IList<ManagementPackTypeProjection> irTPs = emg.EntityTypes.GetTypeProjections(irTPCriteria);
                IList<ManagementPackTypeProjection> srTPs = emg.EntityTypes.GetTypeProjections(srTPCriteria);
                IList<ManagementPackTypeProjection> crTPs = emg.EntityTypes.GetTypeProjections(crTPCriteria);
                IList<ManagementPackTypeProjection> prTPs = emg.EntityTypes.GetTypeProjections(prTPCriteria);

                //third, create some empty lists to add all of the templates
                IList<ManagementPackObjectTemplate> irTemplateList = new List<ManagementPackObjectTemplate>();
                IList<ManagementPackObjectTemplate> srTemplateList = new List<ManagementPackObjectTemplate>();
                IList<ManagementPackObjectTemplate> crTemplateList = new List<ManagementPackObjectTemplate>();
                IList<ManagementPackObjectTemplate> prTemplateList = new List<ManagementPackObjectTemplate>();

                //get the Incident Templates by Type Projection ID/GUID
                foreach (ManagementPackTypeProjection tp in irTPs)
                {
                    ManagementPackObjectTemplateCriteria mpotcIncidents = new ManagementPackObjectTemplateCriteria(string.Format("TypeID = '{0}'", tp.Id.ToString()));
                    IList<ManagementPackObjectTemplate> retrievedTemplates = emg.Templates.GetObjectTemplates(mpotcIncidents);
                    foreach (ManagementPackObjectTemplate template in retrievedTemplates)
                    {
                        irTemplateList.Add(template);
                    }  
                }
                this.IncidentTemplates = irTemplateList;
                this.IncidentTemplates = this.IncidentTemplates.OrderBy(template => template.DisplayName).ToList();

                try {
                    Guid irTemplateGuid = (Guid)emoAdminSetting[null, "DefaultIncidentTemplateGUID"].Value;
                    this.DefaultIncidentTemplate = emg.Templates.GetObjectTemplate(irTemplateGuid);
                }
                catch { }

                //get the Service Request Templates by Type Projection ID/GUID
                foreach (ManagementPackTypeProjection tp in srTPs)
                {
                    ManagementPackObjectTemplateCriteria mpotcServiceRequests = new ManagementPackObjectTemplateCriteria(string.Format("TypeID = '{0}'", tp.Id.ToString()));
                    IList<ManagementPackObjectTemplate> retrievedTemplates = emg.Templates.GetObjectTemplates(mpotcServiceRequests);
                    foreach (ManagementPackObjectTemplate template in retrievedTemplates)
                    {
                        srTemplateList.Add(template);
                    } 
                }
                this.ServiceRequestTemplates = srTemplateList;
                this.ServiceRequestTemplates = this.ServiceRequestTemplates.OrderBy(template => template.DisplayName).ToList();

                try
                {
                    Guid srTemplateGuid = (Guid)emoAdminSetting[null, "DefaultServiceRequestTemplateGUID"].Value;
                    this.DefaultServiceRequestTemplate = emg.Templates.GetObjectTemplate(srTemplateGuid);
                }
                catch { }

                //get the Change Request Templates by Type Projection ID/GUID
                foreach (ManagementPackTypeProjection tp in crTPs)
                {
                    ManagementPackObjectTemplateCriteria mpotcChangeRequests = new ManagementPackObjectTemplateCriteria(string.Format("TypeID = '{0}'", tp.Id.ToString()));
                    IList<ManagementPackObjectTemplate> retrievedTemplates = emg.Templates.GetObjectTemplates(mpotcChangeRequests);
                    foreach (ManagementPackObjectTemplate template in retrievedTemplates)
                    {
                        crTemplateList.Add(template);
                    } 
                }
                this.ChangeRequestTemplates = crTemplateList;
                this.ChangeRequestTemplates = this.ChangeRequestTemplates.OrderBy(template => template.DisplayName).ToList();
                try
                {
                    Guid crTemplateGuid = (Guid)emoAdminSetting[null, "DefaultChangeRequestTemplateGUID"].Value;
                    this.DefaultChangeRequestTemplate = emg.Templates.GetObjectTemplate(crTemplateGuid);
                }
                catch { }

                //get the Problem Templates by Type Projection ID/GUID
                foreach (ManagementPackTypeProjection tp in prTPs)
                {
                    ManagementPackObjectTemplateCriteria mpotcProblems = new ManagementPackObjectTemplateCriteria(string.Format("TypeID = '{0}'", tp.Id.ToString()));
                    IList<ManagementPackObjectTemplate> retrievedTemplates = emg.Templates.GetObjectTemplates(mpotcProblems);
                    foreach (ManagementPackObjectTemplate template in retrievedTemplates)
                    {
                        prTemplateList.Add(template);
                    }
                }
                this.ProblemTemplates = prTemplateList;
                this.ProblemTemplates = this.ProblemTemplates.OrderBy(template => template.DisplayName).ToList();
                try
                {
                    Guid prTemplateGuid = (Guid)emoAdminSetting[null, "DefaultProblemTemplateGUID"].Value;
                    this.DefaultProblemTemplate = emg.Templates.GetObjectTemplate(prTemplateGuid);
                }
                catch { }

            /*Get the additional mailboxes by searching by the Additional Mailbox Class
            ##PowerShell SMlets equivalent:
                $additionalMailboxClass = Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings.AdditionalMailbox$"
                Get-SCSMObject -class $additionalMailboxClass
            */
            this.AdditionalMailboxes = emg.EntityObjects.GetObjectReader<EnterpriseManagementObject>(smletsExchangeConnectorAdditionalMailboxClass, ObjectQueryOptions.Default).ToList();

            //Keywords
            this.KeywordFrom = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordFrom"].ToString();
            this.KeywordAcknowledge = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordAcknowledge"].ToString();
            this.KeywordReactivate = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordReactivate"].ToString();
            this.KeywordResolve = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordResolved"].ToString();
            this.KeywordClose = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordClosed"].ToString();
            this.KeywordHold = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordHold"].ToString();
            this.KeywordTake = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordTake"].ToString();
            this.KeywordCancel = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordCancelled"].ToString();
            this.KeywordComplete = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordCompleted"].ToString();
            this.KeywordSkip = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordSkipped"].ToString();
            this.KeywordApprove = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordApprove"].ToString();
            this.KeywordReject = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordReject"].ToString();
            this.KeywordAnnouncement = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordAnnouncement"].ToString();
            this.KeywordHealth = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMKeywordHealth"].ToString();
            this.KeywordPrivate = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordPrivate"].ToString();

            //Cireson Integration
            this.CiresonPortalURL = emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonPortalURL"].ToString();
            try { this.MinWordCountToSuggestRO = Int32.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "NumberOfWordsToMatchFromEmailToCiresonRequestOffering"].ToString()); }
            catch { this.MinWordCountToSuggestRO = 1; }
            try { this.MinWordCountToSuggestKA = Int32.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "NumberOfWordsToMatchFromEmailToCiresonKnowledgeArticle"].ToString()); }
            catch { this.MinWordCountToSuggestKA = 1; }
            this.KeywordAddWatchlist = emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonKeywordWatchlistAdd"].ToString();
            this.KeywordRemoveWatchlist = emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonKeywordWatchlistRemove"].ToString();
            try { this.IsCiresonIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableCiresonIntegration"].ToString()); }
            catch { this.IsCiresonIntegrationEnabled = false; }
            try { this.IsCiresonFirstResponseDateOnSuggestionsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSetFirstResponseDateOnSuggestions"].ToString()); }
            catch { this.IsCiresonFirstResponseDateOnSuggestionsEnabled = false; }
            try { this.IsCiresonKBSearchEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchKnowledgeBase"].ToString()); }
            catch { this.IsCiresonKBSearchEnabled = false; }
            try { this.IsCiresonROSearchEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchRequestOfferings"].ToString()); }
            catch { this.IsCiresonROSearchEnabled = false; }
            //Run as Account - Cireson Portal
            try
            {
                this.RunAsAccountCiresonPortal = emg.Security.GetSecureReference(new Guid(emoAdminSetting[smletsExchangeConnectorSettingsClass, "SecureReferenceIdCiresonPortal"].ToString()));
            }
            catch
            {
                //a run as account for the cireson portal is not defined
            }

            //Announcements
            this.AuthorizedAnnouncementApproverType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMAnnouncementApprovedMemberType"].ToString();
            this.LowPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordLow"].ToString();
            this.MediumPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordMedium"].ToString();
            this.HighPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordHigh"].ToString();

            if (this.LowPriorityExpiresInHours != null) { this.LowPriorityExpiresInHours = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityLowExpirationInHours"].ToString(); }
            else { this.LowPriorityExpiresInHours = strLowPriorityExpInHours; }

            if (this.MediumPriorityExpiresInHours != null) { this.MediumPriorityExpiresInHours = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityNormalExpirationInHours"].ToString(); }
            else { this.MediumPriorityExpiresInHours = strMedPriorityExpInHours; }

            if (this.HighPriorityExpiresInHours != null) { this.HighPriorityExpiresInHours = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityCriticalExpirationInHours"].ToString(); }
            else { this.HighPriorityExpiresInHours = strHighPriorityExpInHours; }

            try { this.IsAnnouncementIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableAnnouncements"].ToString()); }
            catch { this.IsAnnouncementIntegrationEnabled = false; }

            try { this.IsSCSMAnnouncementsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSCSMAnnouncements"].ToString()); }
            catch { this.IsSCSMAnnouncementsEnabled = false; }

            try { this.IsCiresonAnnouncementsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableCiresonSCSMAnnouncements"].ToString()); }
            catch { this.IsCiresonAnnouncementsEnabled = false; }

            this.SCSMApprovedAnnouncers = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMApprovedAnnouncementUsers"].ToString();
            try
            {
                this.SCSMApprovedGroupGUID = (Guid)emoAdminSetting[null, "SCSMApprovedAnnouncementGroupGUID"].Value;
                EnterpriseManagementObject ScsmApprovedGroupEmoObject;
                ScsmApprovedGroupEmoObject = (EnterpriseManagementObject)emg.EntityObjects.GetObject<EnterpriseManagementObject>(this.SCSMApprovedGroupGUID, ObjectQueryOptions.Default);
                this.SCSMApprovedGroupDisplayName = "CURRENT ANNOUNCEMENT GROUP: " + ScsmApprovedGroupEmoObject.DisplayName;
            }
            catch
            {
                this.SCSMApprovedGroupDisplayName = "NO ANNOUNCEMENT GROUP DEFINED";
            }

            //SCOM Integration
            try { this.IsSCOMIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSCOMIntegration"].ToString()); }
            catch { this.IsSCOMIntegrationEnabled = false; }
            this.SCOMServer = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMmgmtServer"].ToString();
            this.AuthorizedSCOMApproverType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMApprovedMemberType"].ToString();
            this.AuthorizedSCOMUsers = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMApprovedUsers"].ToString();
            try
            {
                this.SCOMApprovedGroupGUID = (Guid)emoAdminSetting[null, "SCOMApprovedGroupGUID"].Value;
                EnterpriseManagementObject ScomApprovedGroupEmoObject;
                ScomApprovedGroupEmoObject = (EnterpriseManagementObject)emg.EntityObjects.GetObject<EnterpriseManagementObject>(this.SCOMApprovedGroupGUID, ObjectQueryOptions.Default);
                this.SCOMApprovedGroupDisplayName = "CURRENT APPROVED GROUP: " + ScomApprovedGroupEmoObject.DisplayName;
            }
            catch
            {
                this.SCOMApprovedGroupDisplayName = "NO SCOM GROUP DEFINED";
            }

            //Artificial Intelligence - Cognitive Services
            try { this.IsArtificialIntelligenceEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableArtificialIntelligence"].ToString()); }
            catch { this.IsArtificialIntelligenceEnabled = false; }
            
            //Artificial Intelligence - Cognitive Services
            this.ACSRegion = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTextAnalyticsRegion"].ToString();
            this.AzureCognitiveServicesAPIKey = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTextAnalyticsAPIKey"].ToString();
            this.AzureCognitiveServicesBoundaries = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSPriorityScoringBoundaries"].ToString();
            if (this.MinimumPercentToCreateServiceRequest != null) { this.MinimumPercentToCreateServiceRequest = emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinACSSentimentToCreateSR"].ToString(); }
            else { this.MinimumPercentToCreateServiceRequest = strMinimumPercentToCreateServiceRequest; }

            try { this.IsAzureCognitiveServicesEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSForNewWorkItem"].ToString()); }
            catch { this.IsAzureCognitiveServicesEnabled = false; }
            try { this.IsACSForKAEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSForCiresonKASuggestion"].ToString()); }
            catch { this.IsACSForKAEnabled = false; }
            try { this.IsACSForROEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSForCiresonROSuggestion"].ToString()); }
            catch { this.IsACSForROEnabled = false; }
            try { this.IsACSForPriorityScoringEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSPriorityScoring"].ToString()); }
            catch { this.IsACSForPriorityScoringEnabled = false; }

            this.IncidentImpactEnums = emg.EntityTypes.GetChildEnumerations(new Guid("11756265-f18e-e090-eed2-3aa923a4c872"), TraversalDepth.Recursive);
            this.IncidentImpactEnums = this.IncidentImpactEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();
            this.IncidentUrgencyEnums = emg.EntityTypes.GetChildEnumerations(new Guid("04b28bfb-8898-9af3-009b-979e58837852"), TraversalDepth.Recursive);
            this.IncidentUrgencyEnums = this.IncidentUrgencyEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();

            this.ServiceRequestUrgencyEnums = emg.EntityTypes.GetChildEnumerations(new Guid("eb35f771-8b0a-41aa-18fb-0432dfd957c4"), TraversalDepth.Recursive);
            this.ServiceRequestUrgencyEnums = this.ServiceRequestUrgencyEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();
            this.ServiceRequestPriorityEnums = emg.EntityTypes.GetChildEnumerations(new Guid("d55e65ea-fae9-f7db-0937-843bfb1367c0"), TraversalDepth.Recursive);
            this.ServiceRequestPriorityEnums = this.ServiceRequestPriorityEnums.OrderBy(enumitem => enumitem.DisplayName).ToList();

            //Artificial Intelligence - Keyword Matching
            try { this.IsKeywordMatchingEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableKeywordMatchForNewWorkItem"].ToString()); }
            catch { this.IsKeywordMatchingEnabled = false; }
            this.KeywordMatchingRegexPattern = emoAdminSetting[smletsExchangeConnectorSettingsClass, "KeywordMatchRegexForNewWorkItem"].ToString();
            this.KeywordWorkItemTypeOverride = emoAdminSetting[smletsExchangeConnectorSettingsClass, "KeywordMatchWorkItemType"].ToString();

            //Artificial Intelligence - Azure Machine Learning
            try { this.IsAzureMachineLearningEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableAML"].ToString()); }
            catch { this.IsAzureMachineLearningEnabled = false; }
            this.AzureMachineLearningAPIKey = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLAPIKey"].ToString();
            this.AzureMachineLearningURL = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLurl"].ToString();
            if (this.AzureMachineLearningWIConfidence != null) { this.AzureMachineLearningWIConfidence = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemType"].ToString(); }
            else { this.AzureMachineLearningWIConfidence = decAMLWIMinConfidence; }
            if (this.AzureMachineLearningClassificationConfidence != null) { this.AzureMachineLearningClassificationConfidence = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemClassification"].ToString(); }
            else { this.AzureMachineLearningClassificationConfidence = decAMLClassificationMinConfidence; }
            if (this.AzureMachineLearningSupportGroupConfidence != null) { this.AzureMachineLearningSupportGroupConfidence = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemSupportGroup"].ToString(); }
            else { this.AzureMachineLearningSupportGroupConfidence = decAMLSupportGroupMinConfidence; }
            if (this.AzureMachineLearningAffectedConfigItemConfidence != null) { this.AzureMachineLearningAffectedConfigItemConfidence = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceImpactedConfigItem"].ToString(); }
            else { this.AzureMachineLearningAffectedConfigItemConfidence = decAMLConfigItemMinConfidence; }

            ManagementPackClass incidentClass = emg.EntityTypes.GetClass(new Guid("a604b942-4c7b-2fb2-28dc-61dc6f465c68"));
            ManagementPackClass serviceRequestClass = emg.EntityTypes.GetClass(new Guid("04b69835-6343-4de2-4b19-6be08c612989"));
            IList<ManagementPackProperty> irPropertyList = incidentClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            IList<ManagementPackProperty> srPropertyList = serviceRequestClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            List<ManagementPackProperty> irTempPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> srTempPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> irTempEnumPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> srTempEnumPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> irTempStringPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> srTempStringPropertyList = new List<ManagementPackProperty>();
            
            //Load the Class Extension lists into temporary lists as long as the Property type is an dec/enum and does not come from the stock Class
            //Load the Drop Down's currently Selected Item in the list if the stored GUID matches a property's respective id
            foreach (ManagementPackProperty irproperty in irPropertyList)
            {
                //load all of the decimal properties that aren't from the Microsoft MP into a list
                if ((irproperty.Type.ToString() == "decimal") && ((irproperty.GetManagementPack().Name != "System.WorkItem.Incident.Library")))
                {
                    irTempPropertyList.Add(irproperty);
                }

                //of the decimal items being loaded, if they are the ones currently written to the MP set them as the selected item in the drop down
                if ((irproperty.Type.ToString() == "decimal") && ((irproperty.GetManagementPack().Name != "System.WorkItem.Incident.Library")))
                {
                    try
                    {
                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentConfidenceClassExtensionGUID"].Value)
                        {
                            this.AMLIncidentConfidenceDecExtension = irproperty;
                        }
                    }
                    catch { }
                    try
                    {
                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentClassificationConfidenceClassExtensionGUID"].Value)
                        {
                            this.AMLIncidentClassificationConfidenceDecExtension = irproperty;
                        }
                    }
                    catch { }
                    try
                    {
                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentSupportGroupConfidenceClassExtensionGUID"].Value)
                        {
                            this.AMLIncidentSupportGroupConfidenceDecExtension = irproperty;
                        }
                    }
                    catch { }
                    try
                    {

                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSentimentScoreIncidentClassExtensionGUID"].Value)
                        {
                            this.ACSIncidentSentimentDecExtension = irproperty;
                        }
                    }
                    catch { }
                }

                //load all of the enum properties that aren't from the Microsoft MP into a list
                if ((irproperty.Type.ToString() == "enum") && ((irproperty.GetManagementPack().Name != "System.WorkItem.Incident.Library")) && ((irproperty.GetManagementPack().Name != "System.WorkItem.Library")))
                {
                    irTempEnumPropertyList.Add(irproperty);

                    //if the enum property is the same guid as the one saved, make it the current selection in the list
                    try
                    {
                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentClassificationPredictionClassExtensionGUID"].Value)
                        {
                            this.AMLIncidentClassificationPredictionEnumExtension = irproperty;
                        }
                    }
                    catch { }

                    //if the enum property is the same guid as the one saved, make it the current selection in the list
                    try
                    {
                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentSupportGroupPredictionClassExtensionGUID"].Value)
                        {
                            this.AMLIncidentSupportGroupPredictionEnumExtension = irproperty;
                        }
                    }
                    catch { }
                }

                //load all of the string properties that aren't from the Microsoft MP into a list
                if ((irproperty.Type.ToString() == "string") && ((irproperty.GetManagementPack().Name != "System.WorkItem.Incident.Library")) && ((irproperty.GetManagementPack().Name != "System.WorkItem.Library")) && ((irproperty.GetManagementPack().Name != "System.Library")))
                {
                    irTempStringPropertyList.Add(irproperty);

                    //if the string property is the same guid as the one saved, make it the current selection in the list
                    try
                    {
                        if (irproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIRWorkItemTypePredictionClassExtensionGUID"].Value)
                        {
                            this.AMLIncidentWIPredictionExtension = irproperty;
                        }
                    }
                    catch { }
                }
            }

            //Load the Class Extension lists into temporary lists as long as the Property type is an dec/enum and does not come from the stock Class
            //Load the Drop Down's currently Selected Item in the list if the stored GUID matches a property's respective id
            foreach (ManagementPackProperty srproperty in srPropertyList)
            {                
                if ((srproperty.Type.ToString() == "decimal") && ((srproperty.GetManagementPack().Name != "System.WorkItem.ServiceRequest.Library")))
                {
                    srTempPropertyList.Add(srproperty);

                    try
                    {
                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestConfidenceClassExtensionGUID"].Value)
                        {
                            this.AMLServiceRequestConfidenceDecExtension = srproperty;
                        }
                    }
                    catch { }
                    try
                    {
                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestClassificationConfidenceClassExtensionGUID"].Value)
                        {
                            this.AMLServiceRequestClassificationConfidenceDecExtension = srproperty;
                        }
                    }
                    catch { }
                    try
                    {
                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestSupportGroupConfidenceClassExtensionGUID"].Value)
                        {
                            this.AMLServiceRequestSupportGroupConfidenceDecExtension = srproperty;
                        }
                    }
                    catch { }
                    try
                    {

                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSentimentScoreServiceRequestClassExtensionGUID"].Value)
                        {
                            this.ACSServiceRequestSentimentDecExtension = srproperty;
                        }
                    }
                    catch { }
                }

                //load all of the enum properties that aren't from the Microsoft MP into a list
                if ((srproperty.Type.ToString() == "enum") && ((srproperty.GetManagementPack().Name != "System.WorkItem.ServiceRequest.Library")))
                {
                    srTempEnumPropertyList.Add(srproperty);

                    //if the enum property is the same guid as the one saved, make it the current selection in the list
                    try
                    {
                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestClassificationPredictionClassExtensionGUID"].Value)
                        {
                            this.AMLServiceRequestClassificationPredictionEnumExtension = srproperty;
                        }
                    }
                    catch { }

                    //if the enum property is the same guid as the one saved, make it the current selection in the list
                    try
                    {
                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestSupportGroupPredictionClassExtensionGUID"].Value)
                        {
                            this.AMLServiceRequestSupportGroupPredictionEnumExtension = srproperty;
                        }
                    }
                    catch { }
                }

                //load all of the string properties that aren't from the Microsoft MP into a list
                if ((srproperty.Type.ToString() == "string") && ((srproperty.GetManagementPack().Name != "System.WorkItem.ServiceRequest.Library")) && ((srproperty.GetManagementPack().Name != "System.WorkItem.Library")) && ((srproperty.GetManagementPack().Name != "System.Library")))
                {
                    srTempStringPropertyList.Add(srproperty);

                    //if the string property is the same guid as the one saved, make it the current selection in the list
                    try
                    {
                        if (srproperty.Id == (Guid)emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLSRWorkItemTypePredictionClassExtensionGUID"].Value)
                        {
                            this.AMLServiceRequestWIPredictionExtension = srproperty;
                        }
                    }
                    catch { }
                }
            }
            //Processing Logic - load the Class Extension Lists from the temporary lists
            this.IncidentDecExtensions = irTempPropertyList.ToList();
            this.IncidentDecExtensions = irTempPropertyList.OrderBy(irextensions => irextensions.DisplayName).ToList();
            this.IncidentEnumExtensions = irTempEnumPropertyList.ToList();
            this.IncidentEnumExtensions = irTempEnumPropertyList.OrderBy(irextensions => irextensions.DisplayName).ToList();
            this.ServiceRequestDecExtensions = srTempPropertyList.ToList();
            this.IncidentStringExtensions = irTempStringPropertyList.OrderBy(irextensions => irextensions.DisplayName).ToList();
            this.IncidentStringExtensions = irTempStringPropertyList.ToList();
            this.ServiceRequestDecExtensions = srTempPropertyList.OrderBy(srextensions => srextensions.DisplayName).ToList();
            this.ServiceRequestEnumExtensions = srTempEnumPropertyList.ToList();
            this.ServiceRequestEnumExtensions = srTempEnumPropertyList.OrderBy(srextensions => srextensions.DisplayName).ToList();
            this.ServiceRequestStringExtensions = srTempStringPropertyList.ToList();
            this.ServiceRequestStringExtensions = srTempStringPropertyList.OrderBy(srextensions => srextensions.DisplayName).ToList();

            //azure translate
            try { this.IsAzureTranslationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSTranslate"].ToString()); }
            catch { this.IsAzureTranslationEnabled = false; }
            this.AzureTranslateAPIKey = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTranslateAPIKey"].ToString();
            this.AzureTranslateDefaultLanguageCode = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTranslateDefaultLanguageCode"].ToString();

            //azure vision
            try { this.IsAzureVisionEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSVision"].ToString()); }
            catch { this.IsAzureVisionEnabled = false; }
            this.AzureVisionAPIKey = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSVisionAPIKey"].ToString();
            this.AzureVisionRegion = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSVisionRegion"].ToString();

            //azure speech
            try { this.IsAzureSpeechEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSSpeech"].ToString()); }
            catch { this.IsAzureSpeechEnabled = false; }
            this.AzureSpeechAPIKey = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSpeechAPIKey"].ToString();
            this.AzureSpeechRegion = emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSpeechRegion"].ToString();

            //load the MP
            this.EnterpriseManagementObjectID = emoAdminSetting.Id;
        }

        //This Saves the values that have been set throughout the Wizard Pages back into the management pack so the
        //SMLets Exchange Connector PowerShell can read them via ((Get-SCSMObject -Class (Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings$")))
        public override void AcceptChanges(WizardMode wizardMode)
        {
            //Get the server name to connect to and connect 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

            //Get the SMLets Settings class and Additional Mailbox class by GUID
            ManagementPackClass smletsExchangeConnectorSettingsClass = emg.EntityTypes.GetClass(new Guid("a0022e87-75a8-65ee-4581-d923ff06a564"));
            ManagementPackClass smletsExchangeConnectorAdditionalMailboxClass = emg.EntityTypes.GetClass(new Guid("8b132799-638b-cc6f-469a-8f0638f917c0"));

            //Get the SMLets Exchange Connector Settings object using the object GUID 
            EnterpriseManagementObject emoAdminSetting = emg.EntityObjects.GetObject<EnterpriseManagementObject>(this.EnterpriseManagementObjectID, ObjectQueryOptions.Default);

            //General Settings
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMmgmtServer"].Value = this.SCSMmanagementServer;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "WorkflowEmailAddress"].Value = this.WorkflowEmailAddress;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseAutoDiscover"].Value = this.IsAutodiscoverEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExchangeAutodiscoverURL"].Value = this.ExchangeAutodiscoverURL;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseExchangeOnline"].Value = this.IsExchangeOnline;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AzureClientID"].Value = this.AzureClientID;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AzureTenantID"].Value = this.AzureTenantID;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CreateUsersNotInCMDB"].Value = this.CreateUsersNotFoundtInCMDB;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "IncludeWholeEmail"].Value = this.IncludeWholeEmail;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AttachEmailToWorkItem"].Value = this.AttachEmailToWorkItem;
            
            //Run As Account - Exchange
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SecureReferenceIdEWS"].Value = this.RunAsAccountEWS.Id.ToString();

            //DLL Paths - EWS, Mimekit, PII regex, HTML Suyggestions, Custom Events
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathEWSDLL"].Value = this.EWSFilePath;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathSMExcoPS"].Value = this.SMExcoFilePath;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathMimeKitDLL"].Value = this.MimeKitFilePath;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathPIIRegex"].Value = this.PIIRegexFilePath;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathHTMLSuggestionTemplates"].Value = this.HTMLSuggestionTemplatesFilePath;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathCustomEvents"].Value = this.CustomEventsFilePath;

            //Processing Logic
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "VoteOnBehalfOfADGroup"].Value = this.ProcessADGroupVote;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "TakeRequiresSupportGroupMembership"].Value = this.EnforceSupportGroupTake;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "RemovePII"].Value = this.RemovePII;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseMailboxRedirection"].Value = this.UseMailboxRedirection;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ProcessCalendarAppointments"].Value = this.ProcessCalendarAppointments;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "MergeReplies"].Value = this.ProcessMergeReplies;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "DynamicAnalystAssignmentType"].Value = this.DynamicAnalystAssignment;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CreateNewWorkItemIfWorkItemClosed"].Value = this.ProcessClosedWorkItemsToNewWorkItems;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ProcessDigitallyEncryptedMessages"].Value = this.ProcessEncryptedEmails;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ProcessDigitallySignedMessages"].Value = this.ProcessDigitallySignedEmails;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CertificateStore"].Value = this.CertificateStore;

            //Processing Logic - Set Incident Resolution, Service Request Implentation, and Problem Resolution Categories by their Enum Name
            emoAdminSetting[null, "IncidentResolutionCategory"].Value = this.DefaultIncidentResolutionCategoryEnum;
            emoAdminSetting[null, "ServiceRequestImplementationCategory"].Value = this.DefaultServiceRequestImplementationCategoryEnum;
            emoAdminSetting[null, "ProblemResolutionCategory"].Value = this.DefaultProblemResolutionCategoryEnum;

            //Processing Logic - Change Incident Status on Reply
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ChangeIncidentStatusOnReply"].Value = this.ChangeIncidentStatusOnReply;
            emoAdminSetting[null, "IncidentStatusOnAffectedUserReply"].Value = this.DefaultIncidentStatusOnAUReplyEnum;
            emoAdminSetting[null, "IncidentStatusOnAssignedToReply"].Value = this.DefaultIncidentStatusOnATReplyEnum;
            emoAdminSetting[null, "IncidentStatusOnRelatedUserReply"].Value = this.DefaultIncidentStatusOnRelReplyEnum;

            //Processing Logic - Related user comment control, if nothing was set ensure we default like the stock connector Analyst Comments
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentTypeIR"].Value = this.ExternalIRCommentType;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentTypeSR"].Value = this.ExternalSRCommentType;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentPrivacyIR"].Value = this.ExternalIRCommentPrivate;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ExternalPartyCommentPrivacySR"].Value = this.ExternalSRCommentPrivate;

            //Processing Logic - Extended Support Groups
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "CRSupportGroupGUID"].Value = this.ChangeRequestSupportGroupExtensionEnum.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "MASupportGroupGUID"].Value = this.ManualActivitySupportGroupExtensionEnum.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "PRSupportGroupGUID"].Value = this.ProblemSupportGroupExtensionEnum.Id; }
            catch { }

            //File Attachments
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinimumFileAttachmentSize"].Value = this.MinFileAttachmentSize;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnforceFileAttachmentSettings"].Value = this.IsMaxFileSizeAttachmentsEnabled;

            //Default Templates
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultWorkItemType"].Value = this.DefaultWorkItem;
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultIncidentTemplateGUID"].Value = this.DefaultIncidentTemplate.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultServiceRequestTemplateGUID"].Value = this.DefaultServiceRequestTemplate.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultChangeRequestTemplateGUID"].Value = this.DefaultChangeRequestTemplate.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultProblemTemplateGUID"].Value = this.DefaultProblemTemplate.Id; }
            catch { }

            //Multiple Mailbox Routing/Redirection - foreach entry returned from the form either Create or Update
            if (MultipleMailboxesToAdd != null)
            {
                foreach (string mailObject in MultipleMailboxesToAdd)
                {
                    string[] sections = mailObject.Split(';');
                    try
                    {
                        //retrieve the mailbox in the loop by its GUID
                        EnterpriseManagementObject mailboxExists = emg.EntityObjects.GetObject<EnterpriseManagementObject>(new Guid(sections[0]), ObjectQueryOptions.Default);

                        /*update the mailbox
                        ##PowerShell SMlets equivalent:
                            $additionalMailboxClass = Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings.AdditionalMailbox$"
                            $additionalMailbox = Get-SCSMObject -Class $additionalMailboxClass -filter "MailboxAddress -eq 'user@corp.net'"
                            Set-SCSMObject -SMObject $additionalMailbox -Property MailboxAddress -Value "differentuser@corp.net"
                        */
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxAddress"].Value = sections[1].ToString();
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxTemplateWorkItemType"].Value = sections[2].ToString();
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxIRTemplateDisplayName"].Value = sections[3].ToString();
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxIRTemplateGUID"].Value = new Guid(sections[4]);
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxSRTemplateDisplayName"].Value = sections[5].ToString();
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxSRTemplateGUID"].Value = new Guid(sections[6]);
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxCRTemplateDisplayName"].Value = sections[7].ToString();
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxCRTemplateGUID"].Value = new Guid(sections[8]);
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxPRTemplateDisplayName"].Value = sections[9].ToString();
                        mailboxExists[smletsExchangeConnectorAdditionalMailboxClass, "MailboxPRTemplateGUID"].Value = new Guid(sections[10]);
                        mailboxExists.Commit();
                    }
                    catch
                    {
                        /*mailbox doesn't exist, create it
                        ##PowerShell SMlets equivalent:
                            $additionalMailboxClass = Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings.AdditionalMailbox$"
                            New-SCSMObject -Class $additionalMailboxClass -propertyhashtable @{}
                        */
                        CreatableEnterpriseManagementObject emoMailboxObject = new CreatableEnterpriseManagementObject(emg, smletsExchangeConnectorAdditionalMailboxClass);
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxID"].Value = new Guid(sections[0]);
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxAddress"].Value = sections[1].ToString();
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxTemplateWorkItemType"].Value = sections[2].ToString();
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxIRTemplateDisplayName"].Value = sections[3].ToString();
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxIRTemplateGUID"].Value = new Guid(sections[4]);
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxSRTemplateDisplayName"].Value = sections[5].ToString();
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxSRTemplateGUID"].Value = new Guid(sections[6]);
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxCRTemplateDisplayName"].Value = sections[7].ToString();
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxCRTemplateGUID"].Value = new Guid(sections[8]);
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxPRTemplateDisplayName"].Value = sections[9].ToString();
                        emoMailboxObject[smletsExchangeConnectorAdditionalMailboxClass, "MailboxPRTemplateGUID"].Value = new Guid(sections[10]);
                        emoMailboxObject.Commit();
                    }
                }
            }

            //Multiple Mailbox Routing/Redirection - foreach entry returned from the form Delete
            if (MultipleMailboxesToRemove != null)
            {
                try
                {
                    /*SMLets - Remove-SCSMObject
                     SMLets requires the -Force parameter when the object in question does not derive from the Configuration Item class
                    ##PowerShell SMlets equivalent:
                        $additionalMailboxClass = Get-SCSMClass -Name "SMLets.Exchange.Connector.AdminSettings.AdditionalMailbox$"
                        $additionalMailbox = Get-SCSMObject -Class $additionalMailboxClass -filter "MailboxAddress -eq 'user@corp.net'"
                        $additionalMailbox | Remove-SCSMObject -Force
                        https://github.com/SMLets/SMLets/blob/55f1bac3bc7a7011a461b24f6d7787ba89fe2624/SMLets.Shared/Code/EntityObjects.cs#L1402
                    */
                    IncrementalDiscoveryData idd = new IncrementalDiscoveryData();
                    foreach (string mailboxToDeleteGuid in MultipleMailboxesToRemove)
                    {
                        EnterpriseManagementObject mailboxToDelete = emg.EntityObjects.GetObject<EnterpriseManagementObject>(new Guid(mailboxToDeleteGuid), ObjectQueryOptions.Default);
                        if (mailboxToDelete != null)
                        {
                            //mark the object for deletion
                            idd.Remove(mailboxToDelete);
                        }
                    }
                    //delete the mailbox object(s)
                    idd.Commit(emg);
                }
                catch
                {

                }
            }

            //Keywords
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordFrom"].Value = this.KeywordFrom;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordAcknowledge"].Value = this.KeywordAcknowledge;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordReactivate"].Value = this.KeywordReactivate;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordResolved"].Value = this.KeywordResolve;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordClosed"].Value = this.KeywordClose;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordHold"].Value = this.KeywordHold;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordTake"].Value = this.KeywordTake;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordCancelled"].Value = this.KeywordCancel;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordCompleted"].Value = this.KeywordComplete;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordSkipped"].Value = this.KeywordSkip;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordApprove"].Value = this.KeywordApprove;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordReject"].Value = this.KeywordReject;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordAnnouncement"].Value = this.KeywordAnnouncement;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMKeywordHealth"].Value = this.KeywordHealth;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMKeywordPrivate"].Value = this.KeywordPrivate;

            //Cireson Integration
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableCiresonIntegration"].Value = this.IsCiresonIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonKeywordWatchlistAdd"].Value = this.KeywordAddWatchlist;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonKeywordWatchlistRemove"].Value = this.KeywordRemoveWatchlist;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchKnowledgeBase"].Value = this.IsCiresonKBSearchEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchRequestOfferings"].Value = this.IsCiresonROSearchEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonPortalURL"].Value = this.CiresonPortalURL;
            if (!(this.CiresonPortalURL.EndsWith("/"))) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonPortalURL"].Value = this.CiresonPortalURL + "/"; }
            else { emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonPortalURL"].Value = this.CiresonPortalURL; }
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "NumberOfWordsToMatchFromEmailToCiresonRequestOffering"].Value = this.MinWordCountToSuggestRO;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "NumberOfWordsToMatchFromEmailToCiresonKnowledgeArticle"].Value = this.MinWordCountToSuggestKA;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSetFirstResponseDateOnSuggestions"].Value = this.IsCiresonFirstResponseDateOnSuggestionsEnabled;
            //Run As Account - Cireson
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SecureReferenceIdCiresonPortal"].Value = this.RunAsAccountCiresonPortal.Id.ToString();

            //Announcements
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableAnnouncements"].Value = this.IsAnnouncementIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSCSMAnnouncements"].Value = this.IsSCSMAnnouncementsEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMAnnouncementApprovedMemberType"].Value = this.AuthorizedAnnouncementApproverType;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableCiresonSCSMAnnouncements"].Value = this.IsCiresonAnnouncementsEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordLow"].Value = this.LowPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordMedium"].Value = this.MediumPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordHigh"].Value = this.HighPriorityAnnouncementKeyword;
            try { decimal.Parse(this.LowPriorityExpiresInHours); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityLowExpirationInHours"].Value = this.LowPriorityExpiresInHours; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityLowExpirationInHours"].Value = 0; }
            try { decimal.Parse(this.MediumPriorityExpiresInHours); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityNormalExpirationInHours"].Value = this.MediumPriorityExpiresInHours; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityNormalExpirationInHours"].Value = 0; }
            try { decimal.Parse(this.HighPriorityExpiresInHours); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityCriticalExpirationInHours"].Value = this.HighPriorityExpiresInHours; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityCriticalExpirationInHours"].Value = 0; }
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMApprovedAnnouncementUsers"].Value = this.SCSMApprovedAnnouncers;
            //if the Announcement group is set, don't null it
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMApprovedAnnouncementGroupGUID"].Value = this.SCSMApprovedAnnouncementGroup["Id"]; }
            catch { }

            //SCOM Integration
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSCOMIntegration"].Value = this.IsSCOMIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMmgmtServer"].Value = this.SCOMServer;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMApprovedMemberType"].Value = this.AuthorizedSCOMApproverType;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMApprovedUsers"].Value = this.AuthorizedSCOMUsers;
            //if the SCOM approved group is set, don't null it
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCOMApprovedGroupGUID"].Value = this.SCOMApprovedGroup["Id"]; }
            catch { }

            //Artifical Intelligence - enabled/disabled
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableArtificialIntelligence"].Value = this.IsArtificialIntelligenceEnabled;

            //Artifical Intelligence - Azure Cognitive Services
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSForNewWorkItem"].Value = this.IsAzureCognitiveServicesEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTextAnalyticsRegion"].Value = this.ACSRegion;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTextAnalyticsAPIKey"].Value = this.AzureCognitiveServicesAPIKey;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSPriorityScoringBoundaries"].Value = this.AzureCognitiveServicesBoundaries.ToString();
            try { decimal.Parse(this.MinimumPercentToCreateServiceRequest); emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinACSSentimentToCreateSR"].Value = this.MinimumPercentToCreateServiceRequest; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinACSSentimentToCreateSR"].Value = 95; }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSentimentScoreIncidentClassExtensionGUID"].Value = this.ACSIncidentSentimentDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSentimentScoreServiceRequestClassExtensionGUID"].Value = this.ACSServiceRequestSentimentDecExtension.Id; }
            catch { }
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSForCiresonKASuggestion"].Value = this.IsACSForKAEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSForCiresonROSuggestion"].Value = this.IsACSForROEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSPriorityScoring"].Value = this.IsACSForPriorityScoringEnabled;

            //Artifical Intelligence - Keyword Match
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableKeywordMatchForNewWorkItem"].Value = this.IsKeywordMatchingEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "KeywordMatchRegexForNewWorkItem"].Value = this.KeywordMatchingRegexPattern;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "KeywordMatchWorkItemType"].Value = this.KeywordWorkItemTypeOverride;

            //Artifical Intelligence - Azure Machine Learning
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableAML"].Value = this.IsAzureMachineLearningEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLAPIKey"].Value = this.AzureMachineLearningAPIKey;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLurl"].Value = this.AzureMachineLearningURL;
            try { decimal.Parse(this.AzureMachineLearningWIConfidence); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemType"].Value = this.AzureMachineLearningWIConfidence; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemType"].Value = 95; }
            try { decimal.Parse(this.AzureMachineLearningClassificationConfidence); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemClassification"].Value = this.AzureMachineLearningClassificationConfidence; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemClassification"].Value = 95; }
            try { decimal.Parse(this.AzureMachineLearningSupportGroupConfidence); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemSupportGroup"].Value = this.AzureMachineLearningSupportGroupConfidence; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemSupportGroup"].Value = 95; }
            try { decimal.Parse(this.AzureMachineLearningAffectedConfigItemConfidence); emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceImpactedConfigItem"].Value = this.AzureMachineLearningAffectedConfigItemConfidence; }
            catch { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceImpactedConfigItem"].Value = 95; }

            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentConfidenceClassExtensionGUID"].Value = this.AMLIncidentConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIRWorkItemTypePredictionClassExtensionGUID"].Value = this.AMLIncidentWIPredictionExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentClassificationConfidenceClassExtensionGUID"].Value = this.AMLIncidentClassificationConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentSupportGroupConfidenceClassExtensionGUID"].Value = this.AMLIncidentSupportGroupConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentClassificationPredictionClassExtensionGUID"].Value = this.AMLIncidentClassificationPredictionEnumExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentSupportGroupPredictionClassExtensionGUID"].Value = this.AMLIncidentSupportGroupPredictionEnumExtension.Id; }
            catch { }

            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestConfidenceClassExtensionGUID"].Value = this.AMLServiceRequestConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLSRWorkItemTypePredictionClassExtensionGUID"].Value = this.AMLServiceRequestWIPredictionExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestClassificationConfidenceClassExtensionGUID"].Value = this.AMLServiceRequestClassificationConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestSupportGroupConfidenceClassExtensionGUID"].Value = this.AMLServiceRequestSupportGroupConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestClassificationPredictionClassExtensionGUID"].Value = this.AMLServiceRequestClassificationPredictionEnumExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestSupportGroupPredictionClassExtensionGUID"].Value = this.AMLServiceRequestSupportGroupPredictionEnumExtension.Id; }
            catch { }

            //azure translate
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSTranslate"].Value = this.IsAzureTranslationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTranslateAPIKey"].Value = this.AzureTranslateAPIKey;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSTranslateDefaultLanguageCode"].Value = this.AzureTranslateDefaultLanguageCode;

            //azure vision
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSVision"].Value = this.IsAzureVisionEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSVisionAPIKey"].Value = this.AzureVisionAPIKey;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSVisionRegion"].Value = this.AzureVisionRegion;

            //azure speech
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableACSSpeech"].Value = this.IsAzureSpeechEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSpeechAPIKey"].Value = this.AzureSpeechAPIKey;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "ACSSpeechRegion"].Value = this.AzureSpeechRegion;
            
            //When creating a Run as Account in the UI, it is force stored in the unsealed "Service Manager Linking Framework Configuration" management pack
            //By working with this MP, we can work directly alongside native SCSM connectors and Run as Accounts
            //Credit to Travis Wright for an always detailed walkthrough of Run As Accounts in the SMLets notes
            //https://github.com/SMLets/SMLets/blob/795242d4a44cf8857957a7628cca67655e2e4252/SMLets.Shared/Code/Security.cs#L819
            try
            {
                //get the Service Manager Linking Framework Configuration unsealed MP and the rule. If the rule isn't found, the catch will engage
                ManagementPack scsmLFXConfig = emg.ManagementPacks.GetManagementPack(new Guid("50daaf82-06ce-cacb-8cf5-3950aebae0b0"));
                ManagementPackRule RunSMExco = scsmLFXConfig.GetRule("SMLets.Exchange.Connector.15d8b765a2f8b63ead14472f9b3c12f0");
                if (this.RunAsAccountEWS != null)
                {
                    //the rule exists and a run as account was selected. update rule
                    RunSMExco.Status = ManagementPackElementStatus.PendingUpdate;
                    
                    //Is the workflow enabled?
                    if (this.IsSMExcoWorkflowEnabled == true)
                    {
                        RunSMExco.Enabled = ManagementPackMonitoringLevel.@true;
                    }
                    else
                    {
                        RunSMExco.Enabled = ManagementPackMonitoringLevel.@false;
                    }

                    //set the workflow interval from the value in the GUI
                    RunSMExco.DataSourceCollection[0].Configuration = string.Format("\r\n <Scheduler>\r\n <SimpleReccuringSchedule>\r\n <Interval Unit=\"Seconds\">{0}</Interval>\r\n </SimpleReccuringSchedule>\r\n <ExcludeDates />\r\n </Scheduler>", this.SMExcoIntervalSeconds);
                    
                    //Get the Secure Reference's Management Pack's, Aliased Name, from the SCSM LFX unsealed mp
                    ManagementPack secRefMP = this.RunAsAccountEWS.GetManagementPack();
                    string mpAlias = null;
                    foreach (KeyValuePair<string, ManagementPackReference> reference in scsmLFXConfig.References)
                    {
                        if (secRefMP.Id == reference.Value.Id)
                        {
                            mpAlias = reference.Key;
                        }
                    }
                    
                    //set the Run As Account reference to use in the workflow
                    if (this.RunAsAccountEWS.Name.StartsWith("SecureReference"))
                    {
                        //if the secure reference actually begins with "SecureReference" it's defined in the Linking Framework Configuration MP. Just use it by name.
                        string EWSRunAsName = this.RunAsAccountEWS.Name;
                        RunSMExco.WriteActionCollection[0].Configuration = string.Format("\r\n<Subscription>\r\n <WindowsWorkflowConfiguration>\r\n <AssemblyName>SMLets.Exchange.Connector.Resources</AssemblyName>\r\n <WorkflowTypeName>SMLets.Exchange.Connector.Resources.RunScript</WorkflowTypeName>\r\n <WorkflowParameters>\r\n <WorkflowParameter Name=\"ExchangeDomain\" Type=\"string\">$RunAs[Name=\"{0}\"]/Domain$</WorkflowParameter><WorkflowParameter Name=\"ExchangeUsername\" Type=\"string\">$RunAs[Name=\"{0}\"]/UserName$</WorkflowParameter><WorkflowParameter Name=\"ExchangePassword\" Type=\"string\">$RunAs[Name=\"{0}\"]/Password$</WorkflowParameter></WorkflowParameters><RetryExceptions/><RetryDelaySeconds>60</RetryDelaySeconds><MaximumRunningTimeSeconds>300</MaximumRunningTimeSeconds></WindowsWorkflowConfiguration></Subscription>", EWSRunAsName);
                    }
                    else
                    {
                        //if it doesn't begin with SecureReference, it's defined in the Core MP which is already referenced in the Linking Framework Configuration MP
                        string EWSRunAsName = mpAlias + "!" + this.RunAsAccountEWS.Name;
                        RunSMExco.WriteActionCollection[0].Configuration = string.Format("\r\n<Subscription>\r\n <WindowsWorkflowConfiguration>\r\n <AssemblyName>SMLets.Exchange.Connector.Resources</AssemblyName>\r\n <WorkflowTypeName>SMLets.Exchange.Connector.Resources.RunScript</WorkflowTypeName>\r\n <WorkflowParameters>\r\n <WorkflowParameter Name=\"ExchangeDomain\" Type=\"string\">$RunAs[Name=\"{0}\"]/Domain$</WorkflowParameter><WorkflowParameter Name=\"ExchangeUsername\" Type=\"string\">$RunAs[Name=\"{0}\"]/UserName$</WorkflowParameter><WorkflowParameter Name=\"ExchangePassword\" Type=\"string\">$RunAs[Name=\"{0}\"]/Password$</WorkflowParameter></WorkflowParameters><RetryExceptions/><RetryDelaySeconds>60</RetryDelaySeconds><MaximumRunningTimeSeconds>300</MaximumRunningTimeSeconds></WindowsWorkflowConfiguration></Subscription>", EWSRunAsName);
                    }

                    //save it
                    scsmLFXConfig.AcceptChanges();
                }
            }
            catch
            {
                //if we couldn't find the rule, it must not exist. define and create it
                if (this.RunAsAccountEWS != null)
                {
                    ManagementPack scsmLFXConfig = emg.ManagementPacks.GetManagementPack(new Guid("50daaf82-06ce-cacb-8cf5-3950aebae0b0"));
                    ManagementPack msftSCLibrary = emg.ManagementPacks.GetManagementPack(new Guid("7cfc5cc0-ae0a-da4f-5ac2-d64540141a55"));
                    ManagementPack scsmSubscriptions = emg.ManagementPacks.GetManagementPack(new Guid("0306141b-bf60-70a1-be18-e979132c873c"));
                    ManagementPack scLibrary = emg.ManagementPacks.GetManagementPack(new Guid("01c8b236-3bce-9dba-6f1c-c119bcdc2972"));
                    
                    //create re-occuring schedule XML and set the interval from the value in the GUI
                    string NewSMEXCORuleDataSourceXML = string.Format("\r\n <Scheduler>\r\n <SimpleReccuringSchedule>\r\n <Interval Unit=\"Seconds\">{0}</Interval>\r\n </SimpleReccuringSchedule>\r\n <ExcludeDates />\r\n </Scheduler>", this.SMExcoIntervalSeconds);

                    //Get the Secure Reference's Management Pack's, Aliased Name, from the SCSM LFX unsealed mp
                    ManagementPack secRefMP = this.RunAsAccountEWS.GetManagementPack();
                    string mpAlias = null;
                    foreach (KeyValuePair<string, ManagementPackReference> reference in scsmLFXConfig.References)
                    {
                        if (secRefMP.Id == reference.Value.Id)
                        {
                            mpAlias = reference.Key;
                        }
                    }

                    //create the rule configuration XML and set the Run As Account references to use in the workflow
                    string NewSMEXCORuleWriteActionXML;
                    if (this.RunAsAccountEWS.Name.StartsWith("SecureReference"))
                    {
                        //if the secure reference actually begins with "SecureReference" it's defined in the Linking Framework Configuration MP
                        string EWSRunAsName = this.RunAsAccountEWS.Name;
                        NewSMEXCORuleWriteActionXML = string.Format("\r\n<Subscription>\r\n <WindowsWorkflowConfiguration>\r\n <AssemblyName>SMLets.Exchange.Connector.Resources</AssemblyName>\r\n <WorkflowTypeName>SMLets.Exchange.Connector.Resources.RunScript</WorkflowTypeName>\r\n <WorkflowParameters>\r\n <WorkflowParameter Name=\"ExchangeDomain\" Type=\"string\">$RunAs[Name=\"{0}\"]/Domain$</WorkflowParameter><WorkflowParameter Name=\"ExchangeUsername\" Type=\"string\">$RunAs[Name=\"{0}\"]/UserName$</WorkflowParameter><WorkflowParameter Name=\"ExchangePassword\" Type=\"string\">$RunAs[Name=\"{0}\"]/Password$</WorkflowParameter></WorkflowParameters><RetryExceptions/><RetryDelaySeconds>60</RetryDelaySeconds><MaximumRunningTimeSeconds>300</MaximumRunningTimeSeconds></WindowsWorkflowConfiguration></Subscription>", EWSRunAsName);
                    }
                    else
                    {
                        //if it doesn't begin with SecureReference, it's defined in the Core MP which is already aliased in the Linking Framework Configuration MP
                        string EWSRunAsName = mpAlias + "!" + this.RunAsAccountEWS.Name;
                        NewSMEXCORuleWriteActionXML = string.Format("\r\n<Subscription>\r\n <WindowsWorkflowConfiguration>\r\n <AssemblyName>SMLets.Exchange.Connector.Resources</AssemblyName>\r\n <WorkflowTypeName>SMLets.Exchange.Connector.Resources.RunScript</WorkflowTypeName>\r\n <WorkflowParameters>\r\n <WorkflowParameter Name=\"ExchangeDomain\" Type=\"string\">$RunAs[Name=\"{0}\"]/Domain$</WorkflowParameter><WorkflowParameter Name=\"ExchangeUsername\" Type=\"string\">$RunAs[Name=\"{0}\"]/UserName$</WorkflowParameter><WorkflowParameter Name=\"ExchangePassword\" Type=\"string\">$RunAs[Name=\"{0}\"]/Password$</WorkflowParameter></WorkflowParameters><RetryExceptions/><RetryDelaySeconds>60</RetryDelaySeconds><MaximumRunningTimeSeconds>300</MaximumRunningTimeSeconds></WindowsWorkflowConfiguration></Subscription>", EWSRunAsName);
                    }

                    //create the new Management Pack Rule using the XML strings defined above.
                    //Since only one instance of the SMLets Exchange Connector ever be deployed, the Rule Name is statically defined
                    //To ensure its uniqueness, it's defined here as Name + the guid of the SMlets Exchange Connector MP sans hypens
                    ManagementPackRule NewSMEXCORule = new ManagementPackRule(scsmLFXConfig, "SMLets.Exchange.Connector.15d8b765a2f8b63ead14472f9b3c12f0");
                    NewSMEXCORule.Target = (ManagementPackElementReference<ManagementPackClass>)msftSCLibrary.GetClass("Microsoft.SystemCenter.SubscriptionWorkflowTarget");

                    //Is the workflow enabled?
                    if (this.IsSMExcoWorkflowEnabled == true)
                    {
                        NewSMEXCORule.Enabled = ManagementPackMonitoringLevel.@true;
                    }
                    else
                    {
                        NewSMEXCORule.Enabled = ManagementPackMonitoringLevel.@false;
                    }
                    
                    //build the Data Sources and Write Actions for the new Rule
                    ManagementPackDataSourceModule dataSource = new ManagementPackDataSourceModule((ManagementPackElement)NewSMEXCORule, "DS1");
                        dataSource.RunAs = (ManagementPackElementReference<ManagementPackSecureReference>)emg.Security.GetSecureReference(new Guid("A7ACDF53-01B7-84DF-7E10-C933F0DC9DC2"));
                        dataSource.TypeID = (ManagementPackElementReference<ManagementPackDataSourceModuleType>)scLibrary.GetModuleType("System.Scheduler");
                        dataSource.Configuration = NewSMEXCORuleDataSourceXML;
                        NewSMEXCORule.DataSourceCollection.Add(dataSource);
                    ManagementPackWriteActionModule writeAction = new ManagementPackWriteActionModule((ManagementPackElement)NewSMEXCORule, "WA1");
                        writeAction.TypeID = (ManagementPackElementReference<ManagementPackWriteActionModuleType>)scsmSubscriptions.GetModuleType("Microsoft.EnterpriseManagement.SystemCenter.Subscription.WindowsWorkflowTaskWriteAction");    
                        writeAction.Configuration = NewSMEXCORuleWriteActionXML;
                        NewSMEXCORule.WriteActionCollection.Add(writeAction);
                    NewSMEXCORule.Status = ManagementPackElementStatus.PendingAdd;
                    
                    //save it
                    scsmLFXConfig.AcceptChanges();
                }
            }

            //Update the MP
            emoAdminSetting.Commit();
            this.WizardResult = WizardResult.Success;
        }
    }
}
