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

        //paths
        private String strEWSFilePath = String.Empty;
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
        private Boolean boolEnableAnnouncementIntegration = false;
        private Boolean boolEnableSCSMAnnouncements = false;
        private Boolean boolEnableCiresonAnnouncements = false;
        private Boolean boolEnableCiresonIntegration = false;
        private Boolean boolEnableCiresonKBSearch = false;
        private Boolean boolEnableCiresonROSearch = false;
        private Boolean boolEnableCiresonFirstResponseDateOnSuggestions = false;

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

            //AML incident custom decimal extensions
            ManagementPackProperty incidentAMLWIConfidenceExtensionDec;
            ManagementPackProperty incidentAMLClassifcationConfidenceExtensionDec;
            ManagementPackProperty incidentAMLSupportGroupIConfidenceExtensionDec;
            ManagementPackProperty incidentAMLClassificationPredictionExtensionEnum;
            ManagementPackProperty incidentAMLSupportGroupPredictionExtensionEnum;

            //AML service request custom decimal extensions
            ManagementPackProperty serviceRequestAMLWIConfidenceExtensionDec;
            ManagementPackProperty serviceRequestAMLClassifcationConfidenceExtensionDec;
            ManagementPackProperty serviceRequestAMLSupportGroupIConfidenceExtensionDec;
            ManagementPackProperty serviceRequestAMLClassificationPredictionExtensionEnum;
            ManagementPackProperty serviceRequestAMLSupportGroupPredictionExtensionEnum;

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
        public IList<ManagementPackProperty> ServiceRequestDecExtensions { get; set; }
        public IList<ManagementPackProperty> ServiceRequestEnumExtensions { get; set; }
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

            //Autodiscover
            try { this.IsAutodiscoverEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "UseAutoDiscover"].ToString()); }
            catch { this.IsAutodiscoverEnabled = false; }

            //DLL Paths - EWS, Mimekit, PII regex, HTML Suyggestions, Custom Events
            this.EWSFilePath = emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathEWSDLL"].ToString();
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
                if ((crproperty.Type.ToString() == "enum") && (!(crproperty.Identifier.ToString().Contains("System.WorkItem.ChangeRequest"))))
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
                if ((prproperty.Type.ToString() == "enum") && (!(prproperty.Identifier.ToString().Contains("System.WorkItem.Problem"))) && (!(prproperty.Identifier.ToString().Contains("Trouble"))))
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
            if (this.MinFileAttachmentSize != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinimumFileAttachmentSize"].ToString(); }
            else { this.MinFileAttachmentSize = strMinFileSize; }

            try { this.IsMaxFileSizeAttachmentsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnforceFileAttachmentSettings"].ToString()); }
            catch { this.IsMaxFileSizeAttachmentsEnabled = false; }

            //Templates
            this.DefaultWorkItem = emoAdminSetting[smletsExchangeConnectorSettingsClass, "DefaultWorkItemType"].ToString();

                /*get the Incident Templates by Type Projection ID/GUID
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
                */
                ManagementPackObjectTemplateCriteria mpotcIncidents = new ManagementPackObjectTemplateCriteria("TypeID = '285cb0a2-f276-bccb-563e-bb721df7cdec'");
                this.IncidentTemplates = emg.Templates.GetObjectTemplates(mpotcIncidents);
                this.IncidentTemplates = this.IncidentTemplates.OrderBy(template => template.DisplayName).ToList();
                try {
                    Guid irTemplateGuid = (Guid)emoAdminSetting[null, "DefaultIncidentTemplateGUID"].Value;
                    this.DefaultIncidentTemplate = emg.Templates.GetObjectTemplate(irTemplateGuid);
                }
                catch { }

                //get the Service Request Templates by Type Projection ID/GUID
                ManagementPackObjectTemplateCriteria mpotcServiceRequests = new ManagementPackObjectTemplateCriteria("TypeID = 'e44b7c06-590d-64d6-56d2-2219c5e763e0'");
                this.ServiceRequestTemplates = emg.Templates.GetObjectTemplates(mpotcServiceRequests);
                this.ServiceRequestTemplates = this.ServiceRequestTemplates.OrderBy(template => template.DisplayName).ToList();
                try
                {
                    Guid srTemplateGuid = (Guid)emoAdminSetting[null, "DefaultServiceRequestTemplateGUID"].Value;
                    this.DefaultServiceRequestTemplate = emg.Templates.GetObjectTemplate(srTemplateGuid);
                }
                catch { }

                //get the Change Request Templates by Type Projection ID/GUID
                ManagementPackObjectTemplateCriteria mpotcChangeRequests = new ManagementPackObjectTemplateCriteria("TypeID = '674194d8-0246-7b90-d871-e1ea015b2ea7'");
                this.ChangeRequestTemplates = emg.Templates.GetObjectTemplates(mpotcChangeRequests);
                this.ChangeRequestTemplates = this.ChangeRequestTemplates.OrderBy(template => template.DisplayName).ToList();
                try
                {
                    Guid crTemplateGuid = (Guid)emoAdminSetting[null, "DefaultChangeRequestTemplateGUID"].Value;
                    this.DefaultChangeRequestTemplate = emg.Templates.GetObjectTemplate(crTemplateGuid);
                }
                catch { }

                //get the Problem Templates by Type Projection ID/GUID
                ManagementPackObjectTemplateCriteria mpotcProblems = new ManagementPackObjectTemplateCriteria("TypeID = '45c1c404-f3fe-1050-dcef-530e1c2533e1'");
                this.ProblemTemplates = emg.Templates.GetObjectTemplates(mpotcProblems);
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
            try { this.IsCiresonIntegrationEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableCiresonIntegration"].ToString()); }
            catch { this.IsCiresonIntegrationEnabled = false; }
            try { this.IsCiresonFirstResponseDateOnSuggestionsEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSetFirstResponseDateOnSuggestions"].ToString()); }
            catch { this.IsCiresonFirstResponseDateOnSuggestionsEnabled = false; }
            try { this.IsCiresonKBSearchEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchKnowledgeBase"].ToString()); }
            catch { this.IsCiresonKBSearchEnabled = false; }
            try { this.IsCiresonROSearchEnabled = Boolean.Parse(emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchRequestOfferings"].ToString()); }
            catch { this.IsCiresonROSearchEnabled = false; }

            //Announcements
            this.AuthorizedAnnouncementApproverType = emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMAnnouncementApprovedMemberType"].ToString();
            this.LowPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordLow"].ToString();
            this.MediumPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordMedium"].ToString();
            this.HighPriorityAnnouncementKeyword = emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordHigh"].ToString();

            if (this.LowPriorityExpiresInHours != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityLowExpirationInHours"].ToString(); }
            else { this.LowPriorityExpiresInHours = strLowPriorityExpInHours; }

            if (this.MediumPriorityExpiresInHours != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityNormalExpirationInHours"].ToString(); }
            else { this.MediumPriorityExpiresInHours = strMedPriorityExpInHours; }

            if (this.HighPriorityExpiresInHours != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityCriticalExpirationInHours"].ToString(); }
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
            if (this.MinimumPercentToCreateServiceRequest != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinACSSentimentToCreateSR"].ToString(); }
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
            if (this.AzureMachineLearningWIConfidence != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemType"].ToString(); }
            else { this.AzureMachineLearningWIConfidence = decAMLWIMinConfidence; }
            if (this.AzureMachineLearningClassificationConfidence != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemClassification"].ToString(); }
            else { this.AzureMachineLearningClassificationConfidence = decAMLClassificationMinConfidence; }
            if (this.AzureMachineLearningSupportGroupConfidence != null) { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemSupportGroup"].ToString(); }
            else { this.AzureMachineLearningSupportGroupConfidence = decAMLSupportGroupMinConfidence; }

            ManagementPackClass incidentClass = emg.EntityTypes.GetClass(new Guid("a604b942-4c7b-2fb2-28dc-61dc6f465c68"));
            ManagementPackClass serviceRequestClass = emg.EntityTypes.GetClass(new Guid("04b69835-6343-4de2-4b19-6be08c612989"));
            IList<ManagementPackProperty> irPropertyList = incidentClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            IList<ManagementPackProperty> srPropertyList = serviceRequestClass.GetProperties(BaseClassTraversalDepth.Recursive, PropertyExtensionMode.All);
            List<ManagementPackProperty> irTempPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> srTempPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> irTempEnumPropertyList = new List<ManagementPackProperty>();
            List<ManagementPackProperty> srTempEnumPropertyList = new List<ManagementPackProperty>();
            
            //Load the Class Extension lists into temporary lists as long as the Property type is an dec/enum and does not come from the stock Class
            //Load the Drop Down's currently Selected Item in the list if the stored GUID matches a property's respective id
            foreach (ManagementPackProperty irproperty in irPropertyList)
            {
                //load all of the decimal properties that aren't from the Microsoft MP into a list
                if ((irproperty.Type.ToString() == "decimal") && (!(irproperty.Identifier.ToString().Contains("System.WorkItem.Incident"))))
                {
                    irTempPropertyList.Add(irproperty);
                }

                //of the decimal items being loaded, if they are the ones currently written to the MP set them as the selected item in the drop down
                if ((irproperty.Type.ToString() == "decimal") && (!(irproperty.Identifier.ToString().Contains("System.WorkItem.Incident"))))
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
                if ((irproperty.Type.ToString() == "enum") && (!(irproperty.Identifier.ToString().Contains("System.WorkItem.Incident"))) && (!(irproperty.Identifier.ToString().Contains("Trouble"))))
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
            }

            //Load the Class Extension lists into temporary lists as long as the Property type is an dec/enum and does not come from the stock Class
            //Load the Drop Down's currently Selected Item in the list if the stored GUID matches a property's respective id
            foreach (ManagementPackProperty srproperty in srPropertyList)
            {                
                if ((srproperty.Type.ToString() == "decimal") && (!(srproperty.Identifier.ToString().Contains("System.WorkItem.ServiceRequest"))))
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
                if ((srproperty.Type.ToString() == "enum") && (!(srproperty.Identifier.ToString().Contains("System.WorkItem.ServiceRequest"))))
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
            }
            //Processing Logic - load the Class Extension Lists from the temporary lists
            this.IncidentDecExtensions = irTempPropertyList.ToList();
            this.IncidentDecExtensions = irTempPropertyList.OrderBy(irextensions => irextensions.DisplayName).ToList();
            this.IncidentEnumExtensions = irTempEnumPropertyList.ToList();
            this.IncidentEnumExtensions = irTempEnumPropertyList.OrderBy(irextensions => irextensions.DisplayName).ToList();
            this.ServiceRequestDecExtensions = srTempPropertyList.ToList();
            this.ServiceRequestDecExtensions = srTempPropertyList.OrderBy(srextensions => srextensions.DisplayName).ToList();
            this.ServiceRequestEnumExtensions = srTempEnumPropertyList.ToList();
            this.ServiceRequestEnumExtensions = srTempEnumPropertyList.OrderBy(srextensions => srextensions.DisplayName).ToList();

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
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CreateUsersNotInCMDB"].Value = this.CreateUsersNotFoundtInCMDB;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "IncludeWholeEmail"].Value = this.IncludeWholeEmail;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AttachEmailToWorkItem"].Value = this.AttachEmailToWorkItem;

            //DLL Paths - EWS, Mimekit, PII regex, HTML Suyggestions, Custom Events
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "FilePathEWSDLL"].Value = this.EWSFilePath;
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
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchKnowledgeBase"].Value = this.IsCiresonKBSearchEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonSearchRequestOfferings"].Value = this.IsCiresonROSearchEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "CiresonPortalURL"].Value = this.CiresonPortalURL;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "NumberOfWordsToMatchFromEmailToCiresonRequestOffering"].Value = this.MinWordCountToSuggestRO;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "NumberOfWordsToMatchFromEmailToCiresonKnowledgeArticle"].Value = this.MinWordCountToSuggestKA;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSetFirstResponseDateOnSuggestions"].Value = this.IsCiresonFirstResponseDateOnSuggestionsEnabled;

            //Announcements
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableAnnouncements"].Value = this.IsAnnouncementIntegrationEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableSCSMAnnouncements"].Value = this.IsSCSMAnnouncementsEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "SCSMAnnouncementApprovedMemberType"].Value = this.AuthorizedAnnouncementApproverType;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "EnableCiresonSCSMAnnouncements"].Value = this.IsCiresonAnnouncementsEnabled;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordLow"].Value = this.LowPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordMedium"].Value = this.MediumPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementKeywordHigh"].Value = this.HighPriorityAnnouncementKeyword;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityLowExpirationInHours"].Value = this.LowPriorityExpiresInHours;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityNormalExpirationInHours"].Value = this.MediumPriorityExpiresInHours;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AnnouncementPriorityCriticalExpirationInHours"].Value = this.HighPriorityExpiresInHours;
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
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "MinACSSentimentToCreateSR"].Value = this.MinimumPercentToCreateServiceRequest;
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
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemType"].Value = this.AzureMachineLearningWIConfidence;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemClassification"].Value = this.AzureMachineLearningClassificationConfidence;
            emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLMinConfidenceWorkItemSupportGroup"].Value = this.AzureMachineLearningSupportGroupConfidence;

            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLIncidentConfidenceClassExtensionGUID"].Value = this.AMLIncidentConfidenceDecExtension.Id; }
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
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestClassificationConfidenceClassExtensionGUID"].Value = this.AMLServiceRequestClassificationConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestSupportGroupConfidenceClassExtensionGUID"].Value = this.AMLServiceRequestSupportGroupConfidenceDecExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestClassificationPredictionClassExtensionGUID"].Value = this.AMLServiceRequestClassificationPredictionEnumExtension.Id; }
            catch { }
            try { emoAdminSetting[smletsExchangeConnectorSettingsClass, "AMLServiceRequestSupportGroupPredictionClassExtensionGUID"].Value = this.AMLServiceRequestSupportGroupPredictionEnumExtension.Id; }
            catch { }

            //Update the MP
            emoAdminSetting.Commit();
            this.WizardResult = WizardResult.Success;
        }
    }
}