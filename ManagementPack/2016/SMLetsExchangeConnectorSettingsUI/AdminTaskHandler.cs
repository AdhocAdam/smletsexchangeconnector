using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Threading;
using System.Windows;
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement.ConnectorFramework;
using Microsoft.EnterpriseManagement.ConsoleFramework;
using Microsoft.EnterpriseManagement.UI.SdkDataAccess;
using Microsoft.EnterpriseManagement.UI.DataModel;
using Microsoft.EnterpriseManagement.UI.WpfWizardFramework;
using Microsoft.EnterpriseManagement.UI.SdkDataAccess.DataAdapters;
using Microsoft.EnterpriseManagement.UI.FormsInfra;
using Microsoft.EnterpriseManagement.UI.Extensions.Shared;
using Microsoft.EnterpriseManagement.GenericForm;
using Microsoft.EnterpriseManagement.Configuration;
using Microsoft.Win32;

namespace SMLetsExchangeConnectorSettingsUI
{
    internal class AdminTaskHandler : ConsoleCommand
    {
        public AdminTaskHandler()
        {

        }

        public override void ExecuteCommand(IList<NavigationModelNodeBase> nodes, NavigationModelNodeTask task, ICollection<string> parameters)
        {
            /* 
            This GUID is generated automatically when you import the Management Pack with the singleton admin setting class in it. 
            You can get this GUID by running a query like: 
            Select BaseManagedEntityID, FullName where FullName like ‘%<enter your class ID here>%’ 
            where the GUID you want is returned in the BaseManagedEntityID column in the result set 
            */
            /*In this above example I imported the MP, and then used:
                Select BaseManagedEntityID, FullName
                From BaseManagedEntity
                where FullName like 'SMLets%'
              against the ServiceManager DB to figure out the following strSingletonBaseManagedObjectID. Per 
              https://blogs.technet.microsoft.com/servicemanager/2011/05/26/getting-management-pack-elements-using-the-sdk/
              the id is always unique and never changes as its based on the MP ID, Element Names, and the key token.
             */

            String strSingletonBaseManagedObjectID = "a0022e87-75a8-65ee-4581-d923ff06a564";

            //Get the server name to connect to and connect to the server 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);
            
            //Get the Object using the GUID from above – since this is a singleton object we can get it by GUID 
            EnterpriseManagementObject emoAdminSetting = emg.EntityObjects.GetObject<EnterpriseManagementObject>(new Guid(strSingletonBaseManagedObjectID), ObjectQueryOptions.Default);

            //Create a new "wizard" (also used for property dialogs as in this case), set the title bar, create the data, and add the pages 
            WizardStory wizard = new WizardStory();
            wizard.WizardWindowTitle = "SMLets Exchange Connector Settings v3.0";
            WizardData data = new AdminSettingWizardData(emoAdminSetting);
            wizard.WizardData = data;
            wizard.AddLast(new WizardStep("General", typeof(GeneralSettingsForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("DLL", typeof(DLLForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("Processing Logic", typeof(ProcessingLogicForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("File Attachments", typeof(FileAttachmentForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("Templates", typeof(TemplateForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("Multi-Mailbox", typeof(MultipleMailboxes), wizard.WizardData));
            wizard.AddLast(new WizardStep("Parsing Keywords", typeof(KeywordsForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("Cireson", typeof(CiresonIntegrationForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("Announcements", typeof(AnnouncementsForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("SCOM", typeof(SCOMIntegrationForm), wizard.WizardData));
            wizard.AddLast(new WizardStep("A.I.", typeof(ArtificialIntelligence), wizard.WizardData));
            wizard.AddLast(new WizardStep("Translation", typeof(AzureTranslate), wizard.WizardData));
            wizard.AddLast(new WizardStep("Vision", typeof(AzureVision), wizard.WizardData));
            wizard.AddLast(new WizardStep("Transcription", typeof(AzureSpeech), wizard.WizardData));
            wizard.AddLast(new WizardStep("Workflow", typeof(WorkflowSettings), wizard.WizardData));
            wizard.AddLast(new WizardStep("Logging", typeof(Logging), wizard.WizardData));
            wizard.AddLast(new WizardStep("About", typeof(AboutForm), wizard.WizardData));

            //Show the property page 
            PropertySheetDialog wizardWindow = new PropertySheetDialog(wizard);
            wizardWindow.Width = 1100;
            wizardWindow.Height = 700;

            //Update the view when done so the new values are shown 
            bool? dialogResult = wizardWindow.ShowDialog();
            if (dialogResult.HasValue && dialogResult.Value)
            {
                RequestViewRefresh();
            }
        }

    }
}
