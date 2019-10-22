using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.EnterpriseManagement.UI.WpfWizardFramework;
using Microsoft.EnterpriseManagement.UI.DataModel;
using Microsoft.EnterpriseManagement.UI;
using Microsoft.EnterpriseManagement.Configuration;
using Microsoft.Win32;
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;


using Microsoft.EnterpriseManagement.Packaging;
using Microsoft.EnterpriseManagement.ConnectorFramework;
using System.ComponentModel;

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for TemplateForm.xaml
    /// </summary>

    public partial class MultipleMailboxes : WizardRegularPageBase
    {
        //1. Create a blank Instance of AdminSettingsWizardData called "settings"
        AdminSettingWizardData settings;
        //create list to store mailboxes GUIDs that will be removed
        public List<Object> MailboxesToRemove = new List<Object>();

        private MultipleMailboxes adminSettingWizardData = null;

        public MultipleMailboxes(WizardData wizardData)
        {
            //2. Turn the blank instance of AdminSettingsWizardData into the one being used when the form was loaded
            settings = wizardData as AdminSettingWizardData;

            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as MultipleMailboxes;

            //This is less than ideal as it means two connections are made to SCSM when loading the form, this was done so the returned GUID from
            //the Additional Mailbox could be looked up and then set as the Selected Item within the drop down
            //Get the server name to connect to and then connect 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);
            try
            {
                foreach (var mb in settings.AdditionalMailboxes)
                {
                    MailboxListViewControl.Items.Add(new AdditionalMailbox()
                    {
                        MailboxID = mb.Id,
                        MailboxAddress = (String)mb.Values[1].ToString(),
                        MailboxWorkItemType = (String)mb.Values[2].ToString(),
                        MailboxIRTemplate = emg.Templates.GetObjectTemplate(new Guid(mb.Values[4].ToString())),
                        MailboxIRTemplates = settings.IncidentTemplates,
                        MailboxSRTemplate = emg.Templates.GetObjectTemplate(new Guid(mb.Values[6].ToString())),
                        MailboxSRTemplates = settings.ServiceRequestTemplates,
                        MailboxCRTemplate = emg.Templates.GetObjectTemplate(new Guid(mb.Values[8].ToString())),
                        MailboxCRTemplates = settings.ChangeRequestTemplates,
                        MailboxPRTemplate = emg.Templates.GetObjectTemplate(new Guid(mb.Values[10].ToString())),
                        MailboxPRTemplates = settings.ProblemTemplates    
                    });
                }
            }
            catch { }
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        //take the URL defined in the WPF and open a browser to it
        private void Hyperlink_RequestNavigate(object sender,System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.AbsoluteUri);
        }

        //create a class so additional mailbox objects can be created per row and later worked with as objects
        public class AdditionalMailbox
        {
            public Guid MailboxID { get; set; }
            public String MailboxAddress { get; set; }
            public String MailboxWorkItemType { get; set; }
            public IList<ManagementPackObjectTemplate> MailboxIRTemplates { get; set; }
            public IList<ManagementPackObjectTemplate> MailboxSRTemplates { get; set; }
            public IList<ManagementPackObjectTemplate> MailboxCRTemplates { get; set; }
            public IList<ManagementPackObjectTemplate> MailboxPRTemplates { get; set; }
            ManagementPackObjectTemplate mailboxIRTemplate;
            public ManagementPackObjectTemplate MailboxIRTemplate
            {
                get
                {
                    return mailboxIRTemplate;
                }
                set
                {
                    if (this.mailboxIRTemplate != value)
                    {
                        mailboxIRTemplate = value;
                    }
                }
            }
            ManagementPackObjectTemplate mailboxSRTemplate;
            public ManagementPackObjectTemplate MailboxSRTemplate
            {
                get
                {
                    return mailboxSRTemplate;
                }
                set
                {
                    if (this.mailboxSRTemplate != value)
                    {
                        mailboxSRTemplate = value;
                    }
                }
            }
            ManagementPackObjectTemplate mailboxCRTemplate;
            public ManagementPackObjectTemplate MailboxCRTemplate
            {
                get
                {
                    return mailboxCRTemplate;
                }
                set
                {
                    if (this.mailboxCRTemplate != value)
                    {
                        mailboxCRTemplate = value;
                    }
                }
            }
            ManagementPackObjectTemplate mailboxPRTemplate;
            public ManagementPackObjectTemplate MailboxPRTemplate
            {
                get
                {
                    return mailboxPRTemplate;
                }
                set
                {
                    if (this.mailboxPRTemplate != value)
                    {
                        mailboxPRTemplate = value;
                    }
                }
            }
        }

        //Click Event - Add an Empty Row to the Multi-Mailbox Listview. Instantiate from the defined Class above
        private void AddMailbox_Click(object sender, RoutedEventArgs e)
        {
            //Now when creating mailbox entries, the Template lists are populated from the lists that have already been created via the settings variable
            MailboxListViewControl.Items.Add(new AdditionalMailbox() {
                MailboxID = Guid.NewGuid(),
                MailboxAddress = String.Empty,
                MailboxWorkItemType = "IR",
                MailboxIRTemplates = settings.IncidentTemplates,
                MailboxSRTemplates = settings.ServiceRequestTemplates,
                MailboxCRTemplates = settings.ChangeRequestTemplates,
                MailboxPRTemplates = settings.ProblemTemplates
            });
        }

        //Click Event - Remove the selected Row from the Multi-Mailbox Listview and add the GUID to the mailboxes to remove array
        private void RemoveMailbox_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //convert the selected item to an AdditionalMailbox and remove it
                var mailboxToRemove = MailboxListViewControl.SelectedItem as AdditionalMailbox;
                MailboxListViewControl.Items.RemoveAt(MailboxListViewControl.Items.IndexOf(mailboxToRemove));
                MailboxesToRemove.Add(mailboxToRemove);
            }
            catch { }
        }

        //when the form is saved, these values/objects will be passed back over to AdminSettingsWizardData's "AcceptChanges" event
        public override void SaveState(SavePageEventArgs e)
        {
            int totalMailboxesToAdd = MailboxListViewControl.Items.Count;
            string[] mailboxToAddStringArray = new string[totalMailboxesToAdd];
            int totalMailboxesToRemove = MailboxesToRemove.Count;
            string[] mailboxToRemoveStringArray = new string[totalMailboxesToRemove];

            int x = 0;
            foreach (AdditionalMailbox mailboxEntry in MailboxListViewControl.Items)
            {
                mailboxToAddStringArray[x] +=
                    mailboxEntry.MailboxID + ";" +                      //1 - (guid) MailboxID
                    mailboxEntry.MailboxAddress + ";" +                 //2 - (string) MailboxAddress
                    mailboxEntry.MailboxWorkItemType + ";" +            //3 - (string) MailboxWorkType
                    mailboxEntry.MailboxIRTemplate.DisplayName + ";" +  //4 - (string) MailboxIRDisplayName
                    mailboxEntry.MailboxIRTemplate.Id + ";" +           //5 - (guid) MailboxIRTemplateGUID
                    mailboxEntry.MailboxSRTemplate.DisplayName + ";" +  //6 - (string) MailboxSRDisplayName
                    mailboxEntry.MailboxSRTemplate.Id + ";" +           //7 - (guid) MailboxSRTemplateGUID
                    mailboxEntry.MailboxCRTemplate.DisplayName + ";" +  //8 - (string) MailboxCRDisplayName
                    mailboxEntry.MailboxCRTemplate.Id + ";" +           //9 - (guid) MailboxCRTemplateGUID
                    mailboxEntry.MailboxPRTemplate.DisplayName + ";" +  //10 - (string) MailboxPRDisplayName
                    mailboxEntry.MailboxPRTemplate.Id;                  //11 - (guid) MailboxPRTemplateGUID
                x++;
            }

            int y = 0;
            foreach (AdditionalMailbox mailboxToRemove in MailboxesToRemove)
            {
                mailboxToRemoveStringArray[y] +=
                    mailboxToRemove.MailboxID;
                y++;
            }

            //Send back over to AdminSettingsWizardData for Save
            if (mailboxToRemoveStringArray != null) { settings.MultipleMailboxesToRemove = mailboxToRemoveStringArray; }
            if (mailboxToAddStringArray != null) { settings.MultipleMailboxesToAdd = mailboxToAddStringArray; }
        }
    }
}
