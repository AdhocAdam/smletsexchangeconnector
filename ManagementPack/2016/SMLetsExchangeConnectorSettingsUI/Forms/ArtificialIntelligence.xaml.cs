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
using System.Xml.Serialization;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;
using Microsoft.EnterpriseManagement;

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for AzureCognitiveServices.xaml
    /// </summary>
    public partial class ArtificialIntelligence : WizardRegularPageBase
    {
        //1. Create a blank Instance of AdminSettingsWizardData called "settings"
        AdminSettingWizardData settings;

        private ArtificialIntelligence azureCognitiveServicesWizardData = null;

        public ArtificialIntelligence(WizardData wizardData)
        {
            //2. Turn the blank instance of AdminSettingsWizardData into the one being used when the form was loaded
            settings = wizardData as AdminSettingWizardData;

            InitializeComponent();
            this.DataContext = wizardData;
            this.azureCognitiveServicesWizardData = this.DataContext as ArtificialIntelligence;

            //This is less than ideal as it means two connections are made to SCSM when loading the form, this was done so the returned GUID from
            //the Additional Mailbox could be looked up and then set as the Selected Item within the drop down
            //Get the server name to connect to and then connect 
            String strServerName = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Microsoft\\System Center\\2010\\Service Manager\\Console\\User Settings", "SDKServiceMachine", "localhost").ToString();
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup(strServerName);

            //enable Cireson specific controls if these Cireson features are checked off on the Cireson integration form
            if (settings.IsCiresonKBSearchEnabled == true)
            { 
                chkImproveCiresonKASuggestion.IsEnabled = true;
                chkImproveCiresonKASuggestionAML.IsEnabled = true; 
            }
            else
            {
                chkImproveCiresonKASuggestion.IsEnabled = false;
                chkImproveCiresonKASuggestionAML.IsEnabled = false; 
            }
            if (settings.IsCiresonROSearchEnabled == true)
            {
                chkImproveCiresonROSuggestion.IsEnabled = true;
                chkImproveCiresonROSuggestionAML.IsEnabled = true;
            }
            else
            {
                chkImproveCiresonROSuggestion.IsEnabled = false;
                chkImproveCiresonROSuggestionAML.IsEnabled = false;
            }

            //parse ACS Boundary XML back into the Listview
            try
            {
                string xmlBoundaries = settings.AzureCognitiveServicesBoundaries;
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xmlBoundaries);
                foreach (XmlNode node in doc.DocumentElement)
                {
                    if (node.ChildNodes[0].InnerText=="IR")
                    {
                        ACSBoundaryListViewControl.Items.Add(new ACSPriorityBoundary()
                        {
                            WorkItemType = node.ChildNodes[0].InnerText,
                            MinBoundary = node.ChildNodes[1].InnerText,
                            MaxBoundary = node.ChildNodes[2].InnerText,
                            IRImpactSRUrgency = emg.EntityTypes.GetEnumeration(new Guid(node.ChildNodes[3].InnerText)),
                            IRImpactSRUrgencyEnums = settings.IncidentImpactEnums,
                            IRUrgencySRPriority = emg.EntityTypes.GetEnumeration(new Guid(node.ChildNodes[4].InnerText)),
                            IRUrgencySRPriorityEnums = settings.IncidentUrgencyEnums
                        });
                    }
                    else
                    {
                        ACSBoundaryListViewControl.Items.Add(new ACSPriorityBoundary()
                        {
                            WorkItemType = node.ChildNodes[0].InnerText,
                            MinBoundary = node.ChildNodes[1].InnerText,
                            MaxBoundary = node.ChildNodes[2].InnerText,
                            IRImpactSRUrgency = emg.EntityTypes.GetEnumeration(new Guid(node.ChildNodes[3].InnerText)),
                            IRImpactSRUrgencyEnums = settings.ServiceRequestUrgencyEnums,
                            IRUrgencySRPriority = emg.EntityTypes.GetEnumeration(new Guid(node.ChildNodes[4].InnerText)),
                            IRUrgencySRPriorityEnums = settings.ServiceRequestPriorityEnums
                        });
                    }
                }
            }
            catch { }
        }

        /*
        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
        */

        //take the URL defined in the WPF and open a browser to it
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.AbsoluteUri);
        }

        //create a class so additional acs boundaries can be created per row and later worked with as xml objects
        public class ACSPriorityBoundary
        {
            public String WorkItemType { get; set; }
            public String MinBoundary { get; set; }
            public String MaxBoundary { get; set; }
            public IList<ManagementPackEnumeration> IRImpactSRUrgencyEnums { get; set; }
            public IList<ManagementPackEnumeration> IRUrgencySRPriorityEnums { get; set; }
            ManagementPackEnumeration irImpactSRUrgency;
            public ManagementPackEnumeration IRImpactSRUrgency
            {
                get
                {
                    return irImpactSRUrgency;
                }
                set
                {
                    if (this.irImpactSRUrgency != value)
                    {
                        irImpactSRUrgency = value;
                    }
                }
            }
            ManagementPackEnumeration irUrgencySRPriority;
            public ManagementPackEnumeration IRUrgencySRPriority
            {
                get
                {
                    return irUrgencySRPriority;
                }
                set
                {
                    if (this.irUrgencySRPriority != value)
                    {
                        irUrgencySRPriority = value;
                    }
                }
            }
        }

        //Click Event - Add an Incident Row to the ACS Boundary Listview. Instantiate from the defined Class above
        private void AddIRACSBoundary_Click(object sender, RoutedEventArgs e)
        {
            //Now when creating ACS boundary entries, the Template lists are populated from the lists that have already been created via the settings variable
            ACSBoundaryListViewControl.Items.Add(new ACSPriorityBoundary()
            {
                WorkItemType = "IR",
                MinBoundary = "0",
                MaxBoundary = "0",
                IRImpactSRUrgencyEnums = settings.IncidentImpactEnums,
                IRUrgencySRPriorityEnums = settings.IncidentUrgencyEnums
            });
        }

        //Click Event - Add a Service Request Row to the ACS Boundary Listview. Instantiate from the defined Class above
        private void AddSRACSBoundary_Click(object sender, RoutedEventArgs e)
        {
            //Now when creating ACS boundary entries, the Template lists are populated from the lists that have already been created via the settings variable
            ACSBoundaryListViewControl.Items.Add(new ACSPriorityBoundary()
            {
                WorkItemType = "SR",
                MinBoundary = "0",
                MaxBoundary = "0",
                IRImpactSRUrgencyEnums = settings.ServiceRequestUrgencyEnums,
                IRUrgencySRPriorityEnums = settings.ServiceRequestPriorityEnums
            });
        }

        //Click Event - Remove the selected Row from the ACS Boundary Listview
        private void RemoveACSBoundary_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //convert the selected item to an ACSPriorityBoundary and remove it
                var acsBoundaryToRemove = ACSBoundaryListViewControl.SelectedItem as ACSPriorityBoundary;
                ACSBoundaryListViewControl.Items.RemoveAt(ACSBoundaryListViewControl.Items.IndexOf(acsBoundaryToRemove));
            }
            catch { }
        }

        //when the form is saved, these values/objects will be passed back over to AdminSettingsWizardData's "AcceptChanges" event
        public override void SaveState(SavePageEventArgs e)
        {
            int totalACSBoundariesToAdd = ACSBoundaryListViewControl.Items.Count;
            string[] acsBoundariesToAddStringArray = new string[totalACSBoundariesToAdd];

            int x = 0;
            foreach (ACSPriorityBoundary acsBoundary in ACSBoundaryListViewControl.Items)
            {
                acsBoundariesToAddStringArray[x] +=
                    "<ACSPriorityBoundary>" +
                        "<WorkItemType>" + acsBoundary.WorkItemType + "</WorkItemType>" +
                        "<Min>" + acsBoundary.MinBoundary + "</Min>" +
                        "<Max>" + acsBoundary.MaxBoundary + "</Max>" +
                        "<IRImpactSRUrgencyEnum>" + acsBoundary.IRImpactSRUrgency.Id + "</IRImpactSRUrgencyEnum>" +
                        "<IRUrgencySRPriorityEnum>" + acsBoundary.IRUrgencySRPriority.Id + "</IRUrgencySRPriorityEnum>" +
                    "</ACSPriorityBoundary>";
                x++;
            }

            //prepare the XML string
            string ACSBoundariesXMLString = "<ACSPriorityBoundaries>" + string.Join("", acsBoundariesToAddStringArray) + "</ACSPriorityBoundaries>";

            //Send back over to AdminSettingsWizardData for Save
            if (acsBoundariesToAddStringArray != null) { settings.AzureCognitiveServicesBoundaries = ACSBoundariesXMLString; }
        }
    }
}
