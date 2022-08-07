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
using Microsoft.EnterpriseManagement.UI.WpfControls;
using Microsoft.EnterpriseManagement.UI.WpfWizardFramework;
using Microsoft.EnterpriseManagement.UI.DataModel;
using Microsoft.EnterpriseManagement.UI;
using Microsoft.EnterpriseManagement.UI.Controls;
using Microsoft.EnterpriseManagement.ServiceManager.Application.Common;
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement.Configuration;
using Microsoft.EnterpriseManagement.UI.FormsInfra;
using Microsoft.Win32;
using System.Collections.ObjectModel;

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for HistoryForm.xaml
    /// </summary>
    public partial class HistoryForm : WizardRegularPageBase
    {
        private HistoryForm adminSettingWizardData = null;
        AdminSettingWizardData settings;

        public HistoryForm(WizardData wizardData)
        {
            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as HistoryForm;

            //retrieve smlets exchange connector settings
            settings = wizardData as AdminSettingWizardData;

            //build a custom list based on SMexco configuration history
            //Based on SMlets Get-SCSMObjectHistory - https://github.com/SMLets/SMLets/blob/795242d4a44cf8857957a7628cca67655e2e4252/SMLets.Shared/Code/EntityObjects.cs#L1181-L1216
            List<CustomHistoryObject> smexcoConfigHistory = new List<CustomHistoryObject>();

            foreach (EnterpriseManagementObjectHistoryTransaction historyItem in settings.SMExcoSettingsHistory)
            {
                string propertyName = null;
                string oldVal = null;
                string newVal = null;

                foreach (KeyValuePair<Guid, EnterpriseManagementObjectHistory> h in historyItem.ObjectHistory)
                {
                    foreach (EnterpriseManagementObjectClassHistory ch in h.Value.ClassHistory)
                    {
                        foreach (KeyValuePair<ManagementPackProperty, Microsoft.EnterpriseManagement.Common.Pair<EnterpriseManagementSimpleObject, EnterpriseManagementSimpleObject>> hpc in ch.PropertyChanges)
                        {
                            //Simplify properties to variables
                            propertyName = hpc.Key.Name;
                            // propertyName = hpc.Key.DisplayName;
                            try { oldVal = hpc.Value.First.Value.ToString(); } catch { oldVal = ""; }
                            try { newVal = hpc.Value.Second.Value.ToString(); } catch { newVal = ""; }

                            //Create a new custom history object
                            smexcoConfigHistory.Add(new CustomHistoryObject()
                            {
                                UserName = historyItem.UserName,
                                Property = propertyName,
                                OldValue = oldVal,
                                NewValue = newVal,
                                DateOccured = historyItem.DateOccurred.ToLocalTime(),
                                DateAndUser = historyItem.DateOccurred.ToLocalTime().ToString() + "   " + historyItem.UserName
                            });
                        }
                    }
                }
            }

            //With the Custom History Objects created, update the list to all of them. Then define a grouping property for the XAML
            dgHistory.ItemsSource = smexcoConfigHistory;
            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(dgHistory.ItemsSource);
            view.SortDescriptions.Add(new System.ComponentModel.SortDescription("DateOccured", System.ComponentModel.ListSortDirection.Descending));
            PropertyGroupDescription groupDescription = new PropertyGroupDescription("DateAndUser");
            view.GroupDescriptions.Add(groupDescription);
        }

        //create a class for a custom history object to display, this is similiar to the MultipleMailboxes XAML and Code behind
        public class CustomHistoryObject
        {
            public string UserName { get; set; }
            public string Property { get; set; }
            public string OldValue { get; set; }
            public string NewValue { get; set; }
            public DateTime DateOccured { get; set; }
            public string DateAndUser { get; set; }
        }

        /*
        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
        */
    }
}
