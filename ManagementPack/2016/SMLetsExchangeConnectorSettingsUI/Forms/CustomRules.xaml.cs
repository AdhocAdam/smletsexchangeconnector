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
using Microsoft.Win32;
using System.IO;
using System.Windows.Forms;
using Microsoft.EnterpriseManagement.Configuration;

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for CustomRules.xaml
    /// </summary>
    public partial class CustomRules : WizardRegularPageBase
    {
        //1. Create a blank Instance of AdminSettingsWizardData called "settings"
        AdminSettingWizardData settings;

        //create list to store custom rule GUIDs that will be removed
        public List<Object> CustomRulesToRemove = new List<Object>();

        private CustomRules adminSettingWizardData = null;

        public CustomRules(WizardData wizardData)
        {
            //2. Turn the blank instance of AdminSettingsWizardData into the one being used when the form was loaded
            settings = wizardData as AdminSettingWizardData;

            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as CustomRules;

            //load any custom rules
            try
            {
                foreach (var rule in settings.CustomRules)
                {
                    CustomRulesListViewControl.Items.Add(new AdditionalCustomRule()
                    {
                        CustomRuleID = rule.Id,
                        CustomRuleDisplayName = (String)rule.Values[1].ToString(),
                        CustomRuleMessageClassEnum = (ManagementPackEnumeration)rule.Values[2].Value,
                        CustomRuleMessageClasses = settings.MessageClassEnums,
                        CustomRuleMessagePart = (String)rule.Values[3].ToString(),
                        CustomRuleRegex = (String)rule.Values[4].ToString(),
                        CustomRuleItemType = (String)rule.Values[5].ToString(),
                        CustomRuleRegexMatchProperty = (String)rule.Values[6].ToString()
                    });
                }
            }
            catch { }
        }

        /*
        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
        */

        //create a class so additional custom rule object can be created per row and later worked with as objects
        public class AdditionalCustomRule
        {
            public Guid CustomRuleID { get; set; }
            public String CustomRuleDisplayName { get; set; }
            public ManagementPackEnumeration CustomRuleMessageClassEnum { get; set; }
            public IList<ManagementPackEnumeration> CustomRuleMessageClasses { get; set; }
            public String CustomRuleMessagePart { get; set; }
            public String CustomRuleRegex { get; set; }
            public String CustomRuleItemType { get; set; }
            public String CustomRuleRegexMatchProperty { get; set; }
        }

        //take the URL defined in the WPF and open a browser to it
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.AbsoluteUri);
        }

        //Click Event - Add an Empty Row to the Custom Rules Listview. Instantiate from the defined Class above
        private void AddButton_Click(object sender, RoutedEventArgs e)
        { 
            //Now when creating custom rule entries, the Template lists are populated from the lists that have already been created via the settings variable
            CustomRulesListViewControl.Items.Add(new AdditionalCustomRule()
            {
                CustomRuleID = Guid.NewGuid(),
                CustomRuleDisplayName = String.Empty,
                CustomRuleMessageClasses = settings.MessageClassEnums,
                CustomRuleMessagePart = "Subject",
                CustomRuleRegex = String.Empty,
                CustomRuleItemType = "IR",
                CustomRuleRegexMatchProperty = ""
            });
        }

        //Click Event - Remove the selected Row from the Custom Rules Listview and add the GUID to the mailboxes to remove array
        private void RemoveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //convert the selected item to an CustomRule and remove it
                var customRuleToRemove = CustomRulesListViewControl.SelectedItem as AdditionalCustomRule;
                CustomRulesListViewControl.Items.RemoveAt(CustomRulesListViewControl.Items.IndexOf(customRuleToRemove));
                CustomRulesToRemove.Add(customRuleToRemove);
            }
            catch { }
        }

        //when the form is saved, these values/objects will be passed back over to AdminSettingsWizardData's "AcceptChanges" event
        public override void SaveState(SavePageEventArgs e)
        {
            int totalCustomRulesToAdd = CustomRulesListViewControl.Items.Count;
            string[] customRuleToAddStringArray = new string[totalCustomRulesToAdd];
            int totalCustomRulesToRemove = CustomRulesToRemove.Count;
            string[] customRuleToRemoveStringArray = new string[totalCustomRulesToRemove];

            int x = 0;
            foreach (AdditionalCustomRule customRuleEntry in CustomRulesListViewControl.Items)
            {
                customRuleToAddStringArray[x] +=
                    customRuleEntry.CustomRuleID + ";" +                    //1 - (guid) RuleID
                    customRuleEntry.CustomRuleDisplayName + ";" +           //2 - (string) CustomRuleDisplayName
                    customRuleEntry.CustomRuleMessageClassEnum.Id + ";" +   //3 - (enum) CustomRuleMessageClass
                    customRuleEntry.CustomRuleRegex + ";" +                 //4 - (string) CustomRuleRegexPattern
                    customRuleEntry.CustomRuleMessagePart + ";" +           //5 - (string) CustomRuleMessagePart
                    customRuleEntry.CustomRuleItemType + ";" +              //6 - (string) CustomRuleWorkItemToCreate
                    customRuleEntry.CustomRuleRegexMatchProperty;           //7 - (string) CustomRuleRegexMatchProperty
                x++;
            }

            int y = 0;
            foreach (AdditionalCustomRule customRuleToRemove in CustomRulesToRemove)
            {
                customRuleToRemoveStringArray[y] +=
                    customRuleToRemove.CustomRuleID;
                y++;
            }

            //Send back over to AdminSettingsWizardData for Save
            if (customRuleToRemoveStringArray != null) { settings.MultipleCustomRulesToRemove = customRuleToRemoveStringArray; }
            if (customRuleToAddStringArray != null) { settings.MultipleCustomRulesToAdd = customRuleToAddStringArray; }
        }
    }
}
