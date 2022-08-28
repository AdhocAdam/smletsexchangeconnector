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
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Configuration;
using System.Collections;

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for Logging.xaml
    /// </summary>
    public partial class Logging : WizardRegularPageBase
    {
        //1. Create a blank Instance of AdminSettingsWizardData called "settings"
        AdminSettingWizardData settings;

        private Logging adminSettingWizardData = null;

        public Logging(WizardData wizardData)
        {
            //2. Turn the blank instance of AdminSettingsWizardData into the one being used when the form was loaded
            settings = wizardData as AdminSettingWizardData;

            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as Logging;

            //set default values within the statically defined drop downs within the WPF form
            if (settings.LoggingLevel == "(null)") { cbLoggingLevel.SelectedIndex = 0; }
            if (settings.LoggingType == "(null)") { cbLoggingType.SelectedIndex = 0; }
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
        
        //when the form is saved, these values/objects will be passed back over to AdminSettingsWizardData's "AcceptChanges" event
        public override void SaveState(SavePageEventArgs e)
        {
            settings.LoggingLevel = cbLoggingLevel.Text;
            settings.LoggingType = cbLoggingType.Text;
        }
    }
}
