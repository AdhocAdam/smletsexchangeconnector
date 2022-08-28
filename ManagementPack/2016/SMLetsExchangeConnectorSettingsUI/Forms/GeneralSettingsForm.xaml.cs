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

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for SettingsUI.xaml
    /// </summary>
    public partial class GeneralSettingsForm : WizardRegularPageBase
    {
        private GeneralSettingsForm adminSettingWizardData = null;

        public GeneralSettingsForm(WizardData wizardData)
        {
            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as GeneralSettingsForm;
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

        //toggle AutoDiscover URL textbox availability
        private void chkUseAutoDiscover_Checked(object sender, RoutedEventArgs e)
        {
            this.txtAutodiscoverURL.IsEnabled = false;
        }

        private void chkUseAutoDiscover_Unchecked(object sender, RoutedEventArgs e)
        {
            this.txtAutodiscoverURL.IsEnabled = true;
        }
    }
}
