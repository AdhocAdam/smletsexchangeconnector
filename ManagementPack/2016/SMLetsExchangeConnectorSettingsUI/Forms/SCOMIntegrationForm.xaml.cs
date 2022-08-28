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
    /// Interaction logic for SCOMIntegrationForm.xaml
    /// </summary>

    public partial class SCOMIntegrationForm : WizardRegularPageBase
    {
        private SCOMIntegrationForm adminSettingWizardData = null;

        public SCOMIntegrationForm(WizardData wizardData)
        {
            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as SCOMIntegrationForm;
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
