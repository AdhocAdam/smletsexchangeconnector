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
    /// Interaction logic for AzureTranslate.xaml
    /// </summary>
    public partial class AzureTranslate : WizardRegularPageBase
    {
        AdminSettingWizardData settings;

        private AzureTranslate adminSettingWizardData = null;

        public AzureTranslate(WizardData wizardData)
        {
            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as AzureTranslate;
        }

        /*
        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
        */
    }
}
