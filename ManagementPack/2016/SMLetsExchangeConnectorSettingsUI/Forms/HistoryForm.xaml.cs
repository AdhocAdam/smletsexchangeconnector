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
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}