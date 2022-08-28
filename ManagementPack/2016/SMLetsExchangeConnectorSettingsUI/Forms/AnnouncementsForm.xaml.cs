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
    /// Interaction logic for AnnouncementsForm.xaml
    /// </summary>

    public partial class AnnouncementsForm : WizardRegularPageBase
    {
        //1. Create a blank Instance of AdminSettingsWizardData called "settings"
        AdminSettingWizardData settings;

        private AnnouncementsForm adminSettingWizardData = null;

        public AnnouncementsForm(WizardData wizardData)
        {
            //2. Turn the blank instance of AdminSettingsWizardData into the one being used when the form was loaded
            settings = wizardData as AdminSettingWizardData;

            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as AnnouncementsForm;

            //enable Cireson specific controls if the MP is imported per the AdminSettingsWizardData
            if (settings.IsCiresonPortalMPImported == false) { chkPostAnnouncementToCiresonPortal.IsEnabled = false; }
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }
    }
}
