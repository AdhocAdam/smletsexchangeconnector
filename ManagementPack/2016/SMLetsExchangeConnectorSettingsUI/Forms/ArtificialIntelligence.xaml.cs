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
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }

        //take the URL defined in the WPF and open a browser to it
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.AbsoluteUri);
        }

    }
}