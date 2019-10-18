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
    /// Interaction logic for ProcessingLogicForm.xaml
    /// </summary>

    public partial class ProcessingLogicForm : WizardRegularPageBase
    {
        //1. Create a blank Instance of AdminSettingsWizardData called "settings"
        AdminSettingWizardData settings;

        private ProcessingLogicForm adminSettingWizardData = null;

        public ProcessingLogicForm(WizardData wizardData)
        {
            //2. Turn the blank instance of AdminSettingsWizardData into the one being used when the form was loaded
            settings = wizardData as AdminSettingWizardData;

            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as ProcessingLogicForm;

            //enable Cireson specific controls if the MP is imported per the AdminSettingsWizardData
            if (settings.IsCiresonPortalMPImported == false) { chkTakeRequiresCiresonGroupMembership.IsEnabled = false; }

            //set default values within the statically defined drop downs within the WPF form
            if (settings.ExternalIRCommentType == "(null)") { CBExternalIRCommentType.SelectedIndex = 1; }
            if (settings.ExternalSRCommentType == "(null)") { CBExternalSRCommentType.SelectedIndex = 1; }
            if (settings.ExternalIRCommentPrivate == "(null)") { CBExternalIRCommentTypePrivate.SelectedIndex = 0; }
            if (settings.ExternalSRCommentPrivate == "(null)") { CBExternalSRCommentTypePrivate.SelectedIndex = 0; }
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }

        //take the URL defined in the WPF and open a browser to it
        private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.AbsoluteUri);
        }

        //when the form is saved, these values/objects will be passed back over to AdminSettingsWizardData's "AcceptChanges" event
        public override void SaveState(SavePageEventArgs e)
        {
            settings.ExternalIRCommentType = CBExternalIRCommentType.SelectedItem.ToString();
            settings.ExternalSRCommentType = CBExternalSRCommentType.SelectedItem.ToString();
            settings.ExternalIRCommentPrivate = CBExternalIRCommentTypePrivate.SelectedItem.ToString();
            settings.ExternalSRCommentPrivate = CBExternalSRCommentTypePrivate.SelectedItem.ToString();

            //ensure that there wasn't an attempt intentional/accidental to control an End User Comment's privacy as it doesn't exist
            if (CBExternalIRCommentType.SelectedItem.ToString() == "EndUserComment") { CBExternalIRCommentTypePrivate.SelectedItem = "null"; }
            if (CBExternalSRCommentType.SelectedItem.ToString() == "EndUserComment") { CBExternalSRCommentTypePrivate.SelectedItem = "null"; }
        }
        
        //delete key clears depending on what dropdown list you are working with
        private void cbDynamicAssignment_KeyDown(object sender, KeyEventArgs e)
        {
            cbDynamicAssignment.SelectedIndex = -1;
        }

        private void cbIRResCat_KeyDown(object sender, KeyEventArgs e)
        {
            cbIRResCat.SelectedIndex = -1;
        }

        private void crSupportGroup_KeyDown(object sender, KeyEventArgs e)
        {
            crSupportGroup.SelectedIndex = -1;
        }

        private void maSupportGroup_KeyDown(object sender, KeyEventArgs e)
        {
            maSupportGroup.SelectedIndex = -1;
        }

        private void prSupportGroup_KeyDown(object sender, KeyEventArgs e)
        {
            prSupportGroup.SelectedIndex = -1;
        }

        private void cbSRImplementCat_KeyDown(object sender, KeyEventArgs e)
        {
            cbSRImplementCat.SelectedIndex = -1;
        }

        private void cbProblemResCat_KeyDown(object sender, KeyEventArgs e)
        {
            cbProblemResCat.SelectedIndex = -1;
        }

        private void chkChangeIRStatusOnReply_Checked(object sender, RoutedEventArgs e)
        {
            cbIRStatusAUReply.SelectedIndex = 0;
            cbIRStatusATReply.SelectedIndex = 0;
            cbIRStatusRELReply.SelectedIndex = 0;
        }
    }
}
