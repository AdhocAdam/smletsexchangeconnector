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

namespace SMLetsExchangeConnectorSettingsUI
{
    /// <summary>
    /// Interaction logic for TemplateForm.xaml
    /// </summary>

    public partial class DLLForm : WizardRegularPageBase
    {
        private DLLForm adminSettingWizardData = null;

        public DLLForm(WizardData wizardData)
        {
            InitializeComponent();
            this.DataContext = wizardData;
            this.adminSettingWizardData = this.DataContext as DLLForm;
        }

        private void WizardRegularPageBase_Loaded(object sender, RoutedEventArgs e)
        {

        }

        //browse buttons to set the file paths
        //since a mix of WPF dialogs and WinForm dialogs are used, they are explicity referenced
        //the commented Dialog shows the original shorthand which can't be used since it creates
        //ambiguity between the classes given their identical naming conventions
        private void btn_BrowseEWSPath(object sender, RoutedEventArgs e)
        {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "DLL files (*.dll)|*.dll|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtEWSapiDll.Text = openFileDialog.FileName;
            }
        }

        private void btn_BrowseMSALPath(object sender, RoutedEventArgs e)
        {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "DLL files (*.dll)|*.dll|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtMSALDLL.Text = openFileDialog.FileName;
            }
        }

        private void btn_BrowseMimeKitPath(object sender, RoutedEventArgs e)
        {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "DLL files (*.dll)|*.dll|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtMimeKitDLL.Text = openFileDialog.FileName;
            }
        }

        private void btn_BrowsePIIRegexPath(object sender, RoutedEventArgs e)
        {
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == true)
            {
                txtPIIRegex.Text = openFileDialog.FileName;
            }
        }

        private void btn_BrowsePIIHTMLTemplatePath(object sender, RoutedEventArgs e)
        {
            var openFolderDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (openFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtHTMLSuggestionTemplates.Text = openFolderDialog.SelectedPath + "\\";
            }
        }
    }
}
