using System;
using System.ComponentModel;
using System.Workflow.ComponentModel;
using Microsoft.EnterpriseManagement;
using Microsoft.EnterpriseManagement.Common;
using Microsoft.EnterpriseManagement.Workflow.Common;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;
using System.Text;
using System.IO;

namespace SMLets.Exchange.Connector.Resources
{
    public partial class RunScript : WorkflowActivityBase
    {
        //define the Workflow Parameters coming in from the Management Pack
        public static DependencyProperty ExchangeDomainProperty = DependencyProperty.Register("ExchangeDomain", typeof(string), typeof(RunScript));
        public static DependencyProperty ExchangeUsernameProperty = DependencyProperty.Register("ExchangeUsername", typeof(string), typeof(RunScript));
        public static DependencyProperty ExchangePasswordProperty = DependencyProperty.Register("ExchangePassword", typeof(string), typeof(RunScript));
        public static DependencyProperty CiresonPortalDomainProperty = DependencyProperty.Register("CiresonPortalDomain", typeof(string), typeof(RunScript));
        public static DependencyProperty CiresonPortalUsernameProperty = DependencyProperty.Register("CiresonPortalUsername", typeof(string), typeof(RunScript));
        public static DependencyProperty CiresonPortalPasswordProperty = DependencyProperty.Register("CiresonPortalPassword", typeof(string), typeof(RunScript));

        [DescriptionAttribute("The domain of the user")]
        [CategoryAttribute("credential")]
        [BrowsableAttribute(true)]
        [DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Visible)]
        public string ExchangeDomain
        {
            get
            {
                return ((String)(base.GetValue(RunScript.ExchangeDomainProperty)));
            }
            set
            {
                base.SetValue(RunScript.ExchangeDomainProperty, value);
            }
        }

        [DescriptionAttribute("The username of the user")]
        [CategoryAttribute("credential")]
        [BrowsableAttribute(true)]
        [DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Visible)]
        public string ExchangeUsername
        {
            get
            {
                return ((String)(base.GetValue(RunScript.ExchangeUsernameProperty)));
            }
            set
            {
                base.SetValue(RunScript.ExchangeUsernameProperty, value);
            }
        }

        [DescriptionAttribute("The password of the user")]
        [CategoryAttribute("credential")]
        [BrowsableAttribute(true)]
        [DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Visible)]
        public string ExchangePassword
        {
            get
            {
                return ((String)(base.GetValue(RunScript.ExchangePasswordProperty)));
            }
            set
            {
                base.SetValue(RunScript.ExchangePasswordProperty, value);
            }
        }

        [DescriptionAttribute("The domain of the cireson portal user")]
        [CategoryAttribute("credential")]
        [BrowsableAttribute(true)]
        [DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Visible)]
        public string CiresonPortalDomain
        {
            get
            {
                return ((String)(base.GetValue(RunScript.CiresonPortalDomainProperty)));
            }
            set
            {
                base.SetValue(RunScript.CiresonPortalDomainProperty, value);
            }
        }

        [DescriptionAttribute("The username of the cireson portal user")]
        [CategoryAttribute("credential")]
        [BrowsableAttribute(true)]
        [DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Visible)]
        public string CiresonPortalUsername
        {
            get
            {
                return ((String)(base.GetValue(RunScript.CiresonPortalUsernameProperty)));
            }
            set
            {
                base.SetValue(RunScript.CiresonPortalUsernameProperty, value);
            }
        }

        [DescriptionAttribute("The password of the cireson portal user")]
        [CategoryAttribute("credential")]
        [BrowsableAttribute(true)]
        [DesignerSerializationVisibilityAttribute(DesignerSerializationVisibility.Visible)]
        public string CiresonPortalPassword
        {
            get
            {
                return ((String)(base.GetValue(RunScript.CiresonPortalPasswordProperty)));
            }
            set
            {
                base.SetValue(RunScript.CiresonPortalPasswordProperty, value);
            }
        }

        //execute the Workflow
        protected override ActivityExecutionStatus Execute(ActivityExecutionContext SMEXCOContext)
        {
            //connect to SCSM and get the SMLets Exchange Connector Settings MP
            EnterpriseManagementGroup emg = new EnterpriseManagementGroup("localhost");
            EnterpriseManagementObject smexcoAdminSettings = emg.EntityObjects.GetObject<EnterpriseManagementObject>(new Guid("a0022e87-75a8-65ee-4581-d923ff06a564"), ObjectQueryOptions.Default);

            //Retrieve the SMExco File Path from the Settings class
            string smexcoFilePath = smexcoAdminSettings[null, "FilePathSMExcoPS"].Value.ToString();

            //load the script into memory
            string smletsExchangeConnector;
            using (var streamReader = new StreamReader(smexcoFilePath, Encoding.UTF8))
            {
                smletsExchangeConnector = streamReader.ReadToEnd();
            }

            // create a Powershell runspace
            Runspace runspace = RunspaceFactory.CreateRunspace();
            runspace.Open();

            // create a pipeline, update variables in PowerShell, then execute it
            Pipeline pipeline = runspace.CreatePipeline();
            pipeline.Commands.AddScript(smletsExchangeConnector);
            pipeline.Runspace.SessionStateProxy.SetVariable("ewsdomain", this.ExchangeDomain);
            pipeline.Runspace.SessionStateProxy.SetVariable("ewsusername", this.ExchangeUsername);
            pipeline.Runspace.SessionStateProxy.SetVariable("ewspassword", this.ExchangePassword);
            pipeline.Runspace.SessionStateProxy.SetVariable("ciresonPortalRunAsUsername", this.CiresonPortalUsername);
            pipeline.Runspace.SessionStateProxy.SetVariable("ciresonPortalRunAsPassword", this.CiresonPortalPassword);
            Collection<PSObject> results = pipeline.Invoke();
            runspace.Close();

            //finish
            return base.Execute(SMEXCOContext);
        }
    }
}
