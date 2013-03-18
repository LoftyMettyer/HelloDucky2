using System.ComponentModel;
using System.Configuration;

namespace Fusion.Connector.OpenHR
{
	[RunInstaller(true)]
	public partial class Installer : System.Configuration.Install.Installer
	{
		public Installer()
		{
			InitializeComponent();
		}

		public override void Install(System.Collections.IDictionary stateSaver)
		{
			base.Install(stateSaver);

			string targetDirectory = Context.Parameters["targetdir"];
			string param1 = Context.Parameters["Param1"];
			string param2 = Context.Parameters["Param2"];
			string param3 = Context.Parameters["Param3"];

			string exePath = string.Format("{0}Fusion.Connector.OpenHR.DLL", targetDirectory);

			System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(exePath);

			config.AppSettings.Settings["connector_server"].Value = param1;
			config.AppSettings.Settings["connector_db"].Value = param2;
			config.AppSettings.Settings["Param3"].Value = param3;

			config.Save();
		}
	}
}