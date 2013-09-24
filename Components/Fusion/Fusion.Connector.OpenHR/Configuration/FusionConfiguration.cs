using System.Configuration;

namespace Fusion.Connector.OpenHR.Configuration
{
	public class FusionConfiguration : IFusionConfiguration
	{
		public FusionConfiguration()
		{
			ServiceName = ConfigurationManager.AppSettings["Name"];
			Community = ConfigurationManager.AppSettings["Community"];
			SendAsUser = ConfigurationManager.AppSettings["SendAsUser"];
		}

		public string ServiceName { get; private set; }

		public string Community{ get; private set; }

		public string SendAsUser { get; private set; }
	}
}
