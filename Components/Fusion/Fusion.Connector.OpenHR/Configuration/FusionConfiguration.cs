using System.Configuration;

namespace Fusion.Connector.OpenHR.Configuration
{
    public class FusionConfiguration : IFusionConfiguration
    {
        public FusionConfiguration()
        {
            ServiceName = ConfigurationManager.AppSettings["Name"];
            Community = ConfigurationManager.AppSettings["Community"];

        }
        public string ServiceName
        {
            get;
            private set;
        }

        public string Community
        {
            get;
            private set;
        }
    }
}
