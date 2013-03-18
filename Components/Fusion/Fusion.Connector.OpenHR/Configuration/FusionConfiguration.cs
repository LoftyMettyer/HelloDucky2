using System.Configuration;

namespace Fusion.Connector.OpenHR.Configuration
{
    public class FusionConfiguration : IFusionConfiguration
    {
        public FusionConfiguration()
        {
            ServiceName = string.Format("{0}.{1}" ,ConfigurationManager.AppSettings["Name"],ConfigurationManager.AppSettings["OpenHR_db"]);
            InputQueue = ConfigurationManager.AppSettings["InputQueue"];

        }
        public string ServiceName
        {
            get;
            private set;
        }

        public string InputQueue
        {
            get;
            private set;
        }
    }
}
