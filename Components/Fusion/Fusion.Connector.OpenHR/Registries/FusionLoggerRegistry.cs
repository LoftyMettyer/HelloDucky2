using Fusion.Connector.OpenHR.Configuration;
using StructureMap.Configuration.DSL;
using Fusion.Core.Logging;

namespace Fusion.Connector.OpenHR.Registries
{
    public class FusionLoggerRegistry : Registry
    {
        public FusionLoggerRegistry()
        {
            For<IFusionLogService>().Use<FusionLogger>().Ctor<string>("source").Is("Fusion.Connector.OpenHR"); 
        }
    }
}
