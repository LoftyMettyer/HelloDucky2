using Fusion.Core.Logging;
using StructureMap.Configuration.DSL;

namespace Fusion.Republisher.Core.Registries
{
    public class FusionLoggerRegistry : Registry
    {
        public FusionLoggerRegistry()
        {
            For<IFusionLogService>().Use<FusionLogger>().Ctor<string>("source").Is("republisher");
        }
    }
}
