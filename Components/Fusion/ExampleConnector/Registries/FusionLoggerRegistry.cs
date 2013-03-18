using Connector1.Configuration;
using StructureMap.Configuration.DSL;
using Fusion.Core.Logging;

namespace ExampleConnector.Registries
{
    public class FusionLoggerRegistry : Registry
    {
        public FusionLoggerRegistry()
        {
            For<IFusionLogService>().Use<FusionLogger>().Ctor<string>("source").Is("exampleConnector"); 
        }
    }
}
