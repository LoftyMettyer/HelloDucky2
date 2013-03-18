using Fusion.Republisher.Core.MessageProcessors;
using StructureMap.Configuration.DSL;

namespace Fusion.Republisher.Core.Registries
{
    public class FusionMessageProcessorRegistry : Registry
    {
        public FusionMessageProcessorRegistry()
        {
            For<IFusionMessageProcessor>().Use<FusionMessageProcessor>();
        }
    }
}
