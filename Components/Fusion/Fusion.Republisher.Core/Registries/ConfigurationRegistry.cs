using Fusion.Republisher.Core.Configuration;
using StructureMap.Configuration.DSL;

namespace Fusion.Republisher.Core.Registries
{
    public class ConfigurationRegistry : Registry
    {
        public ConfigurationRegistry()
        {
            For<IFusionConfiguration>().Use<FusionConfiguration>();
        }
    }
}
