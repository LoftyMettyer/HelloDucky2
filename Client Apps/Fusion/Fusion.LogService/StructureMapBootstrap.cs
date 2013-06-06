

namespace Fusion.LogService
{
    using Fusion.LogService.Registries;
    using NServiceBus;
    using StructureMap;

    public class StructureMapBootstrap : IWantToRunBeforeConfiguration
    {
        public void Init()
        {
            ObjectFactory.Configure(c =>
            {
                c.AddRegistry<DatabaseAccessRegistry>();
            });
        }
    }
}
