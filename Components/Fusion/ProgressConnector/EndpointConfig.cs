using NServiceBus;
using StructureMap;
using Connector1.Registries;
using ProgressConnector.Registries;

namespace Fusion.ProgressConnector
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Publisher, IWantCustomInitialization {
        public void Init()
        {
            NServiceBus.Configure.With()

               .StructureMapBuilder()
                .JsonSerializer()
                .UnicastBus()
                .DoNotAutoSubscribe()
                .DisableRavenInstall()
                .DisableSecondLevelRetries()
                .DisableTimeoutManager();


            ObjectFactory.Configure(c =>
            {
                c.AddRegistry<BusTypeBuilderRegistry>();
                //c.AddRegistry<ConfigurationRegistry>();
                c.AddRegistry<ProgressInterfaceRegistry>();
            });

        }

    }
}
