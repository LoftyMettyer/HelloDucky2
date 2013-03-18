using NServiceBus;
using StructureMap;
using StructureMap.Attributes;

namespace Fusion.Connector.OpenHR
{
	class EndpointConfig : IConfigureThisEndpoint, AsA_Publisher, IWantCustomInitialization
	{
		public void Init()
		{
            NServiceBus.Configure.With()
                .StructureMapBuilder()
                .JsonSerializer()
                .InMemorySubscriptionStorage()
                .UnicastBus()
                .DoNotAutoSubscribe()
                .DisableRavenInstall()
                .DisableTimeoutManager();

        }
	}
}




