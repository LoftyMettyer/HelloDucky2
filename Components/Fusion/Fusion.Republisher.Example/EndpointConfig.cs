using NServiceBus;
using System.Configuration;

namespace Fusion.Publisher.Example
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Publisher, IWantCustomInitialization {
        public void Init()
        {
            NServiceBus.Configure.With()
                .StructureMapBuilder() 
                .JsonSerializer()
                .MsmqSubscriptionStorage()
                .UnicastBus();
        }
    }
}
