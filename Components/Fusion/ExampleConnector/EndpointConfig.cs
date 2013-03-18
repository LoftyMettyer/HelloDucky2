using NServiceBus;
using StructureMap;
using StructureMap.Attributes;
using Connector1.Configuration;

namespace ExampleConnector
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
                       .DisableTimeoutManager();

        }

    
    }
}
