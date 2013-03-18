using NServiceBus;
using StructureMap;

namespace Fusion.LogService
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Server, IWantCustomInitialization {

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
