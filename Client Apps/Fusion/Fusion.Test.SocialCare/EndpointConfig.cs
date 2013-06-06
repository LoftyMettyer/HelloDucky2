using NServiceBus;

namespace Fusion.Test.SocialCare
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Server, IWantCustomInitialization {

        public void Init()
        {
            NServiceBus.Configure.With()
                .StructureMapBuilder()
                .JsonSerializer()
                .UnicastBus();
        }
    }
}
