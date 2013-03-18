using NServiceBus;

namespace Fusion.Test.SocialCare
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
        }
    }
}
