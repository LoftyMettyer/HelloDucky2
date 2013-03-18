using NServiceBus;
using System.Configuration;

namespace Fusion.Publisher.SocialCare.Staff
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Publisher, IWantCustomInitialization {
        public void Init()
        {
            NServiceBus.Configure.With()
                .StructureMapBuilder()
                .JsonSerializer()
                .InMemorySubscriptionStorage()
                .UnicastBus()
                .DoNotAutoSubscribe()
                .DisableRavenInstall()
                .DisableSecondLevelRetries()
                .DisableTimeoutManager();
        }

    }
}
