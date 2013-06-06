using NServiceBus;
using System.Configuration;

namespace Fusion.Publisher.SocialCare.StaffPublisher
{
    class EndpointConfig : IConfigureThisEndpoint, AsA_Publisher, IWantCustomInitialization {
        public void Init()
        {
            NServiceBus.Configure.With()
                .StructureMapBuilder() 
                .JsonSerializer()
                .UnicastBus()
                .DoNotAutoSubscribe();
        }

    }
}
