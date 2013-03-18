using Fusion.Republisher.Core.Registries;
using NServiceBus;
using StructureMap;

namespace Fusion.Publisher.SocialCare
{
    public class StructureMapBootstrap : IWantToRunBeforeConfiguration
    {
        public void Init()
        {
            ObjectFactory.Configure(c =>
            {
                c.AddRegistry<FusionLoggerRegistry>();
                c.AddRegistry<DatabaseAccessRegistry>();
                c.AddRegistry<MessageStateSerializerRegistry>();
                c.AddRegistry<FusionMessageProcessorRegistry>();
                c.AddRegistry<ConfigurationRegistry>();
            });
        }
    }
}
