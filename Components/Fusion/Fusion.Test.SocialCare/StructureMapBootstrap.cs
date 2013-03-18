
namespace Fusion.Test.SocialCare
{
    using Fusion.Test.Registries;
    using Fusion.Test.SocialCare.Registries;
    using NServiceBus;
    using StructureMap;

    public class StructureMapBootstrap : IWantToRunBeforeConfiguration
    {
        public void Init()
        {
            ObjectFactory.Configure(c =>
            {
                c.AddRegistry<MessageSenderRegistry>();
                c.AddRegistry<FusionXmlMetadataExtractorRegistry>();
                c.AddRegistry<OutboundMessageWatcherRegistry>();
                c.AddRegistry<ConfigurationRegistry>();
            });
        }
    }
}
