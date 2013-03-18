

namespace ExampleConnector
{
    using ExampleConnector.Registries;
    using NServiceBus;
    using StructureMap;

    public class StructureMapBootstrap : IWantToRunBeforeConfiguration
    {
        private static bool IsInitialised = false;
        private static object lockObject = new object();

        public void Init()
        {
            lock (lockObject)
            {
                if (!IsInitialised)
                {
                    IsInitialised = true;

                    ObjectFactory.Configure(c =>
                    {
                        c.AddRegistry<DatabaseAccessRegistry>();
                        c.AddRegistry<OutboundBuilderRegistry>();
                        c.AddRegistry<DatabaseAccessRegistry>();
                        c.AddRegistry<MessageSenderRegistry>();
                        c.AddRegistry<ConfigurationRegistry>();
                        c.AddRegistry<SendFusionMessageRequestBuilderRegistry>();

                        c.AddRegistry<FilterRegistry>();
                        c.AddRegistry<FusionLoggerRegistry>();
                    });
                }
            }

  
        }
    }
}
