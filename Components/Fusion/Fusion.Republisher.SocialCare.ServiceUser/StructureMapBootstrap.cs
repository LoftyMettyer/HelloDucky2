using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus;
using StructureMap;
using Fusion.Republisher.Core.Registries;

namespace Fusion.Publisher.SocialCare.Staff
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
