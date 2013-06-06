using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus;
using StructureMap;
using Fusion.Publisher.SocialCare.Staff.Registries;

namespace Fusion.Publisher.SocialCare.Staff
{
    public class StructureMapBootstrap : IWantToRunBeforeConfiguration
    {
        public void Init()
        {
            ObjectFactory.Configure(c =>
            {
                c.AddRegistry<NHibernateRegistry>();
             
            });
        }
    }
}
