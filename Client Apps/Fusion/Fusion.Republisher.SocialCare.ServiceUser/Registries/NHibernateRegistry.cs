using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using System.Reflection;

namespace Fusion.Publisher.SocialCare.Staff.Registries
{
    public class NHibernateRegistry : Registry
    {
        public NHibernateRegistry()
        {
            For<INHibernateSession>().Singleton().Use<NHibernateSession>();
        }
    }
}
