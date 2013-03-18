using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using Fusion.Test.SocialCare;
using Fusion.Core.Test;

namespace Fusion.Test.Registries
{
    public class ConfigurationRegistry : Registry
    {
        public ConfigurationRegistry()
        {
            For<ITestingConfiguration>().Use<TestingConfiguration>();
        }
    }
}
