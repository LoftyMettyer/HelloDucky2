using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using Connector1.DatabaseAccess;
using Connector1.Configuration;

namespace ExampleConnector.Registries
{
    public class ConfigurationRegistry : Registry
    {
        public ConfigurationRegistry()
        {
            For<IFusionConfiguration>().Use<FusionConfiguration>();
        }
    }
}
