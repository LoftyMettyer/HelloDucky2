using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using Fusion.Connector.OpenHR.DatabaseAccess;
using Fusion.Connector.OpenHR.Configuration;

namespace Fusion.Connector.OpenHR.Registries
{
    public class ConfigurationRegistry : Registry
    {
        public ConfigurationRegistry()
        {
            For<IFusionConfiguration>().Use<FusionConfiguration>();
        }
    }
}
