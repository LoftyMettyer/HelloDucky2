using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using StructureMap.Configuration.DSL;
using ProgressConnector.Configuration;

namespace ProgressConnector.Registries
{
    public class ConfigurationRegistry : Registry
    {
        public ConfigurationRegistry()
        {
            //For<IFusionConfiguration>().Use<FusionConfiguration>();
        }
    }
}
