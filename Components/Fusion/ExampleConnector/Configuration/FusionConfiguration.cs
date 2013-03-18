using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Connector1.Configuration
{
    public class FusionConfiguration : IFusionConfiguration
    {
        public FusionConfiguration()
        {
            this.ServiceName = ConfigurationManager.AppSettings["Name"];
            this.Community =  ConfigurationManager.AppSettings["Community"];

        }
        public string ServiceName
        {
            get;
            private set;
        }

        public string Community
        {
            get;
            private set;
        }
    }
}
