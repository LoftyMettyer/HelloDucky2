using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Fusion.Republisher.Core.Configuration
{
    public class FusionConfiguration : IFusionConfiguration
    {
        public FusionConfiguration()
        {
            string storeState = ConfigurationManager.AppSettings["StoreState"];
            if (String.IsNullOrEmpty(storeState) || Array.IndexOf(new char[] { 'n', 'f', '0' },  Char.ToLowerInvariant(storeState[0])) >= 0)
            {
                this.StoreState = false;
            } else {
                this.StoreState = true;
            }

            this.Community =  ConfigurationManager.AppSettings["Community"];

        }
        public bool StoreState
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
