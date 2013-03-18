using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace Fusion.Core.Test.Configuration
{
    public class SubscriptionsSection : ConfigurationSection
    {
        #region Constructors
        static SubscriptionsSection()
        {
            //s_propName = new ConfigurationProperty(
            //    "name",
            //    typeof(string),
            //    null,
            //    ConfigurationPropertyOptions.IsRequired
            //    );

            s_propSubscriptions = new ConfigurationProperty(
                "",
                typeof(SubscriptionElementCollection),
                null,
                ConfigurationPropertyOptions.IsRequired | ConfigurationPropertyOptions.IsDefaultCollection
                );

            s_properties = new ConfigurationPropertyCollection();

            //s_properties.Add(s_propName);
            s_properties.Add(s_propSubscriptions);
        }
        #endregion

        #region Fields
        private static ConfigurationPropertyCollection s_properties;
        //private static ConfigurationProperty s_propName;
        private static ConfigurationProperty s_propSubscriptions;
        #endregion

        #region Properties
        //public string Name
        //{
        //    get { return (string)base[s_propName]; }
        //    set { base[s_propName] = value; }
        //}

        public SubscriptionElementCollection Subscriptions
        {
            get { return (SubscriptionElementCollection)base[s_propSubscriptions]; }
        }

        protected override ConfigurationPropertyCollection Properties
        {
            get
            {
                return s_properties;
            }
        }
        #endregion
    }

}
