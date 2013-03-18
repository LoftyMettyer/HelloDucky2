using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;

namespace ProgressConnector.Configuration
{
    public class SubscriptionTypeElement : ConfigurationElement
    {
        #region Constructors
        static SubscriptionTypeElement()
        {
            s_typeName = new ConfigurationProperty(
                "type",
                typeof(string),
                null,
                ConfigurationPropertyOptions.IsRequired
                );

            s_properties = new ConfigurationPropertyCollection();

            s_properties.Add(s_typeName);
        }
        #endregion

        #region Fields
        private static ConfigurationPropertyCollection s_properties;
        private static ConfigurationProperty s_typeName;
        private static ConfigurationProperty s_propDesc;
        private static ConfigurationProperty s_propState;
        private static ConfigurationProperty s_propSequence;
        #endregion

        #region Properties
        public string Type
        {
            get { return (string)base[s_typeName]; }
            set { base[s_typeName] = value; }
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
