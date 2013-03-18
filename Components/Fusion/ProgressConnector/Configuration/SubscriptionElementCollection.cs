using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Text;

namespace ProgressConnector.Configuration
{
    public class SubscriptionElementCollection : ConfigurationElementCollection
    {
        #region Constructor
        public SubscriptionElementCollection()
        {
        }
        #endregion

        #region Properties
        public override ConfigurationElementCollectionType CollectionType
        {
            get
            {
                return ConfigurationElementCollectionType.BasicMap;
            }
        }
        protected override string ElementName
        {
            get
            {
                return "subscribe";
            }
        }

        protected override ConfigurationPropertyCollection Properties
        {
            get
            {
                return new ConfigurationPropertyCollection();
            }
        }
        #endregion

        #region Indexers
        public SubscriptionTypeElement this[int index]
        {
            get
            {
                return (SubscriptionTypeElement)base.BaseGet(index);
            }
            set
            {
                if (base.BaseGet(index) != null)
                {
                    base.BaseRemoveAt(index);
                }
                base.BaseAdd(index, value);
            }
        }

        public SubscriptionTypeElement this[string name]
        {
            get
            {
                return (SubscriptionTypeElement)base.BaseGet(name);
            }
        }
        #endregion

        #region Methods
        public void Add(SubscriptionTypeElement item)
        {
            base.BaseAdd(item);
        }

        public void Remove(SubscriptionTypeElement item)
        {
            base.BaseRemove(item);
        }

        public void RemoveAt(int index)
        {
            base.BaseRemoveAt(index);
        }
        #endregion

        #region Overrides
        protected override ConfigurationElement CreateNewElement()
        {
            return new SubscriptionTypeElement();
        }

        protected override object GetElementKey(ConfigurationElement element)
        {
            return (element as SubscriptionTypeElement).Type;
        }
        #endregion
    }

}
