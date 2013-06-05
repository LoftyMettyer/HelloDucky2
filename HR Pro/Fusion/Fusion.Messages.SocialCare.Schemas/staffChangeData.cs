using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Fusion.Messages.SocialCare.Schemas
{
    partial class staffChangeData : Patterns.IDatedObject, Patterns.Statuses.IRescindableRecord
    {
        [XmlIgnore]
        public bool recordStatusSpecified
        {
            get { return true; }
            set
            {
                if (!value)
                {
                    throw new InvalidOperationException("recordStatus attribute is required on " + System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name + ".");
                }
            }
        }
    }

    partial class staffChangeDataStaff : Patterns.IPersonBasicData, Patterns.IAddressData
    {
        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool leavingReasonSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool preferredNameSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool homePhoneNumberSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool workMobileSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool personalMobileSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool emailSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool personalEmailSpecified
        {
            get;
            set;
        }

        [System.Xml.Serialization.XmlIgnoreAttribute()]
        public bool managerRefSpecified
        {
            get;
            set;
        }
    }
}
