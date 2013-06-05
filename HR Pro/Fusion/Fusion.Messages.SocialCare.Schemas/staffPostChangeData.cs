using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Fusion.Messages.SocialCare.Schemas
{
    partial class staffPostChangeData : Patterns.IDatedObject, Patterns.Statuses.IRescindableRecord
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

    partial class staffPostChangeDataStaffPost
    {
        [XmlIgnore]
        public bool siteManagerRefSpecified
        {
            get;
            set;
        }

        [XmlIgnore]
        public bool departmentSpecified
        {
            get;
            set;
        }
    }
}
