using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using Fusion.Messages.SocialCare.Schemas.Patterns;

namespace Fusion.Messages.SocialCare.Schemas
{
    partial class serviceUserChangeData : Patterns.IDatedObject, Patterns.Statuses.IRescindableRecord
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

    partial class serviceUserChangeDataServiceUser : Patterns.IPersonBasicData
    {
        [XmlIgnore]
        public bool roomSpecified
        {
            get;
            set;
        }

        [XmlIgnore]
        public bool locationWithinFacilitySpecified
        {
            get;
            set;
        }

        #region Implementation of IPersonBasicData

        string IPersonBasicData.preferredName
        {
            get { return null; }
            set { } //This property is not supported on service user.
        }

        bool IPersonBasicData.preferredNameSpecified
        {
            get { return false; }
            set { } //This property is not supported on service user.
        }

        #endregion

        string Patterns.IFusionXmlDto.Serialize()
        {
            throw new NotImplementedException();
        }
    }
}
