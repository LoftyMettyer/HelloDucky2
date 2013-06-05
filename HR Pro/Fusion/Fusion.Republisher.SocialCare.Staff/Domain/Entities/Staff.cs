namespace Fusion.Republisher.SocialCare.Domain.Entities
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Fusion.Publisher.SocialCare;

    public class Staff
    {
        public Guid StaffRef { get; set; }
        public string AuditUserName { get; set; }
        public DateTime? EffectiveFrom { get; set; }
        public DateTime? EffectiveTo { get; set; }

        public string XmlData {
            get
            {
                return SerializationHelper.ToXml(Data);
            }
            set
            {
                this.Data = SerializationHelper.FromXml<StaffData>(value);
            }
        }

        public StaffData Data
        {
            get;
            private set;      
        }        
    }
}
