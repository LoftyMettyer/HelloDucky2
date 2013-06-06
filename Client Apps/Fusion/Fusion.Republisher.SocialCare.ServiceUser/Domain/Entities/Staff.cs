namespace Prototype.NHibernateTypeSerialization.Domain.Entities
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;

    public class Staff
    {
        public virtual Guid StaffRef { get; set; }
        public virtual string AuditUserName { get; set; }
        public virtual DateTime? EffectiveFrom { get; set; }
        public virtual DateTime? EffectiveTo { get; set; }

        public virtual StaffData Data { get; set; }

        public override string ToString()
        {
            return "";
            //return String.Format("[{0}] {1} {2} ({3})\n{4}\n{5}\n{6}\n{7}\n{8}\n{9}\nNT:{10}\n", Id, FirstName, LastName, BirthDate, HomeAddress.AddressLine1, HomeAddress.AddressLine2, HomeAddress.Town, HomeAddress.County, HomeAddress.Postcode, HomeAddress.Country, HomeAddress.NewThing ?? "null");
        }
    }
}
