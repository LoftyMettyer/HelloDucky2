namespace Prototype.NHibernateTypeSerialization.Domain.Entities
{
    using System;
    using Prototype.NHibernateTypeSerialization.Helpers;

    [Serializable]
    public class StaffData : ICloneable
    {
        public virtual string Title { get; set; }
        public virtual string Forenames { get; set; }
        public virtual string Surname { get; set; }
        public virtual string PayrollNumber { get; set; }
        public virtual DateTime DOB { get; set; }

        public virtual string EmployeeType { get; set; }
        public virtual string EmploymentStatus { get; set; }

        public virtual string HomePhoneNumber { get; set; }
        public virtual string WorkMobile { get; set; }
        public virtual string PersonalMobile { get; set; }

        public virtual string Email { get; set; }
        public virtual string PersonalEmail { get; set; }

        public virtual string AddressLine1 { get; set; }
        public virtual string AddressLine2 { get; set; }
        public virtual string AddressLine3 { get; set; }
        public virtual string AddressLine4 { get; set; }
        public virtual string AddressLine5 { get; set; }

        public virtual string Postcode { get; set; }
        public virtual string Gender { get; set; }

        public virtual DateTime StartDate { get; set; }
        public virtual DateTime? LeavingDate { get; set; }

        public virtual string LeavingReason { get; set; }
        public virtual string CompanyName { get; set; }

        public virtual string JobTitle { get; set; }
        public virtual Guid? ManagerRef { get; set; }

        public object Clone()
        {
            return CloneHelper.Clone<StaffData>(this);
        }
    }
}
