namespace Prototype.NHibernateTypeSerialization.Domain.Maps
{
    using System;
    using FluentNHibernate.Mapping;
    using NHibernate;
    using Prototype.NHibernateTypeSerialization.Domain.Entities;
    using Prototype.NHibernateTypeSerialization.Persistance.CustomTypes;

    public class StaffMap : ClassMap<Staff>
    {
        public StaffMap()
        {
            Id(x => x.StaffRef);
            Map(x => x.AuditUserName);
            Map(x => x.EffectiveFrom);
            Map(x => x.EffectiveTo);
            Map(x => x.Data).CustomType<StaffDataType>().Not.Nullable();
        }
    }
}
