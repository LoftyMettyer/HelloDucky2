namespace Prototype.NHibernateTypeSerialization.Persistance.Conventions
{
    using System;
    using FluentNHibernate.Conventions;
    using FluentNHibernate.Conventions.Instances;
    using Prototype.NHibernateTypeSerialization.Persistance.CustomTypes;

    public class StaffDataTypeConvention : UserTypeConvention<StaffDataType>
    {
        public override void Apply(IPropertyInstance instance)
        {
            base.Apply(instance);
        }
    }
}
