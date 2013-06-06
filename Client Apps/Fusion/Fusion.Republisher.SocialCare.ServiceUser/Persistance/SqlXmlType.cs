namespace Prototype.NHibernateTypeSerialization.Persistance
{
    using System.Data;
    using NHibernate.SqlTypes;

    public class SqlXmlType : SqlType
    {
        public SqlXmlType() : base(DbType.Xml)
        {
        }
    }
}
