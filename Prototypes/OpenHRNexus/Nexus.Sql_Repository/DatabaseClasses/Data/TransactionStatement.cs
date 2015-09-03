using System;

namespace Nexus.Sql_Repository.DatabaseClasses.Data
{
    public class TransactionStatement
    {
        public Guid Id { get; set; }
        public string Statement { get; set; }
        public DateTime Time { get; set; }
        public Guid UserID { get; set; }
    }
}
