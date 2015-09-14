using Nexus.Common.Classes;
using Nexus.Common.Models;
using System;

namespace Nexus.Sql_Repository.DatabaseClasses.Data
{

    public class ProcessInFlow
    {
        public Guid Id { get; set; }
        public Guid UserId { get; set; }
        public Process Process { get; set; }
        public int WebFormId { get; set; }
    }
}
