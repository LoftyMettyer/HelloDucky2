using Nexus.Common.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace Nexus.Sql_Repository.DatabaseClasses.Data
{

    public class ProcessInFlow
    {
        public Guid Id { get; set; }
        public Guid UserId { get; set; }

        public IEnumerable<ProcessInFlowData> Data { get; set; }
    }
}
