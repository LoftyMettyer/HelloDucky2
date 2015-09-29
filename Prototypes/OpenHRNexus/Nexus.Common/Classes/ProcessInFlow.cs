using Nexus.Common.Classes;
using System;
using System.Collections.Generic;

namespace Nexus.Sql_Repository.DatabaseClasses.Data
{

    public class ProcessInFlow
    {
        public Guid Id { get; set; }
        //public Process Process { get; set; }
        public string ProcessName {get; set; }
        public Guid InitiationUserId { get; set; }
        public DateTime? InitiationDateTime { get; set; }
        public DateTime? CompletionDateTime { get; set; }
        public List<ProcessInFlowData> StepData { get; set; } = new List<ProcessInFlowData>();
        public string Caption { get; set; } = "";

    }
}
