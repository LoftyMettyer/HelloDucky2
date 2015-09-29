using System;

namespace Nexus.Common.Classes
{
    public class ProcessInFlowData
    {
        public Guid Id { get; set; }
        public Guid UserId { get; set; }
        public DateTime StepDateTime { get; set; }
        public string StepData { get; set; }
    }
}
