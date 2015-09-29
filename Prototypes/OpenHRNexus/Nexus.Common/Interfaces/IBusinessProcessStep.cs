using Nexus.Common.Enums;
using OpenHRNexus.Common.Enums;
using System.Collections.Generic;

namespace Nexus.Common.Interfaces
{
    public interface IProcessStep
    {
        int Id { get; set; }
        ProcessElementType Type { get; }
        ProcessStepStatus Validate();
        Dictionary<string, object> Variables { get; set; }
    }
}
