using Nexus.Common.Enums;
using OpenHRNexus.Common.Enums;

namespace Nexus.Common.Interfaces
{
    public interface IProcessStep
    {
        int Id { get; set; }
        ProcessStepType Type { get; }
        ProcessStepStatus Validate();
    }
}
