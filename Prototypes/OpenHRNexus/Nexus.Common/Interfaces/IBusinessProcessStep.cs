using Nexus.Common.Enums;
using OpenHRNexus.Common.Enums;

namespace Nexus.Common.Interfaces
{
    public interface IBusinessProcessStep
    {
        int Id { get; set; }
        BusinessProcessStepType Type { get; }
        BusinessProcessStepStatus Validate();
    }
}
