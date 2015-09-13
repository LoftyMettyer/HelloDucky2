using Nexus.Common.Enums;
using OpenHRNexus.Common.Enums;

namespace Nexus.Common.Interfaces
{
    public interface IProcessStep
    {
        int Id { get; set; }
        ProcessElementType Type { get; }
        ProcessStepStatus Validate();
    }
}
