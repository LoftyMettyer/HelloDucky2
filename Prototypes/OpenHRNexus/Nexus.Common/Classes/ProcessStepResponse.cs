using Nexus.Common.Enums;

namespace Nexus.Common.Classes
{
    public class ProcessStepResponse
    {
        public ProcessStepStatus Status { get; set; }
        public string Message { get; set; }
        public string FollowOnUrl { get; set; }
    }
}
