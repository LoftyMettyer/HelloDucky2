using Nexus.Common.Enums;

namespace Nexus.Common.Classes
{
    public class BusinessProcessStepResponse
    {
        public BusinessProcessStepStatus Status { get; set; }
        public string Message { get; set; }
        public string FollowOnUrl { get; set; }
    }
}
