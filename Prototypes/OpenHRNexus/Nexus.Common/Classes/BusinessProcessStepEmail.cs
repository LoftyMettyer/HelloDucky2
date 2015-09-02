using Nexus.Common.Interfaces;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Enums;

namespace Nexus.Common.Classes
{
    public class BusinessProcessStepEmail : IBusinessProcessStep
    {
        public int Id { get; set; }

        public BusinessProcessStepType Type
        {
            get
            {
                return BusinessProcessStepType.Email;
            }

        }

        public BusinessProcessStepStatus Validate()
        {
            return BusinessProcessStepStatus.Success;
        }

        public string To { get { return "roberto.caballero@advancedcomputersoftware.com"; } }

        public string Message { get { return "Hello Ducky "; } }

    }
}
