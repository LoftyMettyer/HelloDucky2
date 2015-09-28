using Nexus.Common.Interfaces;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Enums;
using System.Collections.Generic;
using Nexus.Common.Models;

namespace Nexus.Common.Classes
{
    public class ProcessStepEmail : IProcessStep
    {
        public int Id { get; set; }

        public ProcessElementType Type
        {
            get
            {
                return ProcessElementType.Email;
            }
        }

        public ProcessStepStatus Validate()
        {
            return ProcessStepStatus.Success;
        }

        public string BodyTemplate { get; set; }

        public string Subject { get; set; }

    }

}
