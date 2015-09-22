using Nexus.Common.Interfaces;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Enums;

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

//        public string To { get; set; }

        public string BodyTemplate { get; set; }

        public string Subject { get; set; }

        public EmailAddressCollection GetEmailDestinations()
        {
            return new EmailAddressCollection()
            {
                From = "nick.gibson@advancedcomputersoftware.com",
                To = "nick.gibson@advancedcomputersoftware.com"
            };

        }

    }

}
