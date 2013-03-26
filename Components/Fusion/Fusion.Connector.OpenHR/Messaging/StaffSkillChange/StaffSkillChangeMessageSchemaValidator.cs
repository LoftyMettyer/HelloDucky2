using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffSkillChange
{
    public class StaffSkillChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffSkillChangeRequest>
    {
        public override bool Handle(StaffSkillChangeRequest message)
        {
            var valid = false;

            if (!checkAlreadySent(message))
            {
                valid = CheckValidity(message);
            }

            return valid;
        }
    }
}
