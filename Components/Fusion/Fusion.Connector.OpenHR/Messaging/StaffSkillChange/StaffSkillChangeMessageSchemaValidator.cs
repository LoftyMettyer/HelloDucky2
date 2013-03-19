using Fusion.Messages.SocialCare;

namespace Fusion.Connector.OpenHR.Messaging.StaffSkillChange
{
    public class StaffSkillChangeMessageSchemaValidator : SchemaValidatorOutboundFilterHandler<StaffSkillChangeMessage>
    {
        public override bool Handle(StaffSkillChangeMessage message)
        {
            var valid = CheckValidity(message);
            return valid;
        }
    }
}
